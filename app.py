import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pathlib import Path
import io
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

st.set_page_config(page_title="REDEFRETE", layout="wide", initial_sidebar_state="expanded")

BASE_DIR = Path(__file__).parent
EXCEL = BASE_DIR / "Base de Prestadores de Serviço.xlsx"
LOGO = BASE_DIR / "logo2.png"

st.markdown("""
<style>
.stApp{background:#FFFFFF}
.block-container{padding-top:0.8rem}
[data-testid="stMetric"]{background:#FFFFFF;border:1px solid #E5E7EB;border-left:4px solid #E10600;border-radius:12px;padding:14px}
section[data-testid="stSidebar"]{background:#F9FAFB;border-right:1px solid #E5E7EB}
</style>
""", unsafe_allow_html=True)

# --- FUNÇÃO QUE CARREGA DO SHAREPOINT OU LOCAL ---
@st.cache_data(ttl=60)
def load_local():
    df = pd.read_excel(EXCEL).fillna("")
    df.columns = [c.strip() for c in df.columns]
    return df

def load_sharepoint(email, senha):
    site_url = "https://vlooz.sharepoint.com/sites/secretaria"
    file_path = "/sites/secretaria/Documentos Compartilhados/Base de Prestadores de Serviço.xlsx"
    ctx = ClientContext(site_url).with_credentials(UserCredential(email, senha))
    file_content = ctx.web.get_file_by_server_relative_url(file_path).read()
    df = pd.read_excel(io.BytesIO(file_content)).fillna("")
    df.columns = [c.strip() for c in df.columns]
    return df

# --- SIDEBAR LOGIN ---
with st.sidebar:
    st.markdown("## 🔐 Atualização SharePoint")
    st.caption("Se deixar em branco, usa o Excel do GitHub")
    email = st.text_input("E-mail Redefrete")
    senha = st.text_input("Senha", type="password")
    btn = st.button("Carregar do SharePoint", use_container_width=True, type="primary")
    
    st.divider()
    st.markdown("## 🔍 Filtros")
    # os filtros vão depois

if btn and email and senha:
    try:
        with st.spinner("Baixando do SharePoint..."):
            df = load_sharepoint(email, senha)
        st.sidebar.success("Dados do SharePoint!")
    except Exception as e:
        st.sidebar.error(f"Erro: {e}")
        st.sidebar.info("Se tem 2 fatores ativado, crie uma senha de app em account.microsoft.com > Segurança")
        df = load_local()
else:
    df = load_local()

C_NOME = "Nome da empresa / Prestador"
C_SERV = "Serviço Prestado"
C_CID = "Cidade"
C_FDS = "Atende FINAL DE SEMANA? (Sim/Não)"
C_CONT = "Contato (telefone / e-mail)"
C_OBS = "Observações"

# HEADER
col_logo, col_title = st.columns([1, 6])
with col_logo:
    if LOGO.exists():
        st.image(str(LOGO), width=180)
with col_title:
    st.markdown("<h1 style='color:#111827;margin:0'>BASE DE PRESTADORES</h1><p style='color:#E10600;margin:0;font-weight:800;letter-spacing:2px'>REDEFRETE</p>", unsafe_allow_html=True)

# FILTROS (continua na sidebar)
with st.sidebar:
    busca = st.text_input("Buscar", placeholder="Nome, cidade, serviço...")
    f_cid = st.multiselect("Cidade", sorted(df[C_CID].dropna().unique()), placeholder="Escolha Cidade")
    f_serv = st.multiselect("Serviço", sorted(df[C_SERV].dropna().unique()), placeholder="Escolha Serviço")
    f_fds = st.selectbox("FDS", ["Todos","Sim","Não"])
    if st.button("Limpar filtros", use_container_width=True):
        st.rerun()

df_f = df.copy()
if busca:
    df_f = df_f[df_f.apply(lambda r: busca.lower() in " ".join(map(str,r.values)).lower(), axis=1)]
if f_cid:
    df_f = df_f[df_f[C_CID].isin(f_cid)]
if f_serv:
    df_f = df_f[df_f[C_SERV].isin(f_serv)]
if f_fds != "Todos":
    df_f = df_f[df_f[C_FDS].astype(str).str.lower() == f_fds.lower()]

sim = len(df_f[df_f[C_FDS].astype(str).str.lower()=="sim"])
nao = len(df_f) - sim

k1,k2,k3,k4,k5 = st.columns(5)
k1.metric("Total Fornecedores", len(df_f))
k2.metric("Atende FDS ✅", sim)
k3.metric("Não atende ❌", nao)
k4.metric("Cidades", df_f[C_CID].nunique())
k5.metric("Serviços", df_f[C_SERV].nunique())

c1,c2,c3 = st.columns(3)
with c1:
    top = df_f[C_CID].value_counts().head(6)
    fig = go.Figure(go.Bar(y=top.index, x=top.values, orientation='h', marker_color='#E10600'))
    fig.update_layout(template="plotly_white", paper_bgcolor="white", plot_bgcolor="white", height=320, title="Top Cidades", margin=dict(l=0,r=20,t=40,b=0))
    st.plotly_chart(fig, use_container_width=True)
with c2:
    fig2 = go.Figure(go.Pie(labels=["Sim","Não"], values=[sim,nao], hole=0.65, marker_colors=["#E10600","#E5E7EB"], textinfo="label+percent"))
    fig2.update_layout(template="plotly_white", paper_bgcolor="white", height=320, title="Final de Semana")
    st.plotly_chart(fig2, use_container_width=True)
with c3:
    top2 = df_f[C_SERV].value_counts().head(6)
    fig3 = go.Figure(go.Pie(labels=top2.index, values=top2.values, hole=0.6))
    fig3.update_layout(template="plotly_white", paper_bgcolor="white", height=320, title="Por Serviço")
    st.plotly_chart(fig3, use_container_width=True)

st.divider()
st.subheader("Base Detalhada")
st.dataframe(df_f[[C_NOME, C_SERV, C_CID, C_FDS, C_CONT, C_OBS]], use_container_width=True, height=700, hide_index=True)
