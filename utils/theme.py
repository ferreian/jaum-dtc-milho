"""
utils/theme.py
CSS e helpers de tema compartilhados entre todas as páginas do Painel de Milho.
(Visual idêntico ao painel de soja — verde Stine, header com logo, sidebar estilizada.)
"""

import base64
from pathlib import Path

import streamlit as st

BASE_DIR = Path(__file__).parent.parent


def aplicar_tema():
    st.markdown("""
<style>
:root {
    --green-primary:  #27AE60;
    --green-dark:     #1E8449;
    --green-light:    #E9F7EF;
    --green-mid:      #A9DFBF;
    --text-primary:   #1A1A1A;
    --text-secondary: #6B7280;
    --bg-main:        #F8FAF9;
    --bg-white:       #FFFFFF;
    --border:         #E5E7EB;
    --shadow:         0 1px 4px rgba(0,0,0,0.07);
}

html, body {
    font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important;
    background-color: var(--bg-main) !important;
    color: var(--text-primary) !important;
}
/* Aplicar fonte globalmente sem sobrescrever componentes externos como AgGrid */
[class*="css"]:not([class*="ag-"]) {
    font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important;
}

/* ── Sidebar limpa ── */
[data-testid="stSidebar"] {
    background-color: var(--bg-white) !important;
    border-right: 1px solid var(--border) !important;
}
[data-testid="stSidebar"] * { color: var(--text-primary) !important; }

/* ── Nav pages ── */
[data-testid="stSidebarNav"] {
    padding-top: 1.5rem !important;
}
[data-testid="stSidebarNav"] a {
    padding: 9px 16px !important;
    border-radius: 8px !important;
    font-size: 14px !important;
    font-weight: 500 !important;
    transition: background 0.15s !important;
    color: var(--text-primary) !important;
    text-decoration: none !important;
    display: flex !important;
    align-items: center !important;
}
[data-testid="stSidebarNav"] a:hover {
    background: var(--green-light) !important;
    color: var(--green-dark) !important;
}
[data-testid="stSidebarNav"] a[aria-current="page"] {
    background: var(--green-light) !important;
    color: var(--green-dark) !important;
    font-weight: 600 !important;
}

/* ── Oculta label padrão "Navigation" do pages ── */
[data-testid="stSidebarNav"]::before { display: none !important; }

/* ── Topo da página — header com logo ── */
.jaum-header {
    display: flex;
    align-items: center;
    gap: 1.2rem;
    padding: 0.5rem 0 1.2rem;
    border-bottom: 1px solid var(--border);
    margin-bottom: 1.5rem;
}
.jaum-header img {
    height: 150px;
    width: auto;
}
.jaum-header-divider {
    width: 1px;
    height: 32px;
    background: var(--border);
}
.jaum-header-text h1 {
    font-size: 2.4rem !important;
    font-weight: 700 !important;
    margin: 0 !important;
    line-height: 1.2 !important;
}
.jaum-header-text p {
    font-size: 15px !important;
    color: var(--text-secondary) !important;
    margin: 2px 0 0 !important;
}

/* ── Conteúdo principal ── */
.main .block-container {
    padding: 1.5rem 2.5rem 2rem !important;
    max-width: 1400px !important;
}

/* ── Títulos ── */
h1 { font-size: 1.9rem !important; font-weight: 700 !important; }
h2 { font-size: 1.5rem !important; font-weight: 600 !important; }
h3 { font-size: 1.2rem !important; font-weight: 600 !important; }

/* ── Botão primário ── */
.stButton > button[kind="primary"] {
    background-color: var(--green-primary) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 14px !important;
    padding: 0.5rem 1.2rem !important;
    transition: background 0.2s !important;
}
.stButton > button[kind="primary"]:hover { background-color: var(--green-dark) !important; }

/* ── Botão secundário ── */
.stButton > button {
    border-radius: 8px !important;
    font-weight: 500 !important;
    font-size: 14px !important;
    border: 1px solid var(--border) !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    border-color: var(--green-primary) !important;
    color: var(--green-primary) !important;
}

/* ── Métricas ── */
[data-testid="stMetric"] {
    background: var(--bg-white) !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    padding: 1rem 1.2rem !important;
    box-shadow: var(--shadow) !important;
}
[data-testid="stMetricLabel"] { color: var(--text-secondary) !important; font-size: 13px !important; }
[data-testid="stMetricValue"] { color: var(--text-primary) !important; font-size: 1.6rem !important; font-weight: 700 !important; }

/* ── Alerts ── */
.stAlert { border-radius: 8px !important; font-size: 14px !important; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
}

/* ── Expander ── */
[data-testid="stExpander"] {
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    background: var(--bg-white) !important;
}

hr { border-color: var(--border) !important; }

/* ── AgGrid — cabeçalho escuro e ícones brancos ── */
.ag-theme-alpine .ag-header,
.ag-theme-alpine .ag-header-row,
.ag-theme-alpine .ag-header-cell,
.ag-theme-alpine .ag-header-cell-resize {
    background-color: #4A4A4A !important;
    background: #4A4A4A !important;
}
.ag-theme-alpine .ag-header-cell-text,
.ag-theme-alpine .ag-header-cell-label,
.ag-theme-alpine .ag-header-group-cell-label {
    color: #FFFFFF !important;
    font-weight: 700 !important;
    font-size: 13px !important;
}
.ag-theme-alpine .ag-icon,
.ag-theme-alpine .ag-header-icon,
.ag-theme-alpine .ag-icon-menu,
.ag-theme-alpine .ag-icon-filter,
.ag-theme-alpine .ag-icon-columns,
.ag-theme-alpine .ag-header-cell-menu-button,
.ag-theme-alpine .ag-header-cell-menu-button span {
    color: #FFFFFF !important;
    opacity: 1 !important;
    fill: #FFFFFF !important;
}
[data-testid="stCaptionContainer"] { color: #374151 !important; font-size: 13px !important; }
.stSpinner > div { border-top-color: var(--green-primary) !important; }

::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: var(--bg-main); }
::-webkit-scrollbar-thumb { background: var(--green-mid); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--green-primary); }
</style>
""", unsafe_allow_html=True)


def logo_base64():
    for ext in ["png", "svg", "jpg", "jpeg"]:
        path = BASE_DIR / "assets" / f"logo.{ext}"
        if path.exists():
            mime = "image/svg+xml" if ext == "svg" else f"image/{ext}"
            data = base64.b64encode(path.read_bytes()).decode()
            return f"data:{mime};base64,{data}"
    return None


def page_header(titulo: str, subtitulo: str = "", imagem: str = ""):
    """
    Renderiza o header padrão de cada página:
    [Logo ou imagem customizada] | [Título da página / subtítulo]
    """
    if imagem:
        path = BASE_DIR / "assets" / imagem
        if path.exists():
            ext = imagem.rsplit(".", 1)[-1].lower()
            mime = "image/svg+xml" if ext == "svg" else f"image/{ext}"
            data = base64.b64encode(path.read_bytes()).decode()
            src = f"data:{mime};base64,{data}"
        else:
            src = logo_base64()
    else:
        src = logo_base64()
    logo_html = (
        f'<img src="{src}" />'
        if src
        else '<span style="font-size:1.4rem;font-weight:700;color:#27AE60;">🌽</span>'
    )
    sub_html = f'<p>{subtitulo}</p>' if subtitulo else ""

    st.markdown(f"""
<div class="jaum-header">
    {logo_html}
    <div class="jaum-header-divider"></div>
    <div class="jaum-header-text">
        <h1>{titulo}</h1>
        {sub_html}
    </div>
</div>
""", unsafe_allow_html=True)


# Mantém compatibilidade com chamadas antigas
def sidebar_header():
    pass


def secao_titulo(label: str, titulo: str, descricao: str = ""):
    """Bloco editorial: label em caps + título grande + descrição."""
    desc_html = (
        f'<p style="font-size:16px;color:#6B7280;margin:6px 0 0;line-height:1.6;">{descricao}</p>'
        if descricao else ""
    )
    st.markdown(f"""
<div style="margin: 1.5rem 0 1rem;">
    <p style="font-size:13px;font-weight:600;color:#6B7280;text-transform:uppercase;
              letter-spacing:0.05em;margin:0 0 6px;">{label}</p>
    <h2 style="font-size:2.0rem;font-weight:700;color:#1A1A1A;margin:0;line-height:1.2;">{titulo}</h2>
    {desc_html}
</div>
""", unsafe_allow_html=True)


def rodape(altura_logo: int = 20):
    """Rodapé padrão: logo Stine + autoria, com link para o LinkedIn.
    Reaproveita a mesma logo do header (assets/logo.*)."""
    src = logo_base64()
    logo_html = (
        f'<img src="{src}" alt="Stine" '
        f'style="height:{altura_logo}px;width:auto;vertical-align:middle;margin-right:10px;">'
        if src else ""
    )
    st.markdown(f"""
<div style="margin-top:2.2rem;padding-top:1rem;border-top:1px solid #E5E7EB;text-align:center;">
  <a href="https://www.linkedin.com/in/eng-agro-andre-ferreira/" target="_blank" rel="noopener"
     style="text-decoration:none;color:#374151;font-size:13px;
            font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;
            display:inline-flex;align-items:center;justify-content:center;">
    {logo_html}<span><strong>Andre Ferreira</strong> &middot; Especialista em Dados &middot; STINE</span>
  </a>
</div>
""", unsafe_allow_html=True)
