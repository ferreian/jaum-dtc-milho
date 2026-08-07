"""
app.py — Painel Analítico JAUM DTC (Milho) — Home
"""
from pathlib import Path
import streamlit as st
from utils.theme import aplicar_tema, page_header, rodape

st.set_page_config(
    page_title="Painel JAUM DTC · Milho",
    page_icon="🌽",
    layout="wide",
    initial_sidebar_state="expanded",
)

aplicar_tema()
st.markdown("<style>.jaum-header img { height: 60px !important; }</style>", unsafe_allow_html=True)
page_header("Painel Analítico de Híbridos de Milho", "JAUM DTC · Stine Seed")

# ── Assinatura em destaque (logo abaixo do título) ───────────────────────────
st.markdown("""
<div style="margin:-0.8rem 0 1.4rem;">
  <a href="https://www.linkedin.com/in/eng-agro-andre-ferreira/" target="_blank" rel="noopener"
     style="display:inline-flex;align-items:center;gap:9px;text-decoration:none;
            background:#E9F7EF;border:1px solid #A9DFBF;border-radius:24px;
            padding:7px 16px 7px 13px;transition:all .2s;">
    <svg width="16" height="16" viewBox="0 0 24 24" fill="#0A66C2" style="flex-shrink:0;">
      <path d="M20.45 20.45h-3.56v-5.57c0-1.33-.02-3.04-1.85-3.04-1.85 0-2.14 1.45-2.14 2.94v5.67H9.35V9h3.41v1.56h.05c.48-.9 1.63-1.85 3.37-1.85 3.6 0 4.27 2.37 4.27 5.46v6.28zM5.34 7.43a2.06 2.06 0 1 1 0-4.13 2.06 2.06 0 0 1 0 4.13zm1.78 13.02H3.55V9h3.57v11.45zM22.22 0H1.77C.79 0 0 .77 0 1.72v20.56C0 23.23.79 24 1.77 24h20.45c.98 0 1.78-.77 1.78-1.72V1.72C24 .77 23.2 0 22.22 0z"/>
    </svg>
    <span style="font-size:14px;color:#374151;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;">
      Desenvolvido por <strong style="color:#1E8449;">Andre Ferreira</strong>
      &nbsp;·&nbsp; Especialista em Dados &nbsp;·&nbsp; STINE
    </span>
  </a>
</div>
""", unsafe_allow_html=True)

BASE_DIR = Path(__file__).parent

# ── Layout: intro + imagem | cards ───────────────────────────────────────────
col_esq, col_dir = st.columns([2, 3], gap="large")

with col_esq:
    st.markdown("""
<div style="margin-top: 1rem;">
    <p style="font-size:15px; color:#1A1A1A; line-height:1.8;">
        Painel multissafra de análise de cultivares de milho do programa
        <strong>JAUM DTC</strong> — produtividade, sanidade, caracterização
        agronômica e efeito de densidade de plantio.
    </p>
    <p style="font-size:13px; color:#6B7280; line-height:1.6; margin-top: 0.5rem;">
        &#127463;&#127479; Departamento Técnico de Culturas · Stine Brasil
    </p>
    <p style="font-size:13px; color:#6B7280; line-height:1.6; margin-top: 0.2rem;">
        Safras 2024/25 e 2025/26 · milho safrinha
    </p>
    <p style="font-size:14px; color:#374151; line-height:1.8; margin-top: 0.8rem;">
        Comece pela <strong>Auditoria</strong> ou pelo <strong>Diagnóstico</strong> para
        conferir o cálculo e o status dos dados antes de usar as análises.
    </p>
</div>
""", unsafe_allow_html=True)
    img_path = BASE_DIR / "assets" / "App development-amico.png"
    if img_path.exists():
        st.image(str(img_path), use_container_width=True)

# ── Cards de navegação (páginas reais do painel de milho) ─────────────────────
PAGINAS = [
    {
        "icone": "🧮",
        "titulo": "Auditoria de Cálculo",
        "subtitulo": "Confira o gold plot a plot",
        "descricao": "Escolha fazenda e material, veja as variáveis cruas e refaça a conta a partir "
                     "delas — a página compara com o gold e mostra se bate. Tem simulador de input em 2025.",
        "tags": ["Conferência", "Réguas", "Rastreável"],
    },
    {
        "icone": "🔄",
        "titulo": "Diagnóstico",
        "subtitulo": "Status dos dados",
        "descricao": "Verifique o carregamento das duas safras, identifique inconsistências e confirme "
                     "que os dados estão prontos antes de usar as análises.",
        "tags": ["Safras", "Integridade", "Pré-análise"],
    },
    {
        "icone": "📊",
        "titulo": "Análise Conjunta",
        "subtitulo": "Produtividade por cultivar (Faixa)",
        "descricao": "Compare produtividade (kg/ha e sacas) entre cultivares no ensaio de Faixa, "
                     "por safra e região, com ranking geral e recorte multissafra.",
        "tags": ["kg/ha", "sacas", "Ranking", "Faixa"],
    },
    {
        "icone": "⚔️",
        "titulo": "Head-to-Head",
        "subtitulo": "Confronto direto entre materiais",
        "descricao": "Classificação de um cultivar versus os adversários local a local — vitórias, "
                     "empates e derrotas em produtividade, com visão por fazenda e região.",
        "tags": ["H2H", "Vitórias", "Confronto", "Local"],
    },
    {
        "icone": "🦠",
        "titulo": "Doenças",
        "subtitulo": "Sanidade da av2",
        "descricao": "Reação dos cultivares às principais doenças do milho — turcicum, cercospora, "
                     "mancha branca, bipolaris, ferrugem tropical e enfezamento — com classificação AS a R.",
        "tags": ["Doenças", "Incidência", "Classificação", "AS–R"],
    },
    {
        "icone": "🌽",
        "titulo": "Caracterização",
        "subtitulo": "Arquitetura de planta (av3)",
        "descricao": "Perfil agronômico: altura de planta e de espiga, dias até o florescimento "
                     "masculino e feminino, e sincronismo — por cultivar, safra e região.",
        "tags": ["Altura", "Florescimento", "Arquitetura"],
    },
    {
        "icone": "🌱",
        "titulo": "Análise de Densidade",
        "subtitulo": "Efeito da população de plantas",
        "descricao": "Como a densidade de plantio afeta a produtividade: resposta do material às "
                     "populações do ensaio de Densidade, com curvas e distribuição.",
        "tags": ["Densidade", "População", "Resposta"],
    },
    {
        "icone": "🗺️",
        "titulo": "Mapa",
        "subtitulo": "Desempenho por região",
        "descricao": "Produtividade média dos cultivares por estado, macro e microrregião de milho. "
                     "Identifique onde cada material se destaca geograficamente.",
        "tags": ["Mapa", "Estado", "Macro", "Micro"],
    },
]

with col_dir:
    st.markdown("""
<div style="margin: 0.2rem 0 1rem;">
    <p style="font-size:12px;font-weight:600;color:#6B7280;text-transform:uppercase;
              letter-spacing:0.07em;margin:0 0 4px;">Páginas do Painel</p>
    <h2 style="font-size:1.4rem;font-weight:700;color:#1A1A1A;margin:0;">
        O que você quer analisar hoje?
    </h2>
</div>
""", unsafe_allow_html=True)

    # Grid 2 colunas dentro da coluna direita
    _linhas = [PAGINAS[i:i+2] for i in range(0, len(PAGINAS), 2)]

    for _linha in _linhas:
        _cols = st.columns(2, gap="small")
        for _ci, _pg in enumerate(_linha):
            with _cols[_ci]:
                _tags_html = "".join([
                    f'<span style="display:inline-block;background:#E9F7EF;color:#1E8449;'
                    f'font-size:10px;font-weight:600;padding:2px 8px;border-radius:20px;'
                    f'margin:2px 2px 0 0;">{t}</span>'
                    for t in _pg["tags"]
                ])
                st.markdown(f"""
<div style="border:1px solid #E5E7EB;border-radius:12px;padding:14px;
            background:#FFFFFF;min-height:180px;
            box-shadow:0 1px 4px rgba(0,0,0,0.06);">
  <div style="font-size:22px;margin-bottom:6px;">{_pg['icone']}</div>
  <p style="font-size:14px;font-weight:700;color:#1A1A1A;margin:0 0 2px;">{_pg['titulo']}</p>
  <p style="font-size:11px;color:#6B7280;margin:0 0 8px;font-weight:500;">{_pg['subtitulo']}</p>
  <p style="font-size:12px;color:#374151;line-height:1.5;margin:0 0 10px;">{_pg['descricao']}</p>
  <div>{_tags_html}</div>
</div>
""", unsafe_allow_html=True)
                st.markdown("<div style='margin-bottom:6px;'></div>", unsafe_allow_html=True)

st.divider()
rodape()
