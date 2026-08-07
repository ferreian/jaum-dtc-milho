"""
pages/3_H2H.py — Head-to-Head · JAUM DTC Milho

Confronto direto entre híbridos, calculado apenas nos locais onde ambos foram
avaliados simultaneamente.

Abas:
  · Tab 1 — Tabela de Classificação: Produto 1 vs todos os adversários
  · Tab 2 — Análise por Local: par específico, cards + donut + mapa + barras
  · Tab 3 — Desvios por Ambiente (a construir)

Fonte: tabela_analitica_faixa das safras 2024/25 e 2025/26 (só Faixa; Densidade
tem página própria).
"""

import io

import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go_plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from utils.theme import aplicar_tema, page_header, secao_titulo, rodape
from utils.loader import carregar_multisafra
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode

st.set_page_config(
    page_title="H2H · JAUM DTC Milho",
    page_icon="⚔️",
    layout="wide",
    initial_sidebar_state="expanded",
)

aplicar_tema()

st.markdown("""
<style>
[data-testid="stCaptionContainer"] p,
[data-testid="stCaptionContainer"] { color: #374151 !important; opacity: 1 !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Constantes
# ─────────────────────────────────────────────────────────────────────────────
STATUS_P1 = ["STINE", "EXP", "DP2"]     # pool de "Produto 1" (materiais Stine)
EMPATE_MARGEM = 1.0                      # ± sc/ha: dentro disso é empate técnico
SAFRA_PADRAO = ("25/26", "2025/26")      # aceita os dois formatos de rótulo

COR_VITORIA = "#27AE60"
COR_EMPATE = "#FFFF00"
COR_DERROTA = "#FF0000"


def classificar_h2h(pct: float):
    """Classe e cor de fundo a partir do % de vitórias."""
    if pd.isna(pct):
        return "—", "#F3F4F6"
    if pct <= 45:
        return "Restrito", "#FF0000"
    if pct <= 55:
        return "Competitivo", "#FFFF00"
    if pct <= 75:
        return "Superior", "#87CEFF"
    return "Alta Performance", "#90EE90"


COR_STATUS_TITULO = {"CHECK": "#F4B184", "STINE": "#2976B6",
                     "EXP": "#00FF00", "DP2": "#C4DFB4"}


# ─────────────────────────────────────────────────────────────────────────────
# Helper Excel — mesma paleta da classificação
# ─────────────────────────────────────────────────────────────────────────────
def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "H2H"

    df = df.reset_index(drop=True)
    df = df.loc[:, ~df.columns.str.startswith("_")].copy()

    thin = Side(style="thin", color="CCCCCC")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ci, col in enumerate(df.columns, 1):
        c = ws.cell(row=1, column=ci, value=str(col))
        c.font = Font(bold=True, name="Arial", size=10, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="4A4A4A")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = brd
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = max(14, len(str(col)) + 4)
    ws.row_dimensions[1].height = 28

    COR_FUNDO_XL = {"Alta Performance": "90EE90", "Superior": "87CEFF",
                    "Competitivo": "FFFF00", "Restrito": "FF0000", "—": "F3F4F6"}
    COR_FONTE_XL = {"Alta Performance": "1A1A1A", "Superior": "1A1A1A",
                    "Competitivo": "1A1A1A", "Restrito": "FFFFFF", "—": "6B7280"}

    idx_classe = df.columns.tolist().index("Classe") if "Classe" in df.columns else None

    for ri, row_data in enumerate(df.itertuples(index=False), start=2):
        classe_val = str(row_data[idx_classe]) if idx_classe is not None else "—"
        bg = COR_FUNDO_XL.get(classe_val, "FFFFFF")
        fg = COR_FONTE_XL.get(classe_val, "1A1A1A")
        for ci, val in enumerate(row_data, 1):
            try:
                val = None if (val is None or (isinstance(val, float) and np.isnan(val))) else val
            except (TypeError, ValueError):
                pass
            c = ws.cell(row=ri, column=ci, value=val)
            if idx_classe is not None and ci == idx_classe + 1:
                c.font = Font(name="Arial", size=10, color=fg, bold=True)
                c.fill = PatternFill("solid", start_color=bg)
            else:
                c.font = Font(name="Arial", size=10, color="1A1A1A")
                c.fill = PatternFill("solid", start_color="FFFFFF")
            c.alignment = Alignment(horizontal="left" if ci == 1 else "center", vertical="center")
            c.border = brd

    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# Helper AgGrid — coluna Classe colorida
# ─────────────────────────────────────────────────────────────────────────────
def ag_table_h2h(df: pd.DataFrame, height: int = 480, estilos_col=None, renderers_col=None):
    """Tabela AgGrid padrão do H2H.
    `estilos_col`: {coluna: JsCode} para cellStyle · `renderers_col`: {coluna: JsCode} para
    cellRenderer (permite desenhar barras dentro da célula)."""
    classe_style = JsCode("""
    function(params) {
        const v = params.value;
        if (v === 'Alta Performance') return {'backgroundColor':'#90EE90','color':'#1A1A1A','fontWeight':'700','textAlign':'center'};
        if (v === 'Superior')         return {'backgroundColor':'#87CEFF','color':'#1A1A1A','fontWeight':'700','textAlign':'center'};
        if (v === 'Competitivo')      return {'backgroundColor':'#FFFF00','color':'#1A1A1A','fontWeight':'700','textAlign':'center'};
        if (v === 'Restrito')         return {'backgroundColor':'#FF0000','color':'#FFFFFF','fontWeight':'700','textAlign':'center'};
        return {};
    }
    """)
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        resizable=True, sortable=True, filter=True, suppressMenu=False,
        menuTabs=["generalMenuTab", "filterMenuTab", "columnsMenuTab"],
        cellStyle={"fontSize": "13px", "color": "#000000",
                   "fontFamily": "Helvetica Neue, sans-serif"})
    if "Classe" in df.columns:
        gb.configure_column("Classe", cellStyle=classe_style, minWidth=140)
    for _c, _e in (estilos_col or {}).items():
        if _c in df.columns:
            gb.configure_column(_c, cellStyle=_e)
    for _c, _r in (renderers_col or {}).items():
        if _c in df.columns:
            gb.configure_column(_c, cellRenderer=_r, minWidth=170)
    gb.configure_grid_options(headerHeight=36, rowHeight=32, domLayout="normal",
                              suppressMenuHide=True, suppressColumnVirtualisation=True,
                              suppressContextMenu=False, enableRangeSelection=True)
    go = gb.build()
    go["onFirstDataRendered"] = JsCode("function(params) { params.api.sizeColumnsToFit(); }")
    AgGrid(df, gridOptions=go, height=height, update_mode=GridUpdateMode.NO_UPDATE,
           fit_columns_on_grid_load=False, columns_auto_size_mode=2,
           allow_unsafe_jscode=True, enable_enterprise_modules=True,
           custom_css={
               ".ag-header": {"background-color": "#4A4A4A !important"},
               ".ag-header-row": {"background-color": "#4A4A4A !important"},
               ".ag-header-cell": {"background-color": "#4A4A4A !important"},
               ".ag-header-cell-label": {"color": "#FFFFFF !important", "font-weight": "700"},
               ".ag-header-cell-text": {"color": "#FFFFFF !important", "font-size": "13px !important",
                                        "font-weight": "700 !important"},
               ".ag-icon": {"color": "#FFFFFF !important", "opacity": "1 !important"},
               ".ag-header-icon": {"color": "#FFFFFF !important", "opacity": "1 !important"},
               ".ag-header-cell-menu-button": {"opacity": "1 !important", "visibility": "visible !important"},
               ".ag-cell": {"font-size": "13px !important"},
               ".ag-row": {"font-size": "13px !important"},
           },
           theme="streamlit", use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# Carregamento — analítica de Faixa das duas safras, já reconciliada pelo
# depara_mestre (sem ele, o mesmo híbrido teria nomes diferentes em cada safra)
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def carregar_concat():
    d = carregar_multisafra()
    df = d.get("tabela_analitica_faixa")
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    df = df.copy()
    if "produtividade_valida_kg_ha" in df.columns:
        df["kg_ha"] = pd.to_numeric(df["produtividade_valida_kg_ha"], errors="coerce")
        if "produtividade_kg_ha" in df.columns:
            df["kg_ha"] = df["kg_ha"].fillna(pd.to_numeric(df["produtividade_kg_ha"], errors="coerce"))
    elif "produtividade_kg_ha" in df.columns:
        df["kg_ha"] = pd.to_numeric(df["produtividade_kg_ha"], errors="coerce")
    if "produtividade_valida_sacas_ha" in df.columns:
        df["sc_ha"] = pd.to_numeric(df["produtividade_valida_sacas_ha"], errors="coerce")
    elif "kg_ha" in df.columns:
        df["sc_ha"] = (df["kg_ha"] / 60).round(1)
    return df


with st.spinner("Carregando dados..."):
    ta_raw = carregar_concat()

if ta_raw.empty:
    st.error("Nenhum dado disponível. Verifique a página de Diagnóstico.")
    st.stop()

page_header("Head-to-Head",
            "Compare híbridos diretamente nos locais em que ambos foram avaliados simultaneamente.",
            imagem="Data analysis-pana.png")


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar — filtros encadeados
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<p style="font-size:11px;font-weight:600;color:#6B7280;text-transform:uppercase;'
                'letter-spacing:0.05em;padding:0.5rem;">Filtros</p>', unsafe_allow_html=True)

    if st.button("Limpar filtros", use_container_width=True):
        for key in list(st.session_state.keys()):
            if (key.startswith("h2h_") or key.startswith("__opts_h2h_")
                    or key.startswith("busca_h2h_")):
                del st.session_state[key]
        st.rerun()

    def _podar_keys(prefix, opcoes, molde):
        """Remove estado de checkboxes de opções que saíram da cascata —
        sem isso elas reaparecem marcadas e o filtro se reaplica sozinho."""
        antigas = st.session_state.get(f"__opts_{prefix}", [])
        atuais = set(map(str, opcoes))
        for o in antigas:
            if str(o) not in atuais:
                st.session_state.pop(molde(o), None)
        st.session_state[f"__opts_{prefix}"] = list(opcoes)

    def checkboxes(opcoes, default_all=True, defaults=None, prefix=""):
        _podar_keys(prefix, opcoes, lambda o: f"{prefix}_{o}")
        sel = []
        for o in opcoes:
            key = f"{prefix}_{o}"
            if key in st.session_state:
                marcado = st.checkbox(str(o), key=key)
            else:
                checked = (o in defaults) if defaults is not None else default_all
                marcado = st.checkbox(str(o), value=checked, key=key)
            if marcado:
                sel.append(o)
        return sel

    def filtro_busca(opcoes, prefix):
        """Busca textual + seleção persistente. A seleção é lida direto dos checkboxes
        (fonte única), inclusive dos itens ocultos pela busca."""
        if f"{prefix}_reset" not in st.session_state:
            st.session_state[f"{prefix}_reset"] = 0
        r = st.session_state[f"{prefix}_reset"]
        _podar_keys(prefix, opcoes, lambda o: f"{prefix}_chk_{r}_{o}")

        busca = st.text_input("Buscar", value="", key=f"busca_{prefix}",
                              placeholder="Digite para filtrar...")
        filtradas = ([c for c in opcoes if busca.strip().lower() in str(c).lower()]
                     if busca.strip() else opcoes)

        if st.button("Limpar seleção", key=f"{prefix}_limpar", use_container_width=True):
            for o in opcoes:
                st.session_state.pop(f"{prefix}_chk_{r}_{o}", None)
            st.session_state.pop(f"__opts_{prefix}", None)
            st.session_state[f"{prefix}_reset"] = r + 1
            st.rerun()

        for c in filtradas:
            st.checkbox(str(c), key=f"{prefix}_chk_{r}_{c}")

        sel = [o for o in opcoes if st.session_state.get(f"{prefix}_chk_{r}_{o}", False)]
        return sel or opcoes

    # 1. Safra — padrão safra atual
    with st.expander("Safra", expanded=True):
        safras_all = sorted(ta_raw["safra"].dropna().unique().tolist())
        safra_default = [s for s in safras_all if str(s) in SAFRA_PADRAO] or safras_all[-1:]
        if "h2h_safra_init" not in st.session_state:
            for o in safras_all:
                st.session_state[f"h2h_safra_{o}"] = (o in safra_default)
            st.session_state["h2h_safra_init"] = True
        safras_sel = checkboxes(safras_all, defaults=safra_default, prefix="h2h_safra")
    ta_f1 = ta_raw[ta_raw["safra"].isin(safras_sel)] if safras_sel else ta_raw.iloc[0:0]

    # 2. Região Macro
    with st.expander("Região Macro", expanded=False):
        macros_sel = checkboxes(sorted(ta_f1["regiao_macro"].dropna().unique().tolist()), prefix="h2h_macro")
    ta_f2 = ta_f1[ta_f1["regiao_macro"].isin(macros_sel)] if macros_sel else ta_f1.iloc[0:0]

    # 3. Região Micro
    with st.expander("Região Micro", expanded=False):
        micros_sel = checkboxes(sorted(ta_f2["regiao_micro"].dropna().unique().tolist()), prefix="h2h_micro")
    ta_f3 = ta_f2[ta_f2["regiao_micro"].isin(micros_sel)] if micros_sel else ta_f2.iloc[0:0]

    # 4. Estado
    with st.expander("Estado", expanded=False):
        estados_sel = filtro_busca(sorted(ta_f3["estado_sigla"].dropna().unique().tolist()), "h2h_estado")
    ta_f4 = ta_f3[ta_f3["estado_sigla"].isin(estados_sel)] if estados_sel else ta_f3.iloc[0:0]

    # 5. Cidade
    with st.expander("Cidade", expanded=False):
        cidades_sel = filtro_busca(sorted(ta_f4["cidade_nome"].dropna().unique().tolist()), "h2h_cidade")
    ta_f5 = ta_f4[ta_f4["cidade_nome"].isin(cidades_sel)] if cidades_sel else ta_f4.iloc[0:0]

    # 6. Fazenda
    with st.expander("Fazenda", expanded=False):
        fazendas_sel = filtro_busca(sorted(ta_f5["nomeFazenda"].dropna().unique().tolist()), "h2h_fazenda")
    ta_f6 = ta_f5[ta_f5["nomeFazenda"].isin(fazendas_sel)] if fazendas_sel else ta_f5.iloc[0:0]

    # 7. Responsável
    with st.expander("Responsável", expanded=False):
        resps_sel = filtro_busca(sorted(ta_f6["nomeResponsavel"].dropna().unique().tolist()), "h2h_resp")
    ta_filtrado = ta_f6[ta_f6["nomeResponsavel"].isin(resps_sel)] if resps_sel else ta_f6.iloc[0:0]

    # 8. Status do adversário (Produto 2) — padrão CHECK
    with st.expander("Status do Adversário (Prod. 2)", expanded=True):
        status_p2_all = sorted(ta_filtrado["status_material"].dropna().unique().tolist())
        status_p2_default = [s for s in ["CHECK"] if s in status_p2_all]
        if "h2h_p2status_init" not in st.session_state and status_p2_all:
            for o in status_p2_all:
                st.session_state[f"h2h_p2status_{o}"] = (o in status_p2_default)
            st.session_state["h2h_p2status_init"] = True
        status_p2_sel = checkboxes(status_p2_all, defaults=status_p2_default, prefix="h2h_p2status")

if ta_filtrado.empty:
    st.warning("Nenhum dado para os filtros selecionados.")
    st.stop()

df_p1 = ta_filtrado[ta_filtrado["status_material"].isin(STATUS_P1)].copy()
df_p2 = (ta_filtrado[ta_filtrado["status_material"].isin(status_p2_sel)].copy()
         if status_p2_sel else pd.DataFrame())


# ─────────────────────────────────────────────────────────────────────────────
# Cruzamento por local (o join pesado — cacheado)
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def cruzar_por_local(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    """Cruza df1 × df2 pelo local. Agrega por local antes do merge para não
    duplicar parcelas. Uma linha por (dePara_1, dePara_2, cod_fazenda)."""
    cols = ["dePara", "status_material", "cod_fazenda", "sc_ha", "kg_ha"]
    d1 = df1[[c for c in cols if c in df1.columns]].dropna(subset=["sc_ha"]).copy()
    d2 = df2[[c for c in cols if c in df2.columns]].dropna(subset=["sc_ha"]).copy()
    if d1.empty or d2.empty:
        return pd.DataFrame()
    d1 = d1.groupby(["dePara", "status_material", "cod_fazenda"], as_index=False)[["sc_ha", "kg_ha"]].mean()
    d2 = d2.groupby(["dePara", "status_material", "cod_fazenda"], as_index=False)[["sc_ha", "kg_ha"]].mean()
    merged = d1.merge(d2, on="cod_fazenda", suffixes=("_1", "_2"))
    return merged[merged["dePara_1"] != merged["dePara_2"]].reset_index(drop=True)


def linha_safra(s, g, sufixo_locais="locais"):
    """Uma linha de contexto por safra, com as regiões/estados que ela realmente contém."""
    partes = [f"<b>Safra:</b> {s}"]
    if "regiao_macro" in g.columns and g["regiao_macro"].notna().any():
        partes.append("Macro: " + ", ".join(sorted(g["regiao_macro"].dropna().unique())))
    if "regiao_micro" in g.columns and g["regiao_micro"].notna().any():
        partes.append("Micro: " + ", ".join(sorted(g["regiao_micro"].dropna().unique())))
    if "estado_sigla" in g.columns and g["estado_sigla"].notna().any():
        partes.append(", ".join(sorted(g["estado_sigla"].dropna().unique())))
    partes.append(f"{g['cidade_nome'].nunique()} cidades")
    partes.append(f"{g['cod_fazenda'].nunique()} {sufixo_locais}")
    return " · ".join(partes)


def montar_contexto(base, sufixo_locais="locais"):
    grupos = list(base.groupby("safra")) if "safra" in base.columns else []
    linhas = [linha_safra(s, g, sufixo_locais) for s, g in grupos]
    if len(grupos) > 1:
        linhas = ["<b>Análise multissafra</b>"] + linhas
    return "<br>".join(linhas)


tab1, tab2, tab3 = st.tabs(["Tabela de Classificação", "Análise por Local", "Desvios por Ambiente"])


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Tabela de Classificação
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    secao_titulo(
        "HEAD-TO-HEAD · TABELA",
        "Classificação vs adversários",
        "Escolha o Produto 1 (STINE / EXP / DP2) e veja como ele se comporta contra cada "
        "adversário nos locais em que ambos foram avaliados.",
    )

    st.markdown(
        '<div style="background:#FFF8E1;border:1px solid #F5D76E;border-left:5px solid #D4A800;'
        'border-radius:8px;padding:14px 18px;margin:6px 0 14px;">'
        '<p style="margin:0;font-size:15.5px;line-height:1.6;color:#4A3B00;'
        'font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">'
        '<span style="font-size:19px;vertical-align:-1px;margin-right:8px;">⚠️</span>'
        'O resultado de cada local <b>não muda</b> com os filtros — o que muda é <b>quais locais '
        'entram na conta</b>. Por isso as médias, a % de vitórias e a classificação se recalculam '
        'sozinhos ao filtrar, e não é preciso rodar a análise de novo; o botão só é necessário ao '
        '<b>trocar o híbrido selecionado</b>.<br><br>'
        '<b>Ao comparar recortes diferentes, olhe sempre o nº de comparações.</b> Uma classificação '
        'obtida em 4 locais não tem o mesmo peso que a mesma classificação em 30 — o subtítulo acima '
        'mostra o recorte vigente e a tabela traz o nº de locais de cada confronto.'
        '</p></div>',
        unsafe_allow_html=True,
    )

    if df_p1.empty:
        st.warning("Nenhum híbrido STINE/EXP/DP2 encontrado com os filtros atuais.")
    elif df_p2.empty:
        st.warning("Nenhum adversário disponível. Verifique o filtro 'Status do Adversário' na barra lateral.")
    else:
        hibridos_p1 = sorted(df_p1["dePara"].dropna().unique())
        col_sel, col_btn = st.columns([4, 1])
        with col_sel:
            p1_t1 = st.selectbox("Produto 1 (STINE / EXP / DP2)", hibridos_p1, key="p1_t1")
        with col_btn:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            btn_t1 = st.button("Rodar análise", type="primary", key="btn_t1", use_container_width=True)

        key_t1 = f"res_t1_raw__{p1_t1}"

        if btn_t1:
            with st.spinner("Calculando confrontos..."):
                # roda contra TODOS os adversários da base — os filtros afetam só a exibição
                _p1_full = ta_raw[(ta_raw["status_material"].isin(STATUS_P1)) & (ta_raw["dePara"] == p1_t1)]
                _p2_full = ta_raw[ta_raw["dePara"] != p1_t1]
                df_cross_raw = cruzar_por_local(_p1_full, _p2_full)
                if not df_cross_raw.empty:
                    _meta = (ta_raw[["cod_fazenda", "safra", "regiao_macro", "regiao_micro",
                                     "estado_sigla", "cidade_nome", "nomeFazenda", "nomeResponsavel"]]
                             .drop_duplicates("cod_fazenda"))
                    df_cross_raw = df_cross_raw.merge(_meta, on="cod_fazenda", how="left")
                st.session_state[key_t1] = df_cross_raw

        if key_t1 not in st.session_state:
            st.info("Selecione o Produto 1 e clique em **Rodar análise** para calcular.")
        else:
            _cross_raw = st.session_state[key_t1]

            if _cross_raw.empty:
                st.info("Nenhum confronto encontrado para este híbrido.")
            else:
                # exibição respeita os filtros: locais ativos + status do adversário
                _locais_ativos = set(ta_filtrado["cod_fazenda"].dropna().unique())
                _cross_filt = _cross_raw[_cross_raw["cod_fazenda"].isin(_locais_ativos)]
                if status_p2_sel and "status_material_2" in _cross_filt.columns:
                    _cross_filt = _cross_filt[_cross_filt["status_material_2"].isin(status_p2_sel)]

                _rows = []
                for prod2, grp in _cross_filt.groupby("dePara_2"):
                    n = len(grp)
                    diff = grp["sc_ha_1"] - grp["sc_ha_2"]
                    vit = int((diff > EMPATE_MARGEM).sum())
                    emp = int((diff.abs() <= EMPATE_MARGEM).sum())
                    der = n - vit - emp
                    pct = round(vit / n * 100, 1) if n > 0 else np.nan
                    sc1, sc2 = grp["sc_ha_1"].mean(), grp["sc_ha_2"].mean()
                    kg1 = grp["kg_ha_1"].mean() if "kg_ha_1" in grp.columns else np.nan
                    kg2 = grp["kg_ha_2"].mean() if "kg_ha_2" in grp.columns else np.nan
                    dif_kg = (kg1 - kg2) if not (np.isnan(kg1) or np.isnan(kg2)) else np.nan
                    dif_sc = (sc1 - sc2) if not (np.isnan(sc1) or np.isnan(sc2)) else np.nan
                    dif_pct = ((sc1 / sc2) - 1) * 100 if (sc2 and not np.isnan(sc2)) else np.nan
                    classe, _ = classificar_h2h(pct)
                    _rows.append({
                        "Produto 1": p1_t1,
                        "kg/ha Prod 1": round(kg1, 0) if not np.isnan(kg1) else None,
                        "sc/ha Prod 1": round(sc1, 1),
                        "Produto 2": prod2,
                        "kg/ha Prod 2": round(kg2, 0) if not np.isnan(kg2) else None,
                        "sc/ha Prod 2": round(sc2, 1),
                        "Vitórias": vit,
                        "Empates": emp,
                        "Derrotas": der,
                        "Nº Comparações": n,
                        "Dif. (%)": round(dif_pct, 1) if not np.isnan(dif_pct) else None,
                        "Dif. (kg/ha)": round(dif_kg, 0) if not np.isnan(dif_kg) else None,
                        "Dif. (sc/ha)": round(dif_sc, 1) if not np.isnan(dif_sc) else None,
                        "% Vitórias": pct,
                        "Classe": classe,
                    })

                df_res = pd.DataFrame(_rows)
                if not df_res.empty:
                    df_res = df_res.sort_values("% Vitórias", ascending=False).reset_index(drop=True)

                if df_res.empty:
                    st.info("Nenhum confronto encontrado — o híbrido não compartilha locais com os "
                            "adversários selecionados.")
                else:
                    contexto_str = montar_contexto(_cross_filt)

                    _st_p1 = (df_p1[df_p1["dePara"] == p1_t1]["status_material"].iloc[0]
                              if not df_p1[df_p1["dePara"] == p1_t1].empty else "")
                    st.markdown(
                        f'<div style="margin:0.5rem 0 0.2rem;">'
                        f'<p style="font-size:13px;font-weight:600;color:#6B7280;text-transform:uppercase;'
                        f'letter-spacing:0.05em;margin:0 0 4px;">Análise H2H · Produto 1</p>'
                        f'<h2 style="font-size:1.9rem;font-weight:700;color:#1A1A1A;margin:0;line-height:1.2;">'
                        f'<span style="color:#27AE60;">{p1_t1}</span>'
                        f'<span style="font-size:1rem;font-weight:500;color:#6B7280;margin-left:10px;">'
                        f'{_st_p1} · {len(df_res)} adversários</span></h2>'
                        f'<p style="font-size:14px;color:#6B7280;margin:4px 0 0;">{contexto_str}</p></div>',
                        unsafe_allow_html=True)

                    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

                    with st.popover("ℹ️ Como interpretar esta tabela", use_container_width=False):
                        st.markdown(f"""
**📌 O que esta tabela mostra**

Cada linha é um confronto direto entre o **Produto 1** e um adversário (Produto 2), calculado
exclusivamente nos locais onde **ambos foram avaliados simultaneamente**.

---

**📐 Como ler as colunas**

- **kg/ha · sc/ha Prod 1 / Prod 2** → médias *apenas nos locais compartilhados* (não é a média
  geral do híbrido).
- **Vitórias / Empates / Derrotas** → nº de locais em cada resultado.
- **Nº Comparações** → total de locais com ambos avaliados.
- **% Vitórias** → Vitórias ÷ Comparações × 100 — é a base da classificação.
- **Dif. (%)** → quanto o Produto 1 produz a mais ou a menos, em termos relativos.
- **Dif. (kg/ha) e (sc/ha)** → diferença absoluta média entre os dois.

> ⚠️ **Empate**: diferença de até **±{EMPATE_MARGEM:.0f} sc/ha** não conta como vitória nem derrota.

---

**🎨 Legenda das cores — % de vitórias**
""")
                        ca, cb, cc, cd = st.columns(4)
                        ca.markdown('<div style="background:#90EE90;border-radius:6px;padding:8px;text-align:center;">'
                                    '<b style="color:#1A1A1A;">Alta Performance</b><br>'
                                    '<span style="font-size:12px;color:#1A1A1A;">&gt; 75% de vitórias</span></div>',
                                    unsafe_allow_html=True)
                        cb.markdown('<div style="background:#87CEFF;border-radius:6px;padding:8px;text-align:center;">'
                                    '<b style="color:#1A1A1A;">Superior</b><br>'
                                    '<span style="font-size:12px;color:#1A1A1A;">56 – 75% de vitórias</span></div>',
                                    unsafe_allow_html=True)
                        cc.markdown('<div style="background:#FFFF00;border-radius:6px;padding:8px;text-align:center;">'
                                    '<b style="color:#1A1A1A;">Competitivo</b><br>'
                                    '<span style="font-size:12px;color:#1A1A1A;">46 – 55% de vitórias</span></div>',
                                    unsafe_allow_html=True)
                        cd.markdown('<div style="background:#FF0000;border-radius:6px;padding:8px;text-align:center;">'
                                    '<b style="color:#FFFFFF;">Restrito</b><br>'
                                    '<span style="font-size:12px;color:#FFFFFF;">≤ 45% de vitórias</span></div>',
                                    unsafe_allow_html=True)
                        st.markdown("""
---

**💡 Como interpretar**

- **Alta Performance** → vence em mais de 3/4 dos locais: consistentemente superior a esse adversário.
- **Superior** → vence na maioria dos locais.
- **Competitivo** → resultado equilibrado, nenhum se destaca claramente.
- **Restrito** → perde na maioria dos locais frente a esse adversário — atenção ao posicionamento.

> As médias mudam conforme o adversário, porque cada confronto usa só os locais em comum entre os
dois. Por isso o mesmo híbrido pode aparecer com médias diferentes em linhas diferentes.
""")

                    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

                    contagem = df_res["Classe"].value_counts()
                    total_cls = len(df_res)
                    c1, c2, c3, c4 = st.columns(4)
                    for col_ui, label, cor_txt in zip(
                            [c1, c2, c3, c4],
                            ["Alta Performance", "Superior", "Competitivo", "Restrito"],
                            ["#27AE60", "#1E40AF", "#F2C811", "#FF0000"]):
                        n_cls = int(contagem.get(label, 0))
                        pct_cl = f"{n_cls / total_cls * 100:.0f}%" if total_cls else "—"
                        col_ui.markdown(
                            f'<div style="border:1px solid #E5E7EB;border-radius:10px;padding:10px 14px;'
                            f'background:#FFFFFF;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,0.07);">'
                            f'<p style="margin:0;font-size:14px;font-weight:600;color:#374151;">{label}</p>'
                            f'<p style="margin:4px 0 0;font-size:2.2rem;font-weight:700;color:{cor_txt};">'
                            f'{n_cls} <span style="font-size:1.2rem;font-weight:500;">({pct_cl})</span></p></div>',
                            unsafe_allow_html=True)

                    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

                    ag_table_h2h(df_res, height=min(680, int((36 + 32 * len(df_res) + 20) * 1.3)))
                    st.download_button("⬇️ Exportar Excel", data=to_excel(df_res),
                                       file_name=f"h2h_{p1_t1}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       key="dl_t1")
                    st.caption(f"ℹ️ Confrontos calculados só nos locais compartilhados por cada par. "
                               f"Empate = diferença de até ±{EMPATE_MARGEM:.0f} sc/ha.")


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Análise por Local (a construir)
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    secao_titulo(
        "HEAD-TO-HEAD · POR LOCAL",
        "Diferença de produtividade por local",
        "Selecione um par específico e veja a diferença de sc/ha em cada local compartilhado.",
    )

    st.markdown(
        '<div style="background:#FFF8E1;border:1px solid #F5D76E;border-left:5px solid #D4A800;'
        'border-radius:8px;padding:14px 18px;margin:6px 0 14px;">'
        '<p style="margin:0;font-size:15.5px;line-height:1.6;color:#4A3B00;'
        'font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">'
        '<span style="font-size:19px;vertical-align:-1px;margin-right:8px;">⚠️</span>'
        'O resultado de cada local <b>não muda</b> com os filtros — o que muda é <b>quais locais '
        'entram na conta</b>. Por isso as médias, a % de vitórias e a classificação se recalculam '
        'sozinhos ao filtrar, e não é preciso rodar a análise de novo; o botão só é necessário ao '
        '<b>trocar o híbrido selecionado</b>.<br><br>'
        '<b>Ao comparar recortes diferentes, olhe sempre o nº de comparações.</b> Uma classificação '
        'obtida em 4 locais não tem o mesmo peso que a mesma classificação em 30 — o subtítulo acima '
        'mostra o recorte vigente e a tabela traz o nº de locais de cada confronto.'
        '</p></div>',
        unsafe_allow_html=True,
    )

    if df_p1.empty or df_p2.empty:
        st.warning("Dados insuficientes com os filtros atuais.")
    else:
        hibridos_p1_t2 = sorted(df_p1["dePara"].dropna().unique())
        col_a2, col_b2, col_c2 = st.columns([2, 2, 1])

        with col_a2:
            p1_t2 = st.selectbox("Produto 1 (STINE / EXP / DP2)", hibridos_p1_t2, key="p1_t2")

        locais_p1_t2 = set(df_p1[df_p1["dePara"] == p1_t2]["cod_fazenda"].dropna().unique())
        adv_disp = sorted(df_p2[df_p2["cod_fazenda"].isin(locais_p1_t2)]["dePara"].dropna().unique())

        with col_b2:
            if adv_disp:
                p2_t2 = st.selectbox("Produto 2 (adversário)", adv_disp, key="p2_t2")
            else:
                st.warning("Nenhum adversário com locais em comum para este híbrido.")
                p2_t2 = None

        with col_c2:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            btn_t2 = st.button("Rodar análise", type="primary", key="btn_t2", use_container_width=True)

        key_t2 = f"res_t2_raw__{p1_t2}__{p2_t2}"

        if btn_t2 and p2_t2:
            with st.spinner("Calculando..."):
                # cruza o par em TODOS os locais da base — os filtros afetam só a exibição
                d1_loc = (ta_raw[ta_raw["dePara"] == p1_t2][["cod_fazenda", "sc_ha", "kg_ha"]]
                          .dropna(subset=["sc_ha"])
                          .groupby("cod_fazenda", as_index=False)[["sc_ha", "kg_ha"]].mean())
                d2_loc = (ta_raw[ta_raw["dePara"] == p2_t2][["cod_fazenda", "sc_ha", "kg_ha"]]
                          .dropna(subset=["sc_ha"])
                          .groupby("cod_fazenda", as_index=False)[["sc_ha", "kg_ha"]].mean())
                st.session_state[key_t2] = d1_loc.merge(d2_loc, on="cod_fazenda",
                                                        suffixes=("_1", "_2")).copy()

        if key_t2 not in st.session_state or not p2_t2:
            st.info("Selecione os dois híbridos e clique em **Rodar análise** para calcular.")
        else:
            # aplica os filtros vigentes na exibição (mesma lógica da Tab 1)
            _loc_ativos_t2 = set(ta_filtrado["cod_fazenda"].dropna().unique())
            df_loc = st.session_state[key_t2]
            df_loc = df_loc[df_loc["cod_fazenda"].isin(_loc_ativos_t2)].copy()
            if not df_loc.empty:
                df_loc["diff_sc"] = df_loc["sc_ha_1"] - df_loc["sc_ha_2"]
                df_loc["resultado"] = df_loc["diff_sc"].apply(
                    lambda x: "Vitória" if x > EMPATE_MARGEM
                    else ("Empate" if abs(x) <= EMPATE_MARGEM else "Derrota"))
                df_loc = df_loc.sort_values("diff_sc").reset_index(drop=True)

            if df_loc.empty:
                st.info("Nenhum local compartilhado encontrado para este par.")
            else:
                n_loc = len(df_loc)
                n_vit = int((df_loc["resultado"] == "Vitória").sum())
                n_emp = int((df_loc["resultado"] == "Empate").sum())
                n_der = int((df_loc["resultado"] == "Derrota").sum())
                vit_sc = df_loc.loc[df_loc["resultado"] == "Vitória", "diff_sc"]
                der_sc = df_loc.loc[df_loc["resultado"] == "Derrota", "diff_sc"]
                max_vit = float(vit_sc.max()) if len(vit_sc) else np.nan
                med_vit = float(vit_sc.mean()) if len(vit_sc) else np.nan
                min_der = float(der_sc.min()) if len(der_sc) else np.nan
                med_der = float(der_sc.mean()) if len(der_sc) else np.nan

                _base_t2 = ta_filtrado[ta_filtrado["cod_fazenda"].isin(df_loc["cod_fazenda"])]
                contexto_t2 = montar_contexto(_base_t2, "locais compartilhados")

                _st1 = (df_p1[df_p1["dePara"] == p1_t2]["status_material"].iloc[0]
                        if not df_p1[df_p1["dePara"] == p1_t2].empty else "")
                _st2 = (df_p2[df_p2["dePara"] == p2_t2]["status_material"].iloc[0]
                        if not df_p2[df_p2["dePara"] == p2_t2].empty else "")

                st.markdown(
                    f'<div style="margin:0.5rem 0 0.2rem;">'
                    f'<p style="font-size:13px;font-weight:600;color:#6B7280;text-transform:uppercase;'
                    f'letter-spacing:0.05em;margin:0 0 4px;">Análise H2H · Confronto direto</p>'
                    f'<h2 style="font-size:1.9rem;font-weight:700;color:#1A1A1A;margin:0;line-height:1.2;">'
                    f'<span style="color:#27AE60;">{p1_t2}</span>'
                    f'<span style="font-size:0.85rem;font-weight:500;color:#6B7280;margin-left:6px;">{_st1}</span>'
                    f'<span style="font-size:1.1rem;font-weight:500;color:#6B7280;margin:0 12px;">vs</span>'
                    f'<span style="color:#1A1A1A;">{p2_t2}</span>'
                    f'<span style="font-size:0.85rem;font-weight:500;color:#6B7280;margin-left:6px;">{_st2}</span>'
                    f'</h2><p style="font-size:14px;color:#6B7280;margin:4px 0 0;">{contexto_t2}</p></div>',
                    unsafe_allow_html=True)

                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

                col_pop1, col_pop2, _ = st.columns([2, 2, 4])
                with col_pop1:
                    with st.popover("ℹ️ Como interpretar", use_container_width=True):
                        st.markdown(f"""
**📌 O que este painel mostra**

Confronto direto entre **Produto 1** e **Produto 2** nos locais onde **ambos foram avaliados
simultaneamente**.

---

**📐 Como ler os cards**

- **Locais avaliados** → total de locais com os dois híbridos presentes.
- **Vitórias** → locais em que o Produto 1 superou por mais de **+{EMPATE_MARGEM:.0f} sc/ha**
  (*Max* = maior diferença positiva; *Média* = média das vitórias).
- **Empates** → diferença dentro de ±{EMPATE_MARGEM:.0f} sc/ha.
- **Derrotas** → locais em que ficou abaixo por mais de {EMPATE_MARGEM:.0f} sc/ha
  (*Min* = pior diferença; *Média* = média das derrotas).

---

**🗺️ Mapa e barras**

- No **mapa**, a cor do ponto indica a magnitude da diferença naquele local, e o símbolo distingue
  a safra quando há mais de uma ativa.
- Nas **barras**, cada linha é um local, ordenado da maior derrota à maior vitória. Verde =
  vitória, amarelo = empate, vermelho = derrota.

> Uma leitura útil: muitos locais próximos de zero indicam materiais equivalentes; poucos locais
com diferenças grandes em ambos os sentidos indicam que o ambiente decide o vencedor.
""")
                with col_pop2:
                    with st.popover("Dicionário de locais", use_container_width=True):
                        df_dic_t2 = (ta_filtrado[ta_filtrado["cod_fazenda"].isin(df_loc["cod_fazenda"])]
                                     [["cod_fazenda", "nomeFazenda", "cidade_nome", "estado_sigla"]]
                                     .drop_duplicates()
                                     .sort_values(["estado_sigla", "cidade_nome", "cod_fazenda"])
                                     .rename(columns={"cod_fazenda": "Código", "nomeFazenda": "Local",
                                                      "cidade_nome": "Cidade", "estado_sigla": "Estado"})
                                     .reset_index(drop=True))
                        st.markdown(f"Referência dos **{len(df_dic_t2)} locais** do confronto.")
                        st.dataframe(df_dic_t2, hide_index=True, use_container_width=True)

                st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

                pct_vit = f"{n_vit / n_loc * 100:.0f}%" if n_loc else "—"
                pct_emp = f"{n_emp / n_loc * 100:.0f}%" if n_loc else "—"
                pct_der = f"{n_der / n_loc * 100:.0f}%" if n_loc else "—"
                COR_EMPATE_CARD = "#D4A800"
                card = ("border:1px solid #E5E7EB;border-radius:10px;padding:12px 16px;"
                        "background:#FFFFFF;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,0.07);")

                k1, k2, k3, k4 = st.columns(4)
                k1.markdown(f'<div style="{card}">'
                            f'<p style="margin:0;font-size:12px;color:#6B7280;">Locais avaliados</p>'
                            f'<p style="margin:6px 0 0;font-size:1.9rem;font-weight:700;color:#1A1A1A;">{n_loc}</p>'
                            f'</div>', unsafe_allow_html=True)

                sub_v = (f'<p style="margin:2px 0;font-size:14px;font-weight:600;color:{COR_VITORIA};">Max: {max_vit:+.1f} sc/ha</p>'
                         f'<p style="margin:0;font-size:14px;font-weight:600;color:{COR_VITORIA};">Média: {med_vit:+.1f} sc/ha</p>'
                         ) if not np.isnan(max_vit) else '<p style="margin:2px 0;font-size:14px;">&nbsp;</p><p style="margin:0;font-size:14px;">&nbsp;</p>'
                k2.markdown(f'<div style="{card}border-top:3px solid {COR_VITORIA};">'
                            f'<p style="margin:0;font-size:15px;font-weight:700;color:#1A1A1A;">Vitórias</p>{sub_v}'
                            f'<p style="margin:6px 0;font-size:1.9rem;font-weight:700;color:{COR_VITORIA};">'
                            f'{n_vit} <span style="font-size:1rem;font-weight:600;">({pct_vit})</span></p></div>',
                            unsafe_allow_html=True)

                k3.markdown(f'<div style="{card}border-top:3px solid {COR_EMPATE_CARD};">'
                            f'<p style="margin:0;font-size:15px;font-weight:700;color:#1A1A1A;">Empates</p>'
                            f'<p style="margin:2px 0;font-size:14px;font-weight:600;color:{COR_EMPATE_CARD};">'
                            f'Entre ±{EMPATE_MARGEM:.0f} sc/ha</p>'
                            f'<p style="margin:0;font-size:14px;">&nbsp;</p>'
                            f'<p style="margin:6px 0;font-size:1.9rem;font-weight:700;color:{COR_EMPATE_CARD};">'
                            f'{n_emp} <span style="font-size:1rem;font-weight:600;">({pct_emp})</span></p></div>',
                            unsafe_allow_html=True)

                sub_d = (f'<p style="margin:2px 0;font-size:14px;font-weight:600;color:{COR_DERROTA};">Min: {min_der:+.1f} sc/ha</p>'
                         f'<p style="margin:0;font-size:14px;font-weight:600;color:{COR_DERROTA};">Média: {med_der:+.1f} sc/ha</p>'
                         ) if not np.isnan(min_der) else '<p style="margin:2px 0;font-size:14px;">&nbsp;</p><p style="margin:0;font-size:14px;">&nbsp;</p>'
                k4.markdown(f'<div style="{card}border-top:3px solid {COR_DERROTA};">'
                            f'<p style="margin:0;font-size:15px;font-weight:700;color:#1A1A1A;">Derrotas</p>{sub_d}'
                            f'<p style="margin:6px 0;font-size:1.9rem;font-weight:700;color:{COR_DERROTA};">'
                            f'{n_der} <span style="font-size:1rem;font-weight:600;">({pct_der})</span></p></div>',
                            unsafe_allow_html=True)

                st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

                col_donut, col_mapa = st.columns([1, 2])

                with col_donut:
                    fig_d = go_plt.Figure(go_plt.Pie(
                        labels=["Vitórias", "Empates", "Derrotas"], values=[n_vit, n_emp, n_der],
                        hole=0.55, marker_colors=[COR_VITORIA, COR_EMPATE, COR_DERROTA],
                        textinfo="label+percent", textposition="outside",
                        textfont=dict(size=12, family="Helvetica Neue, sans-serif", color="#111111"),
                        hovertemplate="%{label}: %{value} local(is) (%{percent})<extra></extra>",
                        sort=False, pull=[0.03, 0.03, 0.03],
                        domain=dict(x=[0.15, 0.85], y=[0.05, 0.90])))
                    fig_d.update_layout(
                        title=dict(text="Resultado geral do confronto",
                                   font=dict(size=13, color="#111111"),
                                   x=0.5, xanchor="center", y=0.99, yanchor="top"),
                        showlegend=False, height=420, margin=dict(t=80, b=20, l=60, r=60),
                        paper_bgcolor="#FFFFFF", font=dict(family="Helvetica Neue, sans-serif"))
                    st.plotly_chart(fig_d, use_container_width=True)

                with col_mapa:
                    _cols_coord = ["cod_fazenda", "latitude", "longitude", "nomeFazenda",
                                   "cidade_nome", "estado_sigla", "safra"]
                    if all(c in ta_filtrado.columns for c in ["latitude", "longitude"]):
                        df_coords = (ta_filtrado[_cols_coord].dropna(subset=["latitude", "longitude"])
                                     .sort_values("safra").drop_duplicates("cod_fazenda", keep="last"))
                        df_map = df_loc.merge(df_coords, on="cod_fazenda", how="left") \
                                       .dropna(subset=["latitude", "longitude"])
                    else:
                        df_map = pd.DataFrame()

                    folium = None
                    if df_map.empty:
                        st.info("Coordenadas não disponíveis para os locais deste confronto. "
                                "Atualize o pipeline e limpe o cache para habilitar o mapa.")
                    else:
                        try:
                            import folium
                            from streamlit_folium import st_folium
                        except ModuleNotFoundError:
                            folium = None
                            st.warning("Mapa indisponível: as bibliotecas de mapa não estão "
                                       "instaladas. Rode `pip install folium==0.20.0 "
                                       "streamlit-folium==0.27.1` e reinicie o app. "
                                       "O restante da análise funciona normalmente.")

                    if not df_map.empty and folium is not None:
                        def _cor_diff(d):
                            if d >= 10:  return "#27AE60"
                            if d >= 5:   return "#52C97A"
                            if d >= 2:   return "#A8E6BC"
                            if d >= -1:  return "#FFFF00"
                            if d >= -5:  return "#FF9900"
                            if d >= -10: return "#FF4400"
                            return "#FF0000"

                        _safras_u = sorted(df_map["safra"].dropna().unique().tolist())
                        _icones = ["circle", "star", "square", "diamond"]
                        _simb = {s: _icones[i % len(_icones)] for i, s in enumerate(_safras_u)}

                        def _svg_circulo(c):
                            return (f'<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24">'
                                    f'<circle cx="12" cy="12" r="10" fill="{c}" stroke="white" stroke-width="2"/></svg>')

                        def _svg_quadrado(c):
                            return (f'<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24">'
                                    f'<rect x="2" y="2" width="20" height="20" rx="3" fill="{c}" stroke="white" stroke-width="2"/></svg>')

                        def _svg_estrela(c):
                            return (f'<svg xmlns="http://www.w3.org/2000/svg" width="26" height="26" viewBox="0 0 26 26">'
                                    f'<polygon points="13,1 16.5,9.5 26,10.5 19.5,17 21.5,26 13,21.5 4.5,26 6.5,17 0,10.5 9.5,9.5" '
                                    f'fill="{c}" stroke="white" stroke-width="1.5"/></svg>')

                        _SVG_FN = {"circle": _svg_circulo, "square": _svg_quadrado, "star": _svg_estrela}

                        # marcadores na MESMA coordenada (ex.: mesma fazenda em duas safras, ou
                        # locais que herdaram a coordenada da cidade) ficariam empilhados e só o
                        # último apareceria — espalha em círculo pequeno para todos ficarem visíveis
                        import math
                        _RAIO_JITTER = 0.03          # ~3 km: separa visualmente sem tirar do município
                        df_map = df_map.reset_index(drop=True)
                        df_map["_dlat"] = 0.0
                        df_map["_dlon"] = 0.0
                        _chave_pos = (df_map["latitude"].round(4).astype(str) + "_"
                                      + df_map["longitude"].round(4).astype(str))
                        _n_empilhados = 0
                        for _, _idx in df_map.groupby(_chave_pos).groups.items():
                            _idx = list(_idx)
                            if len(_idx) > 1:
                                _n_empilhados += len(_idx)
                                for _k, _i in enumerate(_idx):
                                    _ang = 2 * math.pi * _k / len(_idx)
                                    df_map.at[_i, "_dlat"] = _RAIO_JITTER * math.cos(_ang)
                                    df_map.at[_i, "_dlon"] = _RAIO_JITTER * math.sin(_ang)

                        m = folium.Map(location=[df_map["latitude"].mean(), df_map["longitude"].mean()],
                                       zoom_start=5, tiles="OpenStreetMap")
                        for _, row in df_map.iterrows():
                            _cor = _cor_diff(row["diff_sc"])
                            _safra = str(row.get("safra", ""))
                            _icon = _simb.get(_safra, "circle")
                            _svg = _SVG_FN.get(_icon, _svg_circulo)(_cor)
                            _size = 26 if _icon == "star" else 24
                            _deslocado = (row["_dlat"] != 0) or (row["_dlon"] != 0)
                            _nota_pos = ("<br><i style='font-size:11px;color:#6B7280;'>ponto deslocado "
                                         "para não sobrepor outro local na mesma coordenada</i>"
                                         if _deslocado else "")
                            folium.Marker(
                                location=[row["latitude"] + row["_dlat"],
                                          row["longitude"] + row["_dlon"]],
                                popup=folium.Popup(
                                    f"<b>{row['nomeFazenda']}</b><br>{row['cidade_nome']} — {row['estado_sigla']}<br>"
                                    f"Safra: {_safra}<br><b>{p1_t2}:</b> {row['sc_ha_1']:.1f} sc/ha<br>"
                                    f"<b>{p2_t2}:</b> {row['sc_ha_2']:.1f} sc/ha<br>"
                                    f"<b>Diferença:</b> {row['diff_sc']:+.1f} sc/ha{_nota_pos}", max_width=280),
                                tooltip=f"{row['nomeFazenda']} · {_safra} · {row['diff_sc']:+.1f} sc/ha",
                                icon=folium.DivIcon(html=_svg, icon_size=(_size, _size),
                                                    icon_anchor=(_size // 2, _size // 2))).add_to(m)

                        _leg_cores = [("≥ +10 sc/ha", "#27AE60"), ("≥ +5 sc/ha", "#52C97A"),
                                      ("≥ +2 sc/ha", "#A8E6BC"), (f"±{EMPATE_MARGEM:.0f} sc/ha (empate)", "#FFFF00"),
                                      ("< −2 sc/ha", "#FF9900"), ("< −5 sc/ha", "#FF4400"),
                                      ("≤ −10 sc/ha", "#FF0000")]
                        _leg_html = "".join(
                            f'<div style="display:flex;align-items:center;gap:6px;margin:2px 0;">'
                            f'<div style="width:14px;height:14px;border-radius:50%;background:{c};'
                            f'border:1px solid #ccc;flex-shrink:0;"></div>'
                            f'<span style="font-size:11px;color:#374151;">{lbl}</span></div>'
                            for lbl, c in _leg_cores)
                        _SIMB_LEG = {"circle": "●", "square": "■", "star": "★", "diamond": "◆"}
                        _leg_safras = "".join(
                            f'<div style="display:flex;align-items:center;gap:6px;margin:2px 0;">'
                            f'<span style="font-size:16px;color:#374151;">{_SIMB_LEG.get(_simb.get(sf, "circle"), "●")}</span>'
                            f'<span style="font-size:11px;color:#374151;">{sf}</span></div>'
                            for sf in _safras_u)

                        st.markdown('<p style="font-size:13px;font-weight:600;color:#4A4A4A;margin:0 0 6px;">'
                                    'Locais por diferença</p>', unsafe_allow_html=True)
                        _cm, _cl = st.columns([4, 1])
                        with _cm:
                            st_folium(m, use_container_width=True, height=420, returned_objects=[])
                            if _n_empilhados:
                                st.caption(
                                    f"ℹ️ {_n_empilhados} locais dividem a mesma coordenada (mesma fazenda "
                                    "em safras diferentes, ou coordenada herdada da cidade). Eles foram "
                                    "afastados alguns quilômetros entre si só para ficarem visíveis — a "
                                    "posição real é a mesma, e o popup avisa quando o ponto está deslocado.")
                        with _cl:
                            st.markdown(
                                f"<div style='background:white;padding:10px 14px;border-radius:8px;"
                                f"border:1px solid #E5E7EB;box-shadow:0 1px 4px rgba(0,0,0,0.08);"
                                f"display:inline-block;'>"
                                f"<p style='font-size:12px;font-weight:700;color:#1A1A1A;margin:0 0 6px;'>"
                                f"Diferença (sc/ha)</p>{_leg_html}"
                                + (f"<p style='font-size:12px;font-weight:700;color:#1A1A1A;margin:8px 0 4px;'>"
                                   f"Safra</p>{_leg_safras}" if len(_safras_u) > 1 else "")
                                + "</div>", unsafe_allow_html=True)

                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

                cores_bar = df_loc["resultado"].map({"Vitória": COR_VITORIA, "Empate": COR_EMPATE,
                                                     "Derrota": COR_DERROTA}).tolist()
                fig_b = go_plt.Figure(go_plt.Bar(
                    x=df_loc["diff_sc"].round(1), y=df_loc["cod_fazenda"], orientation="h",
                    marker_color=cores_bar, text=df_loc["diff_sc"].round(1), textposition="outside",
                    textfont=dict(size=11, color="#111111"),
                    hovertemplate="<b>%{y}</b><br>Diferença: %{x:+.1f} sc/ha<extra></extra>"))
                fig_b.add_vline(x=0, line_color="#333333", line_width=2)
                fig_b.update_layout(
                    title=dict(text=f"Diferença de produtividade por local — {p1_t2} × {p2_t2}",
                               font=dict(size=13, color="#111111")),
                    xaxis=dict(title="Diferença (sc/ha)", tickfont=dict(size=11, color="#111111"),
                               zerolinecolor="#CCCCCC"),
                    yaxis=dict(title="Local", tickfont=dict(size=11, color="#111111")),
                    height=max(380, n_loc * 28 + 100), margin=dict(t=50, b=50, l=130, r=100),
                    plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
                    font=dict(family="Helvetica Neue, sans-serif", size=12, color="#111111"))
                st.plotly_chart(fig_b, use_container_width=True)

                st.markdown("**Dados por local**")
                df_exp = df_loc[["cod_fazenda", "sc_ha_1", "sc_ha_2", "diff_sc", "resultado"]].copy()
                df_exp.columns = ["Local", f"sc/ha — {p1_t2}", f"sc/ha — {p2_t2}",
                                  "Diferença (sc/ha)", "Resultado"]
                df_exp = df_exp.round(1)
                _est_dif = JsCode("""
                function(params) {
                    const v = Number(params.value);
                    if (isNaN(v)) return {'textAlign':'center'};
                    if (v > 0) return {'color':'#15803D','fontWeight':'800','textAlign':'center'};
                    if (v < 0) return {'color':'#B91C1C','fontWeight':'800','textAlign':'center'};
                    return {'textAlign':'center'};
                }
                """)
                _est_res = JsCode("""
                function(params) {
                    const v = params.value;
                    if (v === 'Vitória') return {'backgroundColor':'rgba(39,174,96,0.18)','fontWeight':'700','textAlign':'center'};
                    if (v === 'Derrota') return {'backgroundColor':'rgba(255,0,0,0.15)','fontWeight':'700','textAlign':'center'};
                    if (v === 'Empate')  return {'backgroundColor':'rgba(255,255,0,0.25)','fontWeight':'700','textAlign':'center'};
                    return {'textAlign':'center'};
                }
                """)
                ag_table_h2h(df_exp, height=min(520, 40 + 32 * min(len(df_exp), 14) + 20),
                             estilos_col={"Diferença (sc/ha)": _est_dif, "Resultado": _est_res})
                st.download_button("⬇️ Exportar Excel", data=to_excel(df_exp),
                                   file_name=f"h2h_local_{p1_t2}_vs_{p2_t2}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="dl_t2")


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Desvios por Ambiente (a construir)
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    secao_titulo(
        "HEAD-TO-HEAD · DESVIOS",
        "Como cada híbrido se comporta em ambientes favoráveis vs. adversos?",
        "Desvio em relação à média do local — pontos acima de zero indicam que o híbrido superou a "
        "média daquele ambiente.",
    )

    st.markdown(
        '<div style="background:#FFF8E1;border:1px solid #F5D76E;border-left:5px solid #D4A800;'
        'border-radius:8px;padding:14px 18px;margin:6px 0 14px;">'
        '<p style="margin:0;font-size:15.5px;line-height:1.6;color:#4A3B00;'
        'font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">'
        '<span style="font-size:19px;vertical-align:-1px;margin-right:8px;">⚠️</span>'
        'A média do local considera <b>todos os híbridos avaliados ali</b>, inclusive os fora do '
        'filtro — assim o desvio de cada local <b>não muda</b> conforme a seleção. Os filtros '
        'definem <b>quais locais</b> entram, e a reta e o R² <b>se recalculam '
        'sozinhos</b>; o botão só é necessário ao <b>trocar os híbridos</b>.<br><br>'
        '<b>Cuidado ao comparar recortes:</b> uma inclinação calculada com 5 locais é muito menos '
        'confiável que a mesma com 30. Confira o nº de locais e o R² antes de concluir.'
        '</p></div>',
        unsafe_allow_html=True,
    )

    if df_p1.empty:
        st.warning("Dados insuficientes com os filtros atuais.")
    else:
        hibridos_p1_t3 = sorted(df_p1["dePara"].dropna().unique())
        col_a3, col_b3, col_c3 = st.columns([2, 2, 1])

        with col_a3:
            p1_t3 = st.selectbox("Produto 1 (STINE / EXP / DP2)", hibridos_p1_t3, key="p1_t3")

        _loc_p1_t3 = set(df_p1[df_p1["dePara"] == p1_t3]["cod_fazenda"].dropna().unique())
        adv_disp_t3 = sorted(ta_filtrado[(ta_filtrado["cod_fazenda"].isin(_loc_p1_t3))
                                         & (ta_filtrado["dePara"] != p1_t3)]["dePara"].dropna().unique())

        with col_b3:
            if adv_disp_t3:
                p2_t3 = st.selectbox("Produto 2 (adversário)", adv_disp_t3, key="p2_t3")
            else:
                st.warning("Nenhum adversário com locais em comum.")
                p2_t3 = None

        with col_c3:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            btn_t3 = st.button("Rodar análise", type="primary", key="btn_t3", use_container_width=True)

        key_t3 = f"res_t3_raw__{p1_t3}__{p2_t3}"

        if btn_t3 and p2_t3:
            with st.spinner("Calculando desvios..."):
                # média do local sobre TODOS os híbridos avaliados ali (não muda com o filtro)
                _base = ta_raw[ta_raw["sc_ha"] > 0]
                _media_loc = _base.groupby("cod_fazenda")["sc_ha"].mean().rename("media_local")

                _l1 = set(_base[_base["dePara"] == p1_t3]["cod_fazenda"].dropna())
                _l2 = set(_base[_base["dePara"] == p2_t3]["cod_fazenda"].dropna())
                _comuns = _l1 & _l2

                def _desvios(hib):
                    return (_base[(_base["dePara"] == hib) & (_base["cod_fazenda"].isin(_comuns))]
                            [["cod_fazenda", "sc_ha"]].dropna()
                            .groupby("cod_fazenda", as_index=False)["sc_ha"].mean()
                            .join(_media_loc, on="cod_fazenda")
                            .assign(desvio=lambda d: d["sc_ha"] - d["media_local"], hibrido=hib))

                st.session_state[key_t3] = (_desvios(p1_t3), _desvios(p2_t3))

        if key_t3 not in st.session_state or not p2_t3:
            st.info("Selecione os dois híbridos e clique em **Rodar análise** para calcular.")
        else:
            _df1, _df2 = st.session_state[key_t3]
            # filtros vigentes agem só na exibição
            _loc_ativos_t3 = set(ta_filtrado["cod_fazenda"].dropna().unique())
            _df1 = _df1[_df1["cod_fazenda"].isin(_loc_ativos_t3)].copy()
            _df2 = _df2[_df2["cod_fazenda"].isin(_loc_ativos_t3)].copy()

            if _df1.empty or _df2.empty:
                st.info("Nenhum local compartilhado dentro do recorte filtrado.")
            else:
                with st.popover("ℹ️ Como interpretar · Desvios por Ambiente", use_container_width=False):
                    st.markdown("""
**📌 O que este gráfico mostra**

Cada ponto é o desempenho de um híbrido **em relação à média de todos os híbridos** daquele local.

- **Eixo X** → produtividade média do local (o quanto o ambiente foi favorável).
- **Eixo Y** → desvio do híbrido: `sc/ha do híbrido − média do local`.

---

**📐 Como ler**

- **Ponto acima de zero** → produziu acima da média naquele local.
- **Ponto abaixo de zero** → ficou abaixo da média.
- **Reta subindo (b > 0)** → ganha vantagem em ambientes melhores: é **responsivo**.
- **Reta descendo (b < 0)** → perde vantagem quando o ambiente melhora, mas se segura nos adversos.
- **Reta horizontal (b ≈ 0)** → comportamento independente do ambiente.

O **R²** indica o quanto o ambiente explica o desvio: alto significa comportamento previsível ao
longo dos ambientes; baixo, que o híbrido reage de forma irregular.

---

**🔍 Diferença para a aba Por Local**

Lá as barras mostram a diferença **entre os dois híbridos**. Aqui cada um é medido contra o
**conjunto** de todos os híbridos do local — então dá para ver se ambos vão bem, ambos vão mal, ou
se um sobe enquanto o outro desce conforme o ambiente muda.

> O **p-valor** vem de um teste t pareado entre os dois híbridos nos locais em comum: abaixo de
0,05 indica que a diferença média entre eles dificilmente é acaso.
""")

                _COR_MAP = {"CHECK": "#E67E22", "STINE": "#2976B6", "EXP": "#009900", "DP2": "#7AAF6A"}
                _st1_t3 = (df_p1[df_p1["dePara"] == p1_t3]["status_material"].iloc[0]
                           if not df_p1[df_p1["dePara"] == p1_t3].empty else "")
                _st2_t3 = (ta_filtrado[ta_filtrado["dePara"] == p2_t3]["status_material"].iloc[0]
                           if not ta_filtrado[ta_filtrado["dePara"] == p2_t3].empty else "")
                _cor1 = _COR_MAP.get(_st1_t3, "#2976B6")
                _cor2 = _COR_MAP.get(_st2_t3, "#E67E22")
                if _cor1 == _cor2:                      # mesmo status: diferencia o adversário
                    _cor2 = "#8E44AD"

                _sc1 = _df1.set_index("cod_fazenda")["sc_ha"]
                _sc2 = _df2.set_index("cod_fazenda")["sc_ha"]
                _idx = _sc1.index.intersection(_sc2.index)
                _dif = _sc1[_idx] - _sc2[_idx]
                _dif_media = float(_dif.mean()) if len(_dif) else 0.0
                _n_vit_t3 = int((_dif > EMPATE_MARGEM).sum())
                _pct_vit_t3 = round(_n_vit_t3 / len(_idx) * 100, 1) if len(_idx) else 0

                _pval_str = ""
                try:
                    from scipy import stats as _stats
                    if len(_idx) >= 2:
                        _, _pv = _stats.ttest_rel(_sc1[_idx], _sc2[_idx])
                        _pval_str = f"p={_pv:.3f}"
                except Exception:
                    _pval_str = ""

                _sinal = "+" if _dif_media >= 0 else ""
                _sub = (f"{_pct_vit_t3:.0f}% de vitórias · diferença média {_sinal}{_dif_media:.1f} sc/ha"
                        + (f" · {_pval_str}" if _pval_str else "")
                        + f" · {len(_idx)} locais")

                st.markdown(
                    f'<p style="font-size:22px;font-weight:700;color:#1A1A1A;margin:0;">'
                    f'<span style="color:{_cor1};">{p1_t3}</span>'
                    f' <span style="font-size:16px;font-weight:400;color:#6B7280;">vs</span> '
                    f'<span style="color:{_cor2};">{p2_t3}</span></p>'
                    f'<p style="font-size:15px;color:#374151;margin:4px 0 14px;">{_sub}</p>',
                    unsafe_allow_html=True)

                _meta_loc = (ta_filtrado[["cod_fazenda", "nomeFazenda", "cidade_nome", "estado_sigla"]]
                             .drop_duplicates("cod_fazenda").set_index("cod_fazenda"))

                fig_dev = go_plt.Figure()
                _retas = []
                for _d, _hib, _cor in [(_df1, p1_t3, _cor1), (_df2, p2_t3, _cor2)]:
                    _d = _d.join(_meta_loc, on="cod_fazenda", how="left")
                    _x = _d["media_local"].values
                    _y = _d["desvio"].values
                    fig_dev.add_trace(go_plt.Scatter(
                        x=_x, y=_y, mode="markers", name=_hib,
                        marker=dict(color=_cor, size=10, opacity=0.80,
                                    line=dict(color="#FFFFFF", width=1.2)),
                        customdata=_d[["cod_fazenda", "nomeFazenda", "cidade_nome", "estado_sigla"]].values,
                        hovertemplate=(f"<b>{_hib}</b><br><b>%{{customdata[0]}}</b> — %{{customdata[1]}}<br>"
                                       "%{customdata[2]}, %{customdata[3]}<br>"
                                       "Média local: %{x:.1f} sc/ha<br>"
                                       "Desvio: %{y:+.1f} sc/ha<extra></extra>")))
                    if len(_x) >= 2:
                        try:
                            _Xr = np.column_stack([np.ones(len(_x)), _x])
                            _beta, _, _, _ = np.linalg.lstsq(_Xr, _y, rcond=None)
                            _xl = np.linspace(_x.min(), _x.max(), 100)
                            _yl = _beta[0] + _beta[1] * _xl
                            _ss_res = np.sum((_y - _Xr @ _beta) ** 2)
                            _ss_tot = np.sum((_y - _y.mean()) ** 2)
                            _r2 = 1 - _ss_res / _ss_tot if _ss_tot > 0 else np.nan
                            fig_dev.add_trace(go_plt.Scatter(
                                x=_xl, y=_yl, mode="lines", line=dict(color=_cor, width=2.5),
                                showlegend=False, hoverinfo="skip"))
                            _retas.append({"hib": _hib, "cor": _cor, "x_end": _xl[-1],
                                           "y_end": _yl[-1], "slope": _beta[1], "r2": _r2})
                        except Exception:
                            pass

                if len(_retas) == 2 and abs(_retas[0]["y_end"] - _retas[1]["y_end"]) < 4:
                    _ay = [-35, 20] if _retas[0]["y_end"] >= _retas[1]["y_end"] else [20, -35]
                else:
                    _ay = [10] * len(_retas)
                for _i, _r in enumerate(_retas):
                    fig_dev.add_annotation(
                        x=_r["x_end"], y=_r["y_end"],
                        text=f"<b>{_r['hib']}</b><br>b={_r['slope']:+.3f} · R²={_r['r2']:.2f}",
                        showarrow=True, arrowhead=0, arrowcolor=_r["cor"], arrowwidth=1.5,
                        ax=25, ay=_ay[_i], xanchor="left",
                        font=dict(size=12, color=_r["cor"], weight="bold"),
                        bgcolor="rgba(255,255,255,0.85)", bordercolor=_r["cor"],
                        borderwidth=1, borderpad=3)

                fig_dev.add_hline(y=0, line=dict(color="#444444", width=1.5, dash="dot"))
                fig_dev.update_layout(
                    height=520, plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
                    font=dict(family="Helvetica Neue, sans-serif", color="#1A1A1A"),
                    showlegend=True,
                    legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0,
                                font=dict(size=13, color="#1A1A1A", weight="bold")),
                    xaxis=dict(title=dict(text="<b>Produtividade média do local (sc/ha)</b>",
                                          font=dict(size=14, color="#1A1A1A", weight="bold")),
                               tickfont=dict(size=12, color="#1A1A1A", weight="bold"),
                               showgrid=True, gridcolor="#E5E5E5", zeroline=False),
                    yaxis=dict(title=dict(text="<b>Desvio em relação à média do local (sc/ha)</b>",
                                          font=dict(size=14, color="#1A1A1A", weight="bold")),
                               tickfont=dict(size=12, color="#1A1A1A", weight="bold"),
                               showgrid=True, gridcolor="#E5E5E5", zeroline=False),
                    margin=dict(t=60, b=60, l=80, r=220))

                st.plotly_chart(fig_dev, use_container_width=True)
                st.caption("ℹ️ Desvio = sc/ha do híbrido − média de todos os híbridos no local · "
                           "b > 0 = ganha vantagem em ambientes favoráveis · b < 0 = perde vantagem "
                           "em ambientes melhores · b ≈ 0 = comportamento independente do ambiente.")

                st.markdown("**Desvios por local**")
                _tab3 = (_df1[["cod_fazenda", "media_local", "sc_ha", "desvio"]]
                         .merge(_df2[["cod_fazenda", "sc_ha", "desvio"]], on="cod_fazenda",
                                suffixes=("_1", "_2"))
                         .sort_values("media_local"))
                _tab3.columns = ["Local", "Média do local (sc/ha)", f"sc/ha — {p1_t3}",
                                 f"Desvio — {p1_t3}", f"sc/ha — {p2_t3}", f"Desvio — {p2_t3}"]
                _tab3 = _tab3.round(1)
                def _js_sinal3():
                    return JsCode("""
                    function(params) {
                        const v = Number(params.value);
                        if (isNaN(v)) return {'textAlign':'center'};
                        if (v > 0) return {'color':'#15803D','fontWeight':'800','textAlign':'center'};
                        if (v < 0) return {'color':'#B91C1C','fontWeight':'800','textAlign':'center'};
                        return {'textAlign':'center'};
                    }
                    """)
                ag_table_h2h(_tab3, height=min(520, 40 + 32 * min(len(_tab3), 14) + 20),
                             estilos_col={f"Desvio — {p1_t3}": _js_sinal3(),
                                          f"Desvio — {p2_t3}": _js_sinal3()})
                st.download_button("⬇️ Exportar Excel", data=to_excel(_tab3),
                                   file_name=f"h2h_desvios_{p1_t3}_vs_{p2_t3}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="dl_t3")




rodape()
