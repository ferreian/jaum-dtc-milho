"""
pages/0_Auditoria_Calculo.py — Auditoria dos cálculos do gold

Página de conferência: escolhe safra + fazenda + material, mostra as variáveis CRUAS que entram
no cálculo, refaz a conta passo a passo AO VIVO e compara com o que o pipeline (gold) produziu.
Se bater, o cálculo está correto.

Safra:
  - 2025: conferência com dado real + SIMULADOR (input manual das notas);
  - 2024: só a conferência (histórico; não faz sentido simular). Nesta safra a av1 não tem a
    nota geral do técnico (qualidade_plot_inicial) — a página lida com a ausência.

Começa pela av1; as demais avaliações entram depois.
"""
import numpy as np
import pandas as pd
import streamlit as st

from utils.theme import aplicar_tema, page_header, secao_titulo, rodape
from utils.loader import carregar_2024, carregar_2025, aplicar_mestre, _mapas_mestre

# Réguas e nomes canônicos vêm do PRÓPRIO pipeline (fonte única — iguais nas duas safras)
from pipeline_milho_2025 import (
    NOTAS_AV1, RENAME_AV1,
    DOENCAS_AV2, RENAME_AV2, _classificar_doenca,
    COLS_ALT_PLANTA, COLS_ALT_ESPIGA, FLOR_MIN, FLOR_MAX, padronizar_altura_cm,
    UMID_PADRAO, UMID_MIN, UMID_MAX, PROD_TETO,
    COLS_ESTANDE_8P, COLS_ESTANDE_5SUB,
    N_SUBAMOSTRAS_2025, N_SUBAMOSTRAS_2024, MAX_SLOTS_SUBAMOSTRA,
    MAPA_PERDAS, FENOMENOS_AV4,
)

st.set_page_config(
    page_title="Auditoria de Cálculo · JAUM DTC",
    page_icon="🧮",
    layout="wide",
    initial_sidebar_state="expanded",
)

aplicar_tema()
page_header("Auditoria de Cálculo",
            "Confira, plot a plot, se o gold bate com a conta a partir das variáveis cruas.",
            imagem="App development-bro.png")

st.markdown("""
<style>
.audit-var  { background:#F9FAFB; border:1px solid #E5E7EB; border-radius:8px; padding:8px 12px; }
.audit-ok   { color:#1E8449; font-weight:700; }
.audit-bad  { color:#C0392B; font-weight:700; }
.audit-step { font-size:13px; color:#374151; line-height:1.9; }
.audit-mono { font-family:ui-monospace,SFMono-Regular,Menlo,monospace; font-size:13px; }
</style>
""", unsafe_allow_html=True)

LABEL_NOTA = {
    "nota_uniformidade": "Uniformidade de emergência",
    "nota_densidade":    "Densidade de plantas",
    "nota_vigor":        "Vigor das plantas",
    "nota_daninhas":     "Presença de daninhas",
    "nota_pragas":       "Presença de pragas",
    "nota_doencas":      "Presença de doenças",
    "nota_homogeneidade": "Homogeneidade de crescimento",
    "nota_solo":         "Estado geral do solo",
}
CRU_POR_CANON = {v: k for k, v in RENAME_AV1.items()}


def _ou_travessao(v):
    """Valor legível ou '—'. Trata None, NaN (que é truthy!) e string vazia — 'x or …' sozinho
    não pega NaN, então checamos com pd.isna antes."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and not v.strip()):
        return "—"
    return v


# ── Régua da av1 (a MESMA lógica do _gold_av1, isolada para reuso) ────────────
def calcular_media_av1(notas_vals):
    """Refaz media_categorias a partir das 8 notas, igual ao _gold_av1.
    Retorna (media, passos, tem_escala_errada, n_validas). 0 = não avaliado; nota ≥6 → NaN no plot."""
    notas = pd.to_numeric(pd.Series(list(notas_vals)), errors="coerce")
    tem_escala_errada = bool((notas >= 6).any())
    validas = notas.where(notas >= 1)          # 0 → NaN (não avaliado)
    n_validas = int(validas.notna().sum())
    soma = float(validas.sum()) if n_validas else np.nan
    if tem_escala_errada or n_validas == 0:
        media = np.nan
    else:
        media = round(soma / n_validas, 1)
    return media, tem_escala_errada, n_validas, soma


# ── av2: rótulos das doenças e régua de incidência/classe (a MESMA do _gold_av2) ──
LABEL_DOENCA = {
    "nota_turcicum":         "Mancha de Turcicum",
    "nota_cercospora":       "Mancha de Cercospora",
    "nota_mancha_branca":    "Mancha branca",
    "nota_bipolaris":        "Mancha de Bipolaris",
    "nota_ferrugem_tropical": "Ferrugem tropical",
    "nota_enfezamento":      "Enfezamento",
}
CRU_POR_CANON_AV2 = {v: k for k, v in RENAME_AV2.items()}


def avaliar_doenca(nota_bruta):
    """Refaz incidência e classe de UMA doença, igual ao _gold_av2.
    Retorna (nota, incidencia, classe): 0 → não avaliado (NaN); incidência 1=presente (nota 1-5),
    0=ausente (6-9); classe AS/S/MT/T/R. Escala 1-9 onde 9 = mais resistente."""
    n = pd.to_numeric(pd.Series([nota_bruta]), errors="coerce").where(lambda x: x > 0).iloc[0]
    if pd.isna(n):
        return (np.nan, pd.NA, np.nan)
    incidencia = 1 if 1 <= n <= 5 else 0
    return (n, incidencia, _classificar_doenca(n))


# Paletas de cor para a av2 (fundo suave + texto escuro, legível).
# Classe: do vermelho (AS, pior) ao verde (R, melhor).
COR_CLASSE = {
    "AS": "#F8D7DA",   # vermelho claro
    "S":  "#FCE3D2",   # laranja-avermelhado claro
    "MT": "#FFF3CD",   # amarelo claro
    "T":  "#E2F0D9",   # verde claro
    "R":  "#C6EFCE",   # verde
}
# Incidência: cinza (não avaliada), verde (ausente), vermelho (presente).
COR_INCIDENCIA = {
    "não avaliada": "#EDEFF2",
    "ausente":      "#C6EFCE",
    "presente":     "#F8D7DA",
}


def _estilo_classe(v):
    cor = COR_CLASSE.get(str(v).strip())
    return f"background-color: {cor}; color: #1A1A1A; font-weight: 600;" if cor else ""


def _estilo_incidencia(v):
    cor = COR_INCIDENCIA.get(str(v).strip())
    return f"background-color: {cor}; color: #1A1A1A;" if cor else ""


def _styler_map(styler, func, subset):
    """Aplica estilo célula a célula, compatível com pandas antigo (.applymap) e novo (.map).
    O .applymap do Styler foi renomeado para .map na 2.1 e removido na 3.0."""
    if hasattr(styler, "map"):
        return styler.map(func, subset=subset)
    return styler.applymap(func, subset=subset)


# ── Seletores: safra + avaliação ──────────────────────────────────────────────
secao_titulo("Escopo", "O que auditar",
             "Escolha a safra e a avaliação. Começamos pela av1 (qualidade inicial) e av2 (sanidade).")

AVALIACOES = {
    "av1 — Qualidade inicial": "av1",
    "av2 — Sanidade (doenças)": "av2",
    "av3 — Caracterização (altura e florescimento)": "av3",
    "av4 — Produtividade e colheita": "av4",
}

col_s, col_a = st.columns(2)
with col_s:
    safra = st.radio("Safra", ["2025 (25/26)", "2024 (24/25)"], horizontal=True, key="audit_safra")
with col_a:
    aval_label = st.radio("Avaliação", list(AVALIACOES.keys()), horizontal=True, key="audit_aval")
aval = AVALIACOES[aval_label]

is_2025 = safra.startswith("2025")

with st.spinner(f"Carregando dados de {'2025' if is_2025 else '2024'}..."):
    dados = carregar_2025() if is_2025 else carregar_2024()

if not dados.get("ok"):
    st.error(f"Não foi possível carregar a safra: {dados.get('erro', 'erro desconhecido')}")
    if dados.get("traceback"):
        with st.expander("Detalhe técnico do erro (traceback)"):
            st.code(dados["traceback"], language="text")
    st.stop()

gold_df = dados.get(f"{aval}_gold")
# nome/status canônicos do mestre — o mesmo híbrido tem um nome só entre safras
# (o nome cru da safra continua na coluna `nome`, e o canônico da safra em `dePara_original`)
if isinstance(gold_df, pd.DataFrame) and not gold_df.empty:
    _mc, _ms = _mapas_mestre()
    gold_df = aplicar_mestre(gold_df, _mc, _ms)
if gold_df is None or gold_df.empty:
    st.warning(f"Não há dados de {aval} nesta safra.")
    st.stop()

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# PARTE 1 — CONFERÊNCIA COM DADO REAL
# ══════════════════════════════════════════════════════════════════════════════
secao_titulo("Conferência", "Plot real do banco",
             "Escolha o responsável, a fazenda e o material; a página refaz a conta a partir das "
             "notas cruas e compara com o gold.")

# ── Seletores de responsável + tipo de ensaio (filtram os plots) ──────────────
TODOS = "(todos os responsáveis)"
TODOS_TT = "(todos)"
col_r, col_t = st.columns(2)
with col_r:
    responsaveis = sorted(gold_df["nomeResponsavel"].dropna().unique().tolist()) \
        if "nomeResponsavel" in gold_df.columns else []
    resp_sel = st.selectbox("Responsável", [TODOS] + responsaveis, key="audit_resp")
with col_t:
    tipos = sorted(gold_df["tipoTeste"].dropna().unique().tolist()) \
        if "tipoTeste" in gold_df.columns else []
    tt_sel = st.selectbox("Tipo de ensaio", [TODOS_TT] + tipos, key="audit_tt")

# base filtrada por responsável E tipo de ensaio (cada filtro é opcional)
base = gold_df.copy()
if resp_sel != TODOS:
    base = base[base["nomeResponsavel"] == resp_sel]
if tt_sel != TODOS_TT:
    base = base[base["tipoTeste"] == tt_sel]

col_f, col_m = st.columns(2)
with col_f:
    # mapa cod_fazenda -> "COD — Nome da Fazenda (Cidade/UF)" para o seletor ficar legível
    fz = base.dropna(subset=["cod_fazenda"]).drop_duplicates("cod_fazenda")
    def _rotulo_fazenda(cod):
        r = fz[fz["cod_fazenda"] == cod]
        if r.empty:
            return cod
        r = r.iloc[0]
        nome = _ou_travessao(r.get("nomeFazenda"))
        cid = r.get("cidade_nome"); uf = r.get("estado_sigla")
        local = f" ({cid}/{uf})" if pd.notna(cid) and pd.notna(uf) else ""
        return f"{cod} — {nome}{local}"
    fazendas = sorted(fz["cod_fazenda"].tolist())
    if not fazendas:
        st.info("Nenhuma fazenda com av1 para esse filtro.")
        st.stop()
    faz_sel = st.selectbox("Fazenda", fazendas, format_func=_rotulo_fazenda, key="audit_faz")
with col_m:
    # cada opção é UM plot: "dePara · trat. N". O valor é o índice da linha (único),
    # então material + tratamento ficam num seletor só.
    plots_faz = base[base["cod_fazenda"] == faz_sel].copy()
    if "indexTratamento" in plots_faz.columns:
        plots_faz = plots_faz[plots_faz["dePara"].notna()].sort_values(
            ["indexTratamento", "dePara"], ascending=[True, True])
    else:
        plots_faz = plots_faz[plots_faz["dePara"].notna()].sort_values("dePara")

    def _rotulo_plot(ix):
        r = plots_faz.loc[ix]
        dep = r.get("dePara")
        # base: "# N · Material"
        if "indexTratamento" in plots_faz.columns and pd.notna(r.get("indexTratamento")):
            it = r.get("indexTratamento")
            it = int(it) if float(it).is_integer() else it
            base = f"# {it} · {dep}"
        else:
            base = f"{dep}"
        # na Densidade, o mesmo material aparece em várias populações — mostra a pop-alvo
        # para localizar o plot certo (é como o técnico pensa o ensaio de densidade)
        tipo = str(r.get("tipoTeste") or "")
        pop = r.get("pop_tratamento")
        if tipo.lower().startswith("dens") and pd.notna(pop):
            try:
                pop_txt = f"{int(float(pop)):,}".replace(",", ".")
                base = f"{base} · {pop_txt} pl/ha"
            except (ValueError, TypeError):
                pass
        return base

    if plots_faz.empty:
        st.info("Nenhum plot com material nessa fazenda.")
        st.stop()
    # reset_index para posições sequenciais (0..N-1): evita KeyError quando os filtros mudam
    # e o valor guardado no estado do selectbox aponta para um índice que não existe mais.
    plots_faz = plots_faz.reset_index(drop=True)
    opcoes = plots_faz.index.tolist()
    plot_sel = st.selectbox("Híbrido", opcoes, format_func=_rotulo_plot, key="audit_plot")

# se o estado guardou uma posição que saiu do intervalo (filtro mudou), cai na primeira
if plot_sel not in plots_faz.index:
    plot_sel = plots_faz.index[0]
linha = plots_faz.loc[plot_sel]
mat_sel = linha.get("dePara")

# ── cabeçalho de contexto do plot (nome do responsável, fazenda, tipo de ensaio) ──
resp = _ou_travessao(linha.get("nomeResponsavel"))
nome_faz = _ou_travessao(linha.get("nomeFazenda"))
cidade = _ou_travessao(linha.get("cidade_nome"))
uf = linha.get("estado_sigla") if pd.notna(linha.get("estado_sigla")) else ""
tipo_teste = _ou_travessao(linha.get("tipoTeste"))
cor_tt = "#1E8449" if str(tipo_teste).lower() == "faixa" else "#B9770E"
_it = linha.get("indexTratamento")
_it_txt = (str(int(_it)) if pd.notna(_it) and float(_it).is_integer() else str(_it)) if pd.notna(_it) else "—"

# chip de população-alvo: só na Densidade (onde a população define o plot)
_pop = linha.get("pop_tratamento")
_chip_pop = ""
if str(tipo_teste).lower().startswith("dens") and pd.notna(_pop):
    try:
        _pop_txt = f"{int(float(_pop)):,}".replace(",", ".")
        _chip_pop = (
            '<div><span style="font-size:11px;color:#6B7280;text-transform:uppercase;">População-alvo</span><br>'
            f'<span style="font-size:14px;font-weight:600;color:#1A1A1A;">{_pop_txt} pl/ha</span></div>')
    except (ValueError, TypeError):
        pass

st.markdown(f"""
<div style="display:flex;flex-wrap:wrap;gap:18px;align-items:center;background:#F9FAFB;
            border:1px solid #E5E7EB;border-radius:10px;padding:10px 16px;margin:4px 0 12px;">
  <div><span style="font-size:11px;color:#6B7280;text-transform:uppercase;">Responsável</span><br>
       <span style="font-size:14px;font-weight:600;color:#1A1A1A;">{resp}</span></div>
  <div><span style="font-size:11px;color:#6B7280;text-transform:uppercase;">Fazenda</span><br>
       <span style="font-size:14px;font-weight:600;color:#1A1A1A;">{nome_faz}</span></div>
  <div><span style="font-size:11px;color:#6B7280;text-transform:uppercase;">Local</span><br>
       <span style="font-size:14px;color:#374151;">{cidade}/{uf}</span></div>
  <div><span style="font-size:11px;color:#6B7280;text-transform:uppercase;">Híbrido</span><br>
       <span style="font-size:14px;font-weight:600;color:#1A1A1A;">{mat_sel}</span></div>
  <div><span style="font-size:11px;color:#6B7280;text-transform:uppercase;">Tratamento</span><br>
       <span style="font-size:14px;font-weight:600;color:#1A1A1A;">{_it_txt}</span></div>
  {_chip_pop}
  <div><span style="font-size:11px;color:#6B7280;text-transform:uppercase;">Ensaio</span><br>
       <span style="font-size:14px;font-weight:700;color:{cor_tt};">{tipo_teste}</span></div>
</div>
""", unsafe_allow_html=True)

if aval == "av1":
    # 1) variáveis cruas
    st.markdown("**1. Variáveis cruas (as 8 notas do plot)**")
    st.dataframe(pd.DataFrame({
        "Categoria": [LABEL_NOTA[c] for c in NOTAS_AV1],
        "Campo no banco": [CRU_POR_CANON.get(c, "?") for c in NOTAS_AV1],
        "Nota": [linha.get(c) for c in NOTAS_AV1],
    }), use_container_width=True, hide_index=True)

    # 2) a conta, explicada para leigos: a regra em palavras + o exemplo com os números do plot
    media_calc, escala_errada, n_validas, soma = calcular_media_av1([linha.get(c) for c in NOTAS_AV1])
    notas_num = pd.to_numeric(pd.Series([linha.get(c) for c in NOTAS_AV1]), errors="coerce")
    notas_validas = [int(x) if float(x).is_integer() else x for x in notas_num.where(notas_num >= 1).dropna().tolist()]
    qtd_zeros = int((notas_num == 0).sum())

    st.markdown("**2. As regras do cálculo**")
    st.markdown(
        "<div class='audit-step'>Na av1 o técnico dá 8 notas (de 1 a 5) para o plot. "
        "A <b>média</b> é a soma dessas notas dividida por quantas foram dadas. "
        "Nota <b>0</b> significa “não avaliei”, então ela não entra na conta. "
        "E se aparecer alguma nota <b>6 ou maior</b> (que não existe na escala de 1 a 5), "
        "o plot é marcado como sem média, para não misturar escalas.</div>",
        unsafe_allow_html=True)

    if escala_errada:
        exemplo = ("Neste plot há uma nota 6 ou maior, fora da escala de 1 a 5. "
                   "Por isso a média fica em branco (sem valor).")
        if not is_2025:
            exemplo += " Isso é comum em 2024, quando algumas fazendas usaram a escala antiga de 1 a 9."
    elif n_validas == 0:
        exemplo = "Neste plot todas as 8 notas são 0 (nada foi avaliado), então não há média."
    else:
        zeros_txt = (f" Tirando {qtd_zeros} nota(s) igual a 0, sobram {n_validas}." if qtd_zeros else "")
        exemplo = (f"Neste plot as notas dadas foram {notas_validas}.{zeros_txt} "
                   f"Somando: {soma:.0f}. Dividindo por {n_validas}: "
                   f"{soma:.0f} ÷ {n_validas} = <b>{media_calc}</b>.")
    st.markdown(f"<div class='audit-var audit-step'><b>Neste caso:</b> {exemplo}</div>",
                unsafe_allow_html=True)

    # 3) conferência
    media_gold = linha.get("media_categorias")
    igual = (pd.isna(media_calc) and pd.isna(media_gold)) or (
        pd.notna(media_calc) and pd.notna(media_gold) and abs(float(media_calc) - float(media_gold)) < 0.05)

    st.markdown("**3. Resultado da conferência**")
    c1, c2, c3 = st.columns(3)
    c1.metric("Pela conta acima", "sem média" if pd.isna(media_calc) else f"{media_calc}")
    c2.metric("Valor no sistema", "sem média" if pd.isna(media_gold) else f"{media_gold}")
    with c3:
        if igual:
            st.markdown("<div class='audit-var'><span class='audit-ok'>✓ CONFERE</span><br>"
                        "<span style='font-size:12px;color:#6B7280'>os dois valores são iguais</span></div>",
                        unsafe_allow_html=True)
        else:
            st.markdown("<div class='audit-var'><span class='audit-bad'>✗ DIVERGE</span><br>"
                        "<span style='font-size:12px;color:#6B7280'>os valores estão diferentes</span></div>",
                        unsafe_allow_html=True)

    if not is_2025 and ("qualidade_plot_inicial" not in linha.index or pd.isna(linha.get("qualidade_plot_inicial"))):
        st.caption("Obs.: em 2024 a nota geral do técnico (qualidade_plot_inicial) não era coletada — "
                   "a média das 8 categorias é o único resumo da av1 nesta safra.")

elif aval == "av2":
    # doenças presentes nesta safra (2024 tem tombamentoVerde/graosArdidos vivos; 2025 não)
    doencas_safra = [d for d in DOENCAS_AV2 if d in gold_df.columns]

    # 1) variáveis cruas + as duas derivações, doença a doença
    st.markdown("**1. Notas de sanidade e o que a regra deriva de cada uma**")
    linhas_tab = []
    for d in doencas_safra:
        nota, inc, classe = avaliar_doenca(linha.get(d))
        linhas_tab.append({
            "Doença": LABEL_DOENCA.get(d, d),
            "Campo no banco": CRU_POR_CANON_AV2.get(d, "?"),
            "Nota (1–9)": "—" if pd.isna(nota) else (int(nota) if float(nota).is_integer() else nota),
            "Incidência": ("não avaliada" if pd.isna(inc) else ("presente" if inc == 1 else "ausente")),
            "Classe": "—" if (isinstance(classe, float) and pd.isna(classe)) else classe,
        })
    _sty = pd.DataFrame(linhas_tab).style
    _sty = _styler_map(_sty, _estilo_incidencia, ["Incidência"])
    _sty = _styler_map(_sty, _estilo_classe, ["Classe"])
    st.dataframe(_sty, use_container_width=True, hide_index=True)

    # 2) a régua explicada em duas tabelas (classe e incidência) + exemplo do plot
    st.markdown("**2. As regras do cálculo**")
    st.markdown(
        "<div class='audit-step'>Para cada doença o técnico dá uma nota de <b>1 a 9</b>: "
        "<b>quanto maior a nota, melhor</b> — a planta está mais sadia e resistente. "
        "A partir dessa nota, a regra deriva a <b>classe de reação</b> e a <b>incidência</b>, "
        "assim:</div>", unsafe_allow_html=True)

    cca, ccb = st.columns(2)

    # tabela de CLASSES (nota → sigla → nome), colorida do pior ao melhor
    CLASSES_INFO = [
        ("1 – 2", "AS", "Altamente suscetível", "muito afetada pela doença"),
        ("3 – 4", "S",  "Suscetível",           "bastante afetada"),
        ("5 – 6", "MT", "Medianamente tolerante", "afetada de forma intermediária"),
        ("7 – 8", "T",  "Tolerante",            "pouco afetada"),
        ("9",     "R",  "Resistente",           "praticamente sem doença"),
    ]
    linhas_classe = "".join(
        f"<tr>"
        f"<td style='padding:6px 10px;font-family:ui-monospace,monospace;'>{faixa}</td>"
        f"<td style='padding:6px 10px;background:{COR_CLASSE[sig]};font-weight:700;text-align:center;'>{sig}</td>"
        f"<td style='padding:6px 10px;'>{nome}</td>"
        f"<td style='padding:6px 10px;color:#6B7280;font-size:12px;'>{desc}</td>"
        f"</tr>"
        for faixa, sig, nome, desc in CLASSES_INFO
    )
    with cca:
        st.markdown(
            "<p style='font-size:13px;font-weight:600;margin:6px 0 4px;'>Classe de reação</p>"
            "<table style='border-collapse:collapse;width:100%;font-size:13px;'>"
            "<thead><tr style='color:#6B7280;font-size:11px;text-transform:uppercase;'>"
            "<th style='text-align:left;padding:4px 10px;'>Nota</th>"
            "<th style='text-align:center;padding:4px 10px;'>Sigla</th>"
            "<th style='text-align:left;padding:4px 10px;'>Significado</th>"
            "<th style='text-align:left;padding:4px 10px;'></th></tr></thead>"
            f"<tbody>{linhas_classe}</tbody></table>", unsafe_allow_html=True)

    # tabela de INCIDÊNCIA (nota → status), com cinza/verde/vermelho
    INCID_INFO = [
        ("1 – 5", "presente",      COR_INCIDENCIA["presente"],      "a planta mostrou sintomas da doença"),
        ("6 – 9", "ausente",       COR_INCIDENCIA["ausente"],       "a planta estava sadia"),
        ("0",     "não avaliada",  COR_INCIDENCIA["não avaliada"],  "a doença não foi avaliada neste plot"),
    ]
    linhas_incid = "".join(
        f"<tr>"
        f"<td style='padding:6px 10px;font-family:ui-monospace,monospace;'>{faixa}</td>"
        f"<td style='padding:6px 10px;background:{cor};font-weight:600;text-align:center;'>{status}</td>"
        f"<td style='padding:6px 10px;color:#6B7280;font-size:12px;'>{desc}</td>"
        f"</tr>"
        for faixa, status, cor, desc in INCID_INFO
    )
    with ccb:
        st.markdown(
            "<p style='font-size:13px;font-weight:600;margin:6px 0 4px;'>Incidência</p>"
            "<table style='border-collapse:collapse;width:100%;font-size:13px;'>"
            "<thead><tr style='color:#6B7280;font-size:11px;text-transform:uppercase;'>"
            "<th style='text-align:left;padding:4px 10px;'>Nota</th>"
            "<th style='text-align:center;padding:4px 10px;'>Status</th>"
            "<th style='text-align:left;padding:4px 10px;'></th></tr></thead>"
            f"<tbody>{linhas_incid}</tbody></table>", unsafe_allow_html=True)

    # escolhe uma doença avaliada do plot para o exemplo (a primeira com nota > 0)
    exemplo_txt = None
    for d in doencas_safra:
        nota, inc, classe = avaliar_doenca(linha.get(d))
        if pd.notna(nota):
            nm = LABEL_DOENCA.get(d, d)
            n_int = int(nota) if float(nota).is_integer() else nota
            pres = "presente" if inc == 1 else "ausente"
            faixa = {"AS": "1–2", "S": "3–4", "MT": "5–6", "T": "7–8", "R": "9"}.get(classe, "")
            exemplo_txt = (f"Em <b>{nm}</b> a nota foi <b>{n_int}</b>. Como {n_int} está "
                           f"{'entre 1 e 5' if inc == 1 else 'entre 6 e 9'}, a doença é "
                           f"<b>{pres}</b>. E como {n_int} cai na faixa {faixa}, a classe de reação é "
                           f"<b>{classe}</b>.")
            break
    if exemplo_txt is None:
        exemplo_txt = "Nenhuma doença foi avaliada neste plot (todas as notas são 0)."
    st.markdown(f"<div class='audit-var audit-step'><b>Neste caso:</b> {exemplo_txt}</div>",
                unsafe_allow_html=True)

    # 3) conferência: recálculo × gold, doença a doença
    st.markdown("**3. Resultado da conferência**")
    conf = []
    todas_batem = True
    for d in doencas_safra:
        nota, inc, classe = avaliar_doenca(linha.get(d))
        inc_gold = linha.get(f"inc_{d}")
        cls_gold = linha.get(f"class_{d}")
        inc_ok = (pd.isna(inc) and pd.isna(inc_gold)) or (pd.notna(inc) and pd.notna(inc_gold) and int(inc) == int(inc_gold))
        cls_ok = (
            (isinstance(classe, float) and pd.isna(classe)) and (isinstance(cls_gold, float) and pd.isna(cls_gold))
        ) or (str(classe) == str(cls_gold))
        ok = inc_ok and cls_ok
        todas_batem = todas_batem and ok
        conf.append({
            "Doença": LABEL_DOENCA.get(d, d),
            "Incidência (recalc.)": ("—" if pd.isna(inc) else ("presente" if inc == 1 else "ausente")),
            "Incidência (sistema)": ("—" if pd.isna(inc_gold) else ("presente" if int(inc_gold) == 1 else "ausente")),
            "Classe (recalc.)": "—" if (isinstance(classe, float) and pd.isna(classe)) else classe,
            "Classe (sistema)": "—" if (isinstance(cls_gold, float) and pd.isna(cls_gold)) else cls_gold,
            "Confere": "✓" if ok else "✗",
        })
    def _estilo_confere(v):
        s = str(v).strip()
        if s == "✓":
            return "background-color: #C6EFCE; color: #1E6B34; font-weight: 700;"
        if s == "✗":
            return "background-color: #F8D7DA; color: #B02A37; font-weight: 700;"
        return ""

    _sty2 = pd.DataFrame(conf).style
    _sty2 = _styler_map(_sty2, _estilo_incidencia, ["Incidência (recalc.)", "Incidência (sistema)"])
    _sty2 = _styler_map(_sty2, _estilo_classe, ["Classe (recalc.)", "Classe (sistema)"])
    _sty2 = _styler_map(_sty2, _estilo_confere, ["Confere"])
    st.dataframe(_sty2, use_container_width=True, hide_index=True)
    if todas_batem:
        st.markdown("<div class='audit-var'><span class='audit-ok'>✓ CONFERE</span><br>"
                    "<span style='font-size:12px;color:#6B7280'>todas as doenças batem com o sistema</span></div>",
                    unsafe_allow_html=True)
    else:
        st.markdown("<div class='audit-var'><span class='audit-bad'>✗ DIVERGE</span><br>"
                    "<span style='font-size:12px;color:#6B7280'>há diferença em ao menos uma doença</span></div>",
                    unsafe_allow_html=True)

elif aval == "av3":
    # a av3 tem duas naturezas: ALTURA (média de 5 subamostras, com padronização) e
    # FLORESCIMENTO (diferença de datas, com validação de plausibilidade).
    av3_det = dados.get("av3_detalhe")

    # ── ALTURA ──
    st.markdown("**1. Altura — as 5 medições e a média**")
    st.markdown(
        "<div class='audit-step'>O técnico mede a altura de <b>5 plantas</b> do plot "
        "(altura da planta e altura da espiga). Antes de tirar a média, cada medição passa por "
        "uma limpeza: valor <b>0</b> é “não medido” e sai da conta; valor muito alto "
        "(<b>acima de 350 cm</b>) é erro de digitação e é descartado; e valor <b>abaixo de 10</b> "
        "foi digitado em metros (ex.: 2,1) e é convertido para centímetros (×100). "
        "A <b>altura final</b> do plot é a média das medições que sobraram.</div>",
        unsafe_allow_html=True)

    # recupera as 5 subamostras (já padronizadas) do detalhe, para este plot
    uuid_plot = linha.get("uuid")
    if av3_det is not None and "uuid" in av3_det.columns and pd.notna(uuid_plot):
        sub_det = av3_det[av3_det["uuid"] == uuid_plot]
        tab_alt = {"Planta": [1, 2, 3, 4, 5]}
        for metrica, rot_cm, rot_m in [
            ("altura_planta_cm", "Altura planta (cm)", "Altura planta (m)"),
            ("altura_espiga_cm", "Altura espiga (cm)", "Altura espiga (m)")]:
            vals = []
            for p in range(1, 6):
                v = sub_det[(sub_det["planta"] == p) & (sub_det["metrica"] == metrica)]["valor"]
                vals.append(v.iloc[0] if len(v) else np.nan)
            tab_alt[rot_cm] = ["—" if pd.isna(x) else f"{float(x):.1f}" for x in vals]
            tab_alt[rot_m] = ["—" if pd.isna(x) else f"{float(x)/100:.2f}" for x in vals]
        st.dataframe(pd.DataFrame(tab_alt), use_container_width=True, hide_index=True)
    else:
        st.caption("Medições individuais não disponíveis no detalhe; conferindo só a média final.")

    # ── FLORESCIMENTO ──
    st.markdown("**Florescimento — a conta dos dias**")
    st.markdown(
        "<div class='audit-step'>Os dias até o florescimento são a diferença entre a data do "
        "florescimento e a data do plantio. O milho safrinha normalmente floresce entre "
        f"<b>{FLOR_MIN} e {FLOR_MAX} dias</b> após o plantio. Se a conta der um valor fora dessa "
        "faixa, provavelmente houve erro ao digitar alguma das datas, então esse valor não entra "
        "na análise.</div>", unsafe_allow_html=True)

    # 2) regras já explicadas acima; 3) conferência com a consolidada
    st.markdown("**2. Resultado da conferência**")

    # altura: recalcula a média das subamostras e compara com a consolidada
    conf_linhas = []
    if av3_det is not None and pd.notna(uuid_plot):
        sub_det = av3_det[av3_det["uuid"] == uuid_plot]
        for metrica, gold_col, rotulo in [
            ("altura_planta_cm", "altura_planta_cm", "Altura da planta (cm)"),
            ("altura_espiga_cm", "altura_espiga_cm", "Altura da espiga (cm)")]:
            serie = sub_det[sub_det["metrica"] == metrica]["valor"]
            vals = pd.to_numeric(serie, errors="coerce")
            n_medidas = len(serie)                    # quantas subamostras existem (esperado 5)
            n_validas = int(vals.notna().sum())       # quantas sobraram após a regra
            media_recalc = round(vals.mean(), 1) if vals.notna().any() else np.nan
            media_gold = linha.get(gold_col)
            ok = (pd.isna(media_recalc) and pd.isna(media_gold)) or (
                pd.notna(media_recalc) and pd.notna(media_gold) and abs(float(media_recalc) - float(media_gold)) < 0.05)
            # verificar: alguma medição foi descartada pela regra OU a conferência divergiu
            houve_descarte = n_validas < n_medidas or n_validas == 0
            verificar = houve_descarte or not ok
            if not verificar:
                motivo = ""
            elif not ok:
                motivo = "valores não batem"
            elif n_validas == 0:
                motivo = "nenhuma medição válida"
            else:
                motivo = f"{n_medidas - n_validas} medição(ões) descartada(s)"
            conf_linhas.append({
                "Medida": rotulo,
                "Pela regra": "—" if pd.isna(media_recalc) else f"{media_recalc:.1f}",
                "No sistema": "—" if pd.isna(media_gold) else f"{float(media_gold):.1f}",
                "Em metros": "—" if pd.isna(media_gold) else f"{float(media_gold)/100:.2f}",
                "Confere": "✓" if ok else "✗",
                "Verificar": f"⚠ {motivo}" if verificar else "",
            })

    # florescimento: confere se a validação do sistema segue a regra (implausível → vazio)
    for col, rotulo in [("dias_flor_masculino", "Florescimento masculino (dias)"),
                        ("dias_flor_feminino", "Florescimento feminino (dias)")]:
        bruto = linha.get(col)
        valido = linha.get(f"{col}_valido")
        obs = linha.get(f"obs_{col}")
        implausivel = str(obs) == "data_implausivel"
        if pd.isna(bruto):
            situacao = "sem data"
        elif implausivel:
            situacao = f"{int(bruto)} dias — fora de {FLOR_MIN}–{FLOR_MAX} (descartado)"
        else:
            situacao = f"{int(bruto)} dias — dentro de {FLOR_MIN}–{FLOR_MAX}"
        # regra: se implausível, o válido tem que estar vazio; se plausível, válido == bruto
        if pd.isna(bruto):
            ok = pd.isna(valido)
        elif implausivel:
            ok = pd.isna(valido)
        else:
            ok = pd.notna(valido) and int(valido) == int(bruto)
        conf_linhas.append({
            "Medida": rotulo,
            "Pela regra": situacao,
            "No sistema": ("vazio" if pd.isna(valido) else f"{int(valido)} dias"),
            "Em metros": "—",
            "Confere": "✓" if ok else "✗",
            "Verificar": ("⚠ data fora da faixa" if (pd.notna(bruto) and implausivel)
                          else ("⚠ valores não batem" if not ok else "")),
        })

    if conf_linhas:
        def _cor_confere(v):
            s = str(v).strip()
            if s == "✓":
                return "background-color:#C6EFCE;color:#1E6B34;font-weight:700;"
            if s == "✗":
                return "background-color:#F8D7DA;color:#B02A37;font-weight:700;"
            return ""
        def _cor_verificar(v):
            return "background-color:#FFF3CD;color:#8A6D00;font-weight:600;" if str(v).strip() else ""
        _sty3 = pd.DataFrame(conf_linhas).style
        _sty3 = _styler_map(_sty3, _cor_confere, ["Confere"])
        _sty3 = _styler_map(_sty3, _cor_verificar, ["Verificar"])
        st.dataframe(_sty3, use_container_width=True, hide_index=True)

elif aval == "av4":
    # A av4 é a mais rica: produtividade (peso/umidade/área/correção), população, PMG,
    # os dois caminhos (medido × estimado) e as perdas. Mostro em blocos temáticos.
    metros_ct = 10 if is_2025 else 10   # metragem da contagem (10 nas duas safras)
    n_sub = N_SUBAMOSTRAS_2025 if is_2025 else N_SUBAMOSTRAS_2024   # protocolo: 4 em 2025, 5 em 2024

    def _num(col, casas=1):
        v = linha.get(col)
        return "—" if pd.isna(v) else f"{float(v):.{casas}f}"

    # ── BLOCO 1: PRODUTIVIDADE (Caminho A — medida pelo peso) ──
    st.markdown("**1. Produtividade — medida pelo peso da parcela**")
    st.markdown(
        "<div class='audit-step'>A produtividade principal vem do <b>peso colhido</b> na parcela. "
        "Primeiro calcula-se a <b>área</b> da parcela (linhas × comprimento × espaçamento). "
        "O peso é corrigido para a umidade padrão de <b>13,5%</b> (para comparar todos os plots "
        "na mesma base) e então convertido para <b>kg por hectare</b>. "
        "Uma saca tem 60 kg, então sacas/ha = kg/ha ÷ 60.</div>", unsafe_allow_html=True)

    # variáveis cruas do Caminho A
    peso_cru = linha.get("pesoParcela")
    umid = linha.get("umidade_pct")
    area = linha.get("area_parcela_m2")
    nlin = linha.get("numeroLinhas"); compr = linha.get("comprimentoLinha"); espac = linha.get("espacamento")
    st.dataframe(pd.DataFrame({
        "Variável": ["Peso da parcela", "Umidade do grão", "Nº de linhas", "Comprimento da linha",
                     "Espaçamento", "Área da parcela"],
        "Valor": [f"{float(peso_cru):.2f} kg" if pd.notna(peso_cru) else "—",
                  f"{float(umid):.1f} %" if pd.notna(umid) else "—",
                  _num("numeroLinhas", 0) if pd.notna(nlin) else "—",
                  f"{float(compr):.1f} m" if pd.notna(compr) else "—",
                  f"{float(espac):.2f} m" if pd.notna(espac) else "—",
                  f"{float(area):.2f} m²" if pd.notna(area) else "—"],
    }), use_container_width=True, hide_index=True)

    # a conta com os números do plot
    if pd.notna(area) and pd.notna(peso_cru) and pd.notna(umid):
        peso_corr = float(peso_cru) * (100 - float(umid)) / (100 - UMID_PADRAO)
        st.markdown(
            f"<div class='audit-var audit-step'><b>Neste caso:</b> "
            f"área = {float(nlin):.0f} × {float(compr):.1f} × {float(espac):.2f} = "
            f"<b>{float(area):.2f} m²</b>. "
            f"Peso corrigido a 13,5% = {float(peso_cru):.2f} × (100−{float(umid):.1f})/(100−13,5) = "
            f"<b>{peso_corr:.2f} kg</b>. "
            f"Produtividade = {peso_corr:.2f} × (10000/{float(area):.2f}) = "
            f"<b>{_num('produtividade_kg_ha')} kg/ha</b> ({_num('produtividade_sacas_ha')} sacas/ha).</div>",
            unsafe_allow_html=True)

    flags = linha.get("flags_produtividade")
    if flags and str(flags).strip():
        st.markdown(f"<div class='audit-var' style='background:#FFF3CD;'>⚠ Sinalizações: "
                    f"<b>{flags}</b>. Flags bloqueantes (não colhido, sem geometria, produtividade "
                    f"impossível) zeram a produtividade válida.</div>", unsafe_allow_html=True)

    # ── BLOCO 2: POPULAÇÃO ──
    st.markdown("**2. População de plantas**")
    st.markdown(
        f"<div class='audit-step'>A população é estimada a partir da contagem de plantas em "
        f"<b>{metros_ct} metros</b> de linha. Dividindo pela metragem e pelo espaçamento, e "
        f"convertendo para hectare (×10000), chega-se ao número de <b>plantas por hectare</b>.</div>",
        unsafe_allow_html=True)
    plantas_10m = linha.get("plantas_10m_media")
    if pd.notna(plantas_10m) and pd.notna(espac):
        st.markdown(
            f"<div class='audit-var audit-step'><b>Neste caso:</b> {float(plantas_10m):.1f} plantas em "
            f"{metros_ct} m, espaçamento {float(espac):.2f} m → "
            f"({float(plantas_10m):.1f} / {metros_ct} / {float(espac):.2f}) × 10000 = "
            f"<b>{_num('populacao_real_plantas_ha', 0)} plantas/ha</b>.</div>", unsafe_allow_html=True)

    # ── BLOCO 3: COMPONENTES E PMG ──
    st.markdown("**3. Componentes da espiga e PMG**")
    st.markdown(
        "<div class='audit-step'>Média das subamostras: número de <b>fileiras</b> por espiga, "
        "<b>grãos por fileira</b> e o <b>peso de mil grãos (PMG)</b> — este já corrigido para a "
        "umidade padrão de 13,5%.</div>", unsafe_allow_html=True)
    st.dataframe(pd.DataFrame({
        "Componente": ["Fileiras por espiga", "Grãos por fileira", "PMG corrigido (g)"],
        "Média": [_num("fileiras_media"), _num("graos_fileira_media"), _num("pmg_corrigido_g")],
    }), use_container_width=True, hide_index=True)

    # ── DECOMPOSIÇÃO: o que compõe cada média (para auditar refazendo a conta) ──
    with st.expander("Ver as medições que compõem as médias"):
        av4_det = dados.get("av4_detalhe")
        uuid_plot = linha.get("uuid")

        # tabela 1: componentes por planta (5 subamostras) → média por coluna
        st.markdown(f"**Componentes por planta ({n_sub} subamostras)**")
        if av4_det is not None and "uuid" in av4_det.columns and pd.notna(uuid_plot):
            sub = av4_det[av4_det["uuid"] == uuid_plot]
            comp_metricas = [("fileiras", "Fileiras"), ("graos_fileira", "Grãos/fileira"),
                             ("pmg_bruto_g", "PMG bruto (g)"), ("umidade_pmg_pct", "Umid. amostra (%)"),
                             ("pmg_corrigido_g", "PMG corrig. (g)")]
            tab = {"Planta": list(range(1, n_sub + 1))}
            medias_calc = {}
            for met, rot in comp_metricas:
                vals = []
                for p in range(1, n_sub + 1):
                    v = sub[(sub["planta"] == p) & (sub["metrica"] == met)]["valor"]
                    vals.append(pd.to_numeric(v, errors="coerce").iloc[0] if len(v) else np.nan)
                serie = pd.Series(vals)
                medias_calc[rot] = serie.mean()
                tab[rot] = ["—" if pd.isna(x) else f"{float(x):.1f}" for x in vals]
            df_comp = pd.DataFrame(tab)
            # linha de média no rodapé
            linha_media = {"Planta": "Média"}
            for _, rot in comp_metricas:
                m_ = medias_calc[rot]
                linha_media[rot] = "—" if pd.isna(m_) else f"{m_:.1f}"
            df_comp = pd.concat([df_comp, pd.DataFrame([linha_media])], ignore_index=True)

            def _destaca_media(row):
                return ["background-color:#EAF3FF;font-weight:700;" if row["Planta"] == "Média" else ""
                        for _ in row.index]
            _styc = df_comp.style.apply(_destaca_media, axis=1)
            st.dataframe(_styc, use_container_width=True, hide_index=True)
            st.caption(f"A linha Média é a média das {n_sub} subamostras — deve bater com os valores "
                       "do bloco acima (o PMG corrigido usa a umidade de cada amostra para ajustar "
                       "a 13,5%).")
        else:
            st.caption("Detalhe das subamostras não disponível para este plot.")

        # tabela 2: estande (8 pontos em 2025, 5 subamostras em 2024) → média + população por ponto
        st.markdown("**Contagem de estande e população**")
        cols_est = COLS_ESTANDE_8P if is_2025 else COLS_ESTANDE_5SUB
        cols_est_presentes = [c for c in cols_est if c in linha.index]
        espac_plot = pd.to_numeric(pd.Series([linha.get("espacamento")]), errors="coerce").iloc[0]
        if cols_est_presentes:
            pontos = pd.to_numeric(pd.Series([linha.get(c) for c in cols_est_presentes]), errors="coerce")
            pontos_validos = pontos.where(pontos > 0)   # 0 não conta (mesma regra do pipeline)
            media_est = pontos_validos.mean()
            # população que cada ponto geraria sozinho (mesma fórmula do gold)
            if pd.notna(espac_plot) and 0 < espac_plot <= 2:
                pop_ponto = (pontos_validos / metros_ct / espac_plot) * 10000
                media_pop = (media_est / metros_ct / espac_plot) * 10000 if pd.notna(media_est) else np.nan
            else:
                pop_ponto = pd.Series([np.nan] * len(pontos))
                media_pop = np.nan
            n_label = f"{len(cols_est_presentes)} pontos" if is_2025 else f"{n_sub} subamostras"
            df_est = pd.DataFrame({
                "Ponto": [f"Ponto {i+1}" for i in range(len(cols_est_presentes))] + ["Média"],
                "Plantas em 10 m": [("—" if pd.isna(x) else f"{float(x):.0f}") for x in pontos]
                                   + ["—" if pd.isna(media_est) else f"{media_est:.1f}"],
                "População (plantas/ha)": [("—" if pd.isna(x) else f"{float(x):,.0f}".replace(",", ".")) for x in pop_ponto]
                                   + ["—" if pd.isna(media_pop) else f"{media_pop:,.0f}".replace(",", ".")],
            })
            def _destaca_media2(row):
                return ["background-color:#EAF3FF;font-weight:700;" if row["Ponto"] == "Média" else ""
                        for _ in row.index]
            st.dataframe(df_est.style.apply(_destaca_media2, axis=1),
                         use_container_width=True, hide_index=True)
            st.caption(f"Estande = média de {n_label} de contagem (o valor 0 não entra). A população de "
                       f"cada ponto usa a mesma fórmula (plantas ÷ {metros_ct} ÷ espaçamento × 10.000); "
                       f"a média da última coluna é a população final do plot.")
        else:
            st.caption("Colunas de contagem de estande não disponíveis para este plot.")

    # ── BLOCO 4: CAMINHO B (validação) ──
    st.markdown("**4. Conferência cruzada — produtividade estimada pelos componentes**")
    st.markdown(
        "<div class='audit-step'>Existe um segundo jeito de estimar a produtividade, a partir dos "
        "componentes (população × fileiras × grãos × peso do grão). Ele não é o valor oficial — "
        "serve para <b>cruzar</b> com a produtividade medida pelo peso. Se os dois batem, dá "
        "confiança; se divergem muito, vale investigar o plot.</div>", unsafe_allow_html=True)
    div = linha.get("divergencia_prod_pct")
    c1, c2, c3 = st.columns(3)
    c1.metric("Medida (peso)", f"{_num('produtividade_valida_kg_ha')} kg/ha")
    c2.metric("Estimada (componentes)", f"{_num('prod_estimada_kg_ha')} kg/ha")
    c3.metric("Divergência", "—" if pd.isna(div) else f"{float(div):+.1f} %")

    # ── BLOCO 5: PERDAS ──
    st.markdown("**5. Perdas de colheita**")
    st.markdown(
        "<div class='audit-step'>Percentual de plantas com problema, sobre o estande final: "
        "<b>acamadas</b> (tombadas), <b>quebradas</b>, <b>dominadas</b> (abafadas por vizinhas) e "
        f"<b>colmo podre</b>. Conta-se em <b>{n_sub} subamostras</b>; cada uma vira um percentual "
        "sobre o estande final do plot e o valor abaixo é a <b>média</b> delas. Contagem <b>0 "
        "conta</b> (o avaliador percorreu os 10 m e não achou); só a subamostra não avaliada fica "
        "de fora. A perda total é a soma dos quatro.</div>", unsafe_allow_html=True)
    perdas_tab = {
        "Perda": ["Acamadas", "Quebradas", "Dominadas", "Colmo podre", "Total"],
        "Percentual": [_num("pct_acamadas"), _num("pct_quebradas"), _num("pct_dominadas"),
                       _num("pct_colmo_podre"), _num("pct_perda_total")],
    }
    st.dataframe(pd.DataFrame(perdas_tab), use_container_width=True, hide_index=True)

    # ── Helpers de conferência (usados pelos blocos 5 e 6) ──────────────────
    # Contagens e percentuais por subamostra vêm do av4_detalhe (long), onde perda e
    # fenômeno convivem na mesma subamostra, como no aplicativo.
    _av4_det = dados.get("av4_detalhe")
    _uuid = linha.get("uuid")
    _est = pd.to_numeric(pd.Series([linha.get("plantas_10m_media")]), errors="coerce").iloc[0]
    _sub_det = (_av4_det[_av4_det["uuid"] == _uuid]
                if _av4_det is not None and "uuid" in _av4_det.columns and pd.notna(_uuid)
                else None)

    def _contagens(chaves):
        """{rótulo: [contagem por subamostra]} para um grupo de variáveis."""
        out = {}
        if _sub_det is None:
            return out
        for _canon, _rot in chaves:
            vals = []
            for p in range(1, n_sub + 1):
                v = _sub_det[(_sub_det["planta"] == p) & (_sub_det["metrica"] == f"{_canon}_n")]["valor"]
                vals.append(pd.to_numeric(v, errors="coerce").iloc[0] if len(v) else np.nan)
            if not all(pd.isna(x) for x in vals):
                out[_rot] = vals
        return out

    def _tabela_contagens(dados_cont, com_total):
        """Contagens por subamostra + (opcional) total da subamostra + linha de média."""
        tab = {"Subamostra": [f"{p}ª" for p in range(1, n_sub + 1)]}
        for rot, vals in dados_cont.items():
            tab[rot] = ["—" if pd.isna(x) else f"{float(x):.0f}" for x in vals]
        _tot = None
        if com_total and dados_cont:
            _tot = [sum(v[i] for v in dados_cont.values() if pd.notna(v[i]))
                    if any(pd.notna(v[i]) for v in dados_cont.values()) else np.nan
                    for i in range(n_sub)]
            tab["Total"] = ["—" if pd.isna(x) else f"{float(x):.0f}" for x in _tot]
        df_c = pd.DataFrame(tab)
        medias = {r: pd.Series(v, dtype="float").mean() for r, v in dados_cont.items()}
        if _tot is not None:
            medias["Total"] = pd.Series(_tot, dtype="float").mean()
        lin = {"Subamostra": "Média"}
        lin.update({k: ("—" if pd.isna(v) else f"{v:.2f}") for k, v in medias.items()})
        df_c = pd.concat([df_c, pd.DataFrame([lin])], ignore_index=True)

        def _dest(row):
            return ["background-color:#EAF3FF;font-weight:700;" if row["Subamostra"] == "Média"
                    else "" for _ in row.index]
        st.markdown("**1. Plantas contadas em cada subamostra**")
        st.dataframe(df_c.style.apply(_dest, axis=1), use_container_width=True, hide_index=True)
        return medias

    def _tabela_conferencia(medias, alvos):
        """Média ÷ estande × 100 comparada com o valor do painel."""
        st.markdown(f"**2. Média ÷ estande final ({_est:.1f} plantas em 10 m) × 100**")
        linhas = []
        for rot, col_gold in alvos:
            m = medias.get(rot)
            calc = np.nan if m is None or pd.isna(m) else round(m / _est * 100, 1)
            gold = (pd.to_numeric(pd.Series([linha.get(col_gold)]), errors="coerce").iloc[0]
                    if col_gold else np.nan)
            ok = (pd.isna(calc) and pd.isna(gold)) or (
                pd.notna(calc) and pd.notna(gold) and abs(calc - gold) <= 0.1)
            linhas.append({
                "Variável": rot,
                "Conta refeita": "—" if pd.isna(calc) else f"{calc:.1f}%",
                "No painel": "—" if (col_gold is None or pd.isna(gold)) else f"{float(gold):.1f}%",
                "Confere": ("—" if col_gold is None else ("sim" if ok else "NÃO")),
            })
        st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)
        st.caption("A contagem 0 entra na média — é medição, não ausência. Subamostra em branco "
                   "fica de fora. Diferença de até 0,1 ponto é arredondamento.")

    # métricas de contagem presentes na carga inteira (não só neste plot) — serve para
    # separar "este plot não tem" de "esta carga não tem a métrica"
    _metricas_carga = (set(_av4_det["metrica"].unique())
                       if _av4_det is not None and "metrica" in _av4_det.columns else set())

    def _conferencia_indisponivel(chaves):
        if _sub_det is None:
            st.caption("Detalhe das subamostras não disponível para este plot.")
        elif pd.isna(_est) or _est <= 0:
            st.caption("Estande final indisponível para este plot — sem denominador não há taxa.")
        elif not any(f"{c}_n" in _metricas_carga for c, _ in chaves):
            st.warning("Estas contagens não existem nos dados carregados. Elas passaram a ser "
                       "geradas quando o cálculo foi unificado no pipeline — **limpe o cache do "
                       "Streamlit e recarregue** para que apareçam. Enquanto isso, o percentual "
                       "mostrado acima vem da carga antiga e não corresponde à régua atual.")
        else:
            st.caption("Nenhuma contagem registrada neste plot.")

    with st.expander("Refazer a conta das perdas subamostra a subamostra"):
        _chaves_perda = [(k, k.replace("_", " ").capitalize()) for k in MAPA_PERDAS]
        _cont_perda = _contagens(_chaves_perda) if (_sub_det is not None and pd.notna(_est) and _est > 0) else {}
        if not _cont_perda:
            _conferencia_indisponivel(_chaves_perda)
        else:
            _m = _tabela_contagens(_cont_perda, com_total=True)
            _alvos = [(r, f"pct_{k}") for k, r in _chaves_perda if r in _cont_perda]
            _alvos.append(("Total", "pct_perda_total"))
            _tabela_conferencia(_m, _alvos)


    # fenômenos (só 2025)
    fen_cols = [("green_snap", "Green snap"), ("morte_prematura", "Morte prematura"),
                ("ma_formacao_espigas", "Má formação de espigas"), ("enfezamento", "Enfezamento")]
    tem_fen = any(f"pct_{c}" in linha.index and pd.notna(linha.get(f"pct_{c}")) for c, _ in fen_cols)
    if tem_fen:
        st.markdown("**6. Fenômenos da colheita (novos em 2025)**")
        st.markdown(
            f"<div class='audit-step'>Mesma régua das perdas, e no aplicativo é a <b>mesma "
            f"subamostra</b>: são as mesmas {n_sub} contagens, cada uma vira um percentual sobre o "
            "estande final e o valor abaixo é a <b>média</b> delas, com o 0 contando. Não entra na "
            "soma da perda total — é outra contagem sobre o mesmo estande.</div>",
            unsafe_allow_html=True)
        st.dataframe(pd.DataFrame({
            "Fenômeno": [nome for _, nome in fen_cols],
            "Percentual": [_num(f"pct_{c}") for c, _ in fen_cols],
        }), use_container_width=True, hide_index=True)

        with st.expander("Refazer a conta dos fenômenos subamostra a subamostra"):
            _chaves_fen = [(k, k.replace("_", " ").capitalize()) for k in FENOMENOS_AV4]
            _cont_fen = (_contagens(_chaves_fen)
                         if (_sub_det is not None and pd.notna(_est) and _est > 0) else {})
            if not _cont_fen:
                _conferencia_indisponivel(_chaves_fen)
            else:
                # sem coluna Total: somar fenômenos só faria sentido se um mesmo pé não
                # pudesse cair em duas categorias, e isso não está definido no protocolo.
                _mf = _tabela_contagens(_cont_fen, com_total=False)
                _alvos_f = [(r, f"pct_{k}") for k, r in _chaves_fen if r in _cont_fen]
                _tabela_conferencia(_mf, _alvos_f)

with st.expander("Contexto do plot (para localizar no aplicativo)"):
    ctx = ["cod_fazenda", "nomeFazenda", "cidade_nome", "estado_sigla", "nomeResponsavel",
           "dePara", "status_material", "tipoTeste", "indexTratamento",
           "qualidade_plot_inicial", "media_categorias"]
    st.dataframe(pd.DataFrame({"campo": [c for c in ctx if c in linha.index],
                               "valor": [linha.get(c) for c in ctx if c in linha.index]}),
                 use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# PARTE 2 — SIMULADOR (só av1, só 2025)
# ══════════════════════════════════════════════════════════════════════════════
if is_2025 and aval == "av1":
    st.divider()
    secao_titulo("Simulador", "Teste as regras com valores seus",
                 "Digite 8 notas quaisquer (0 = não avaliado) e veja a média sair — inclusive o "
                 "caso de escala errada (qualquer nota ≥ 6 zera a média do plot).")

    with st.form("sim_av1"):
        cols = st.columns(4)
        entradas = {}
        for i, canon in enumerate(NOTAS_AV1):
            with cols[i % 4]:
                entradas[canon] = st.number_input(LABEL_NOTA[canon], min_value=0, max_value=9,
                                                  value=5, step=1, key=f"sim_{canon}")
        ok = st.form_submit_button("Calcular", type="primary")

    if ok:
        media_sim, escala_sim, nval_sim, soma_sim = calcular_media_av1(list(entradas.values()))
        if escala_sim:
            explica = "Há nota ≥ 6 (escala 1–9) → média = NaN no plot inteiro."
        elif nval_sim == 0:
            explica = "Todas as notas são 0 (não avaliado) → média = NaN."
        else:
            explica = f"Soma {soma_sim:.0f} / {nval_sim} válidas = {soma_sim/nval_sim:.4f} → {media_sim}"
        cA, cB = st.columns([1, 2])
        cA.metric("media_categorias", "NaN" if pd.isna(media_sim) else f"{media_sim}")
        cB.markdown(f"<div class='audit-var audit-step'>{explica}</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PARTE 3 — SIMULADOR DE PERDAS E FENÔMENOS (as duas safras)
# ══════════════════════════════════════════════════════════════════════════════
st.divider()
secao_titulo("Simulador", "Teste a conta da perda com números seus",
             "Digite o estande e as contagens de uma subamostra por vez e veja a taxa sair. "
             "Serve para conferir um caderno de campo ou para entender por que dois plots "
             "parecidos dão resultados diferentes.")

with st.form("sim_perdas"):
    est_sim = st.number_input(
        "Plantas contadas em 10 metros (estande final)",
        min_value=1.0, max_value=200.0, value=27.5, step=0.5,
        help="É o denominador: a média dos pontos de contagem do plot.")

    st.markdown("**Plantas com o problema em cada subamostra**")
    st.caption("Um campo por subamostra, como no aplicativo. **Deixe vazio** a subamostra que não "
               "foi avaliada — ela sai da conta. **Zero é medição**: percorreu os 10 metros e não "
               "encontrou nada, e isso entra na média.")
    _defaults = [0, 0, 1, 0, None]   # exemplo: 4 subamostras avaliadas, a 5ª fora do protocolo
    _cols_sim = st.columns(MAX_SLOTS_SUBAMOSTRA)
    _entradas = []
    for _i in range(MAX_SLOTS_SUBAMOSTRA):
        with _cols_sim[_i]:
            _entradas.append(st.number_input(
                f"{_i + 1}ª", min_value=0, max_value=999, step=1,
                value=_defaults[_i], placeholder="vazio", key=f"sim_perda_sub{_i + 1}"))
    ok_sim = st.form_submit_button("Calcular", type="primary")

if ok_sim:
    _vals = [np.nan if v is None else float(v) for v in _entradas]
    _s = pd.Series(_vals, dtype="float")
    _n_aval = int(_s.notna().sum())

    _df_sim = pd.DataFrame({
        "Subamostra": [f"{i}ª" for i in range(1, len(_vals) + 1)],
        "Plantas contadas": ["não avaliada" if pd.isna(v) else f"{v:.0f}" for v in _vals],
        "% sobre o estande": ["—" if pd.isna(v) else f"{v / est_sim * 100:.1f}%" for v in _vals],
    })
    st.dataframe(_df_sim, use_container_width=True, hide_index=True)

    if _n_aval == 0:
        st.info("Nenhuma subamostra preenchida — sem taxa. No painel o plot ficaria vazio "
                "nesta variável.")
    else:
        _media = _s.mean()
        _soma = _s.sum()
        _taxa = _media / est_sim * 100
        st.success(f"**Taxa = {_taxa:.1f}%**")
        st.markdown(f"""
**Como se chega lá**

1. Somar as contagens das subamostras avaliadas: **{_soma:.0f} plantas**
2. Dividir pelo número de subamostras avaliadas ({_n_aval}): média de **{_media:.2f} planta**
3. Dividir pelo estande de {est_sim:.1f} e multiplicar por 100: **{_taxa:.1f}%**

Dá no mesmo que somar tudo e dividir pelo total de plantas percorridas:
{_soma:.0f} ÷ ({_n_aval} × {est_sim:.1f}) × 100 = **{_soma / (_n_aval * est_sim) * 100:.1f}%**.
São a mesma conta escrita de dois jeitos.
""")
        if _s.eq(0).any():
            _sem_zero = _s.where(_s > 0)
            _t2 = (_sem_zero.mean() / est_sim * 100) if _sem_zero.notna().any() else np.nan
            st.caption(
                "Se os zeros fossem descartados, como o pipeline fazia antes, esta mesma "
                + (f"medição daria **{_t2:.1f}%**" if pd.notna(_t2) else "medição ficaria vazia")
                + " — por isso as taxas antigas saíam maiores.")

st.divider()

rodape()
