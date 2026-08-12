"""
pages/7_Mapa.py — Mapa de Produtividade por Região (MILHO)

Adaptado da página de soja. O desenho é o mesmo — três abas (Estado, Macro, Micro),
choropleth municipal pintado pela região, lollipop ordenado e legenda com valor.

O que muda em relação à soja:
  - fonte: carregar_multisafra() do milho, com as DUAS safras rodando ao vivo;
  - regiões: colunas macroMilho/microMilho do mesmo Excel de municípios (a soja usa
    macroSoja/microSoja — as divisões regionais das duas culturas são diferentes);
  - seletor de tipo de ensaio (Faixa/Densidade), que só existe no milho;
  - filtros no padrão das demais páginas de milho: encadeados, com busca nas listas
    longas, safra na mais recente e "nada marcado = filtro desligado".
"""
import unicodedata, re as _re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import requests
from pathlib import Path

# Raiz do projeto — resolvida uma vez no escopo global
_BASE_DIR = Path(__file__).parent.parent

from utils.theme import aplicar_tema, page_header, secao_titulo, rodape
from utils.loader import carregar_multisafra

st.set_page_config(
    page_title="Mapa · JAUM DTC Milho",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded",
)
aplicar_tema()
st.markdown("<style>.jaum-header img { height: 60px !important; }</style>", unsafe_allow_html=True)
st.markdown("""
<style>
.stTabs [data-baseweb="tab"] { font-size: 16px !important; font-weight: 600 !important; }
.stTabs [data-baseweb="tab-list"] { gap: 8px; }
</style>
""", unsafe_allow_html=True)
page_header("Mapa de Produtividade",
            "Desempenho por estado, macro e micro — selecione o híbrido.")

# ── Carregamento ──────────────────────────────────────────────────────────────
# O milho carrega as duas safras ao vivo e já empilhadas, com o de-para mestre aplicado —
# por isso basta uma chamada, diferente da soja que junta três dicts.
#
# ATENÇÃO: a analítica de milho NÃO tem a coluna `sc_ha`. Ela traz
# `produtividade_valida_sacas_ha` (e a kg equivalente), e cada página deriva o `sc_ha`.
# É a mesma normalização que a página de Densidade faz em carregar_densidade().
def _normalizar_prod(df):
    """Deriva kg_ha e sc_ha a partir das colunas do pipeline. Prioriza a produtividade
    VÁLIDA (a que passou pelas flags bloqueantes) e cai na bruta quando a válida é nula."""
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    df = df.copy()
    if "produtividade_valida_kg_ha" in df.columns:
        df["kg_ha"] = pd.to_numeric(df["produtividade_valida_kg_ha"], errors="coerce")
        if "produtividade_kg_ha" in df.columns:
            df["kg_ha"] = df["kg_ha"].fillna(
                pd.to_numeric(df["produtividade_kg_ha"], errors="coerce"))
    elif "produtividade_kg_ha" in df.columns:
        df["kg_ha"] = pd.to_numeric(df["produtividade_kg_ha"], errors="coerce")
    if "produtividade_valida_sacas_ha" in df.columns:
        df["sc_ha"] = pd.to_numeric(df["produtividade_valida_sacas_ha"], errors="coerce")
        if "produtividade_sacas_ha" in df.columns:
            df["sc_ha"] = df["sc_ha"].fillna(
                pd.to_numeric(df["produtividade_sacas_ha"], errors="coerce"))
    elif "kg_ha" in df.columns:
        df["sc_ha"] = (df["kg_ha"] / 60).round(1)
    return df


@st.cache_data(show_spinner="Carregando dados de milho...")
def _carregar_ta():
    d = carregar_multisafra()
    return {"faixa": _normalizar_prod(d.get("tabela_analitica_faixa")),
            "densidade": _normalizar_prod(d.get("tabela_analitica_densidade"))}


_TA = {k: v for k, v in _carregar_ta().items()
       if isinstance(v, pd.DataFrame) and not v.empty and "sc_ha" in v.columns}

if not _TA:
    st.error(
        "Nenhum dado com produtividade disponível.\n\n"
        "As analíticas vieram vazias ou sem as colunas de produtividade "
        "(`produtividade_valida_sacas_ha` / `produtividade_valida_kg_ha`). "
        "Verifique a conexão com os bancos das duas safras."
    )
    st.stop()

# ── GeoJSON ───────────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600, show_spinner=False)
def _carregar_geojsons_v3():
    SIGLAS = {
        "Acre":"AC","Alagoas":"AL","Amapá":"AP","Amazonas":"AM","Bahia":"BA",
        "Ceará":"CE","Distrito Federal":"DF","Espírito Santo":"ES","Goiás":"GO",
        "Maranhão":"MA","Mato Grosso":"MT","Mato Grosso do Sul":"MS",
        "Minas Gerais":"MG","Pará":"PA","Paraíba":"PB","Paraná":"PR",
        "Pernambuco":"PE","Piauí":"PI","Rio de Janeiro":"RJ",
        "Rio Grande do Norte":"RN","Rio Grande do Sul":"RS","Rondônia":"RO",
        "Roraima":"RR","Santa Catarina":"SC","São Paulo":"SP","Sergipe":"SE",
        "Tocantins":"TO",
    }
    gj_estados = None
    try:
        r = requests.get(
            "https://raw.githubusercontent.com/codeforamerica/click_that_hood/master/public/data/brazil-states.geojson",
            timeout=10)
        gj_estados = r.json()
        for f in gj_estados["features"]:
            f["properties"]["sigla"] = SIGLAS.get(f["properties"].get("name",""), "")
    except Exception:
        pass

    gj_mun = None
    try:
        # API IBGE v3 — malha municipal por UF, qualidade minima
        # Buscar cada UF separadamente e concatenar features
        UFS = ["AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS",
               "MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC",
               "SE","SP","TO"]
        _features_all = []
        # Buscar todas as UFs de uma vez — API v3 suporta BR inteiro
        # Tentar arquivo local primeiro (gerado por baixar_geojson_municipios.py)
        _path_mun = _BASE_DIR / "assets" / "municipios_br.json"
        if _path_mun.exists():
            try:
                import json as _json
                with open(_path_mun, "r", encoding="utf-8") as _fmun:
                    gj_mun = _json.load(_fmun)
            except Exception:
                gj_mun = None

        # Fallback: tentar URLs remotas
        if gj_mun is None:
            _urls_tentativa = [
                "https://servicodados.ibge.gov.br/api/v3/malhas/paises/BR?formato=application/vnd.geo+json&qualidade=minima&divisao=municipio",
                "https://raw.githubusercontent.com/luizpedropisteli/brazil-geojson/main/brazil_geo.json",
                "https://raw.githubusercontent.com/tbrugz/geodata-br/master/geojson/geojs-100-mun.json",
            ]
            for _url_t in _urls_tentativa:
                try:
                    _r = requests.get(_url_t, timeout=30)
                    if _r.status_code == 200:
                        _gj = _r.json()
                        if len(_gj.get("features",[])) > 500:
                            gj_mun = _gj
                            break
                except Exception:
                    continue

        # Normalizar propriedades — cria objeto NOVO para não mutar o cache
        _IBGE_UF = {
            "11":"RO","12":"AC","13":"AM","14":"RR","15":"PA","16":"AP","17":"TO",
            "21":"MA","22":"PI","23":"CE","24":"RN","25":"PB","26":"PE","27":"AL",
            "28":"SE","29":"BA","31":"MG","32":"ES","33":"RJ","35":"SP","41":"PR",
            "42":"SC","43":"RS","50":"MS","51":"MT","52":"GO","53":"DF",
        }
        if gj_mun:
            _features_norm = []
            for _idx, _f in enumerate(gj_mun.get("features",[])):
                _p = _f.get("properties",{})
                _cod_raw = (_p.get("code_muni") or _p.get("codarea") or _p.get("geocodigo") or
                            _p.get("id") or _p.get("CD_MUN") or _p.get("GEOCODIGO",""))
                _cod  = str(_cod_raw).split(".")[0].strip().zfill(7)
                _nome = (_p.get("name_muni") or _p.get("nome") or _p.get("name") or
                         _p.get("NM_MUN") or _p.get("NM_MUNICIP",""))
                _uf   = str(_p.get("abbrev_state") or _IBGE_UF.get(_cod[:2], "")).upper()
                # Copia a feature com propriedades novas — não muta o original
                _new_props = dict(_p)
                _new_props["ibge_norm"] = _cod
                _new_props["nome"]      = str(_nome)
                _new_props["sigla_uf"]  = _uf
                _new_props["mun_id"]    = str(_idx)  # sempre string, sempre == índice
                _features_norm.append({
                    "type":       _f.get("type","Feature"),
                    "geometry":   _f.get("geometry"),
                    "properties": _new_props,
                })
            gj_mun = {"type":"FeatureCollection","features":_features_norm}
    except Exception:
        pass

    # Tabela município → macro/micro
    df_reg = None
    _path = _BASE_DIR / "config" / "base_municipios_regioes_soja_milho.xlsx"
    if _path.exists():
        try:
            def _norm(s):
                s = unicodedata.normalize("NFKD", str(s)).encode("ascii","ignore").decode()
                return _re.sub(r"\s+","",s).upper()
            # MILHO: macroMilho/microMilho. A soja usa macroSoja/microSoja — as divisões
            # regionais das duas culturas são diferentes e não se misturam.
            df_reg = pd.read_excel(
                _path, usecols=["cidade", "siglaEstado", "ibge", "macroMilho", "microMilho"]
            ).rename(columns={"macroMilho": "regiao_macro", "microMilho": "regiao_micro"})
            df_reg["cidade_key"]       = df_reg["cidade"].apply(_norm)
            df_reg["cidade_estado_key"]= df_reg["cidade_key"] + "_" + df_reg["siglaEstado"].str.upper().fillna("")
            df_reg["ibge_norm"]        = df_reg["ibge"].astype(str).str.split(".").str[0].str.strip().str.zfill(7)
        except Exception:
            df_reg = None

    # Pré-computar mapeamento ibge→macro e ibge→micro (evita loop pesado depois)
    df_ibge_macro, df_ibge_micro = pd.DataFrame(), pd.DataFrame()
    if gj_mun is not None and df_reg is not None and "ibge_norm" in df_reg.columns:
        _rows_macro, _rows_micro, _rows_c_macro, _rows_c_micro = [], [], [], []
        for feat in gj_mun["features"]:
            _p  = feat.get("properties", {})
            _ib = _p.get("ibge_norm","")
            if not _ib: continue
            _match = df_reg[df_reg["ibge_norm"] == _ib]
            _macro = _match.iloc[0]["regiao_macro"] if not _match.empty else None
            _micro = _match.iloc[0]["regiao_micro"] if not _match.empty else None
            if _macro == "Não Identificado": _macro = None
            if _micro == "Não Identificado": _micro = None
            _rows_macro.append({"ibge_norm":_ib,"regiao":_macro})
            _rows_micro.append({"ibge_norm":_ib,"regiao":_micro})
            # Centroide
            try:
                _cf = []
                def _fl(c):
                    if isinstance(c,list) and len(c)>0:
                        if isinstance(c[0],(int,float)): _cf.append(c)
                        else:
                            for x in c: _fl(x)
                _fl(feat.get("geometry",{}).get("coordinates",[]))
                if _cf:
                    _lt = sum(c[1] for c in _cf)/len(_cf)
                    _ln = sum(c[0] for c in _cf)/len(_cf)
                    if _macro: _rows_c_macro.append({"regiao":_macro,"lat":_lt,"lon":_ln})
                    if _micro: _rows_c_micro.append({"regiao":_micro,"lat":_lt,"lon":_ln})
            except Exception:
                pass
        df_ibge_macro = pd.DataFrame(_rows_macro) if _rows_macro else pd.DataFrame()
        df_ibge_micro = pd.DataFrame(_rows_micro) if _rows_micro else pd.DataFrame()
        if _rows_c_macro:
            _dc_macro = pd.DataFrame(_rows_c_macro).groupby("regiao")[["lat","lon"]].mean().reset_index()
            df_ibge_macro = df_ibge_macro.merge(_dc_macro, on="regiao", how="left")
        if _rows_c_micro:
            _dc_micro = pd.DataFrame(_rows_c_micro).groupby("regiao")[["lat","lon"]].mean().reset_index()
            df_ibge_micro = df_ibge_micro.merge(_dc_micro, on="regiao", how="left")

    return gj_estados, gj_mun, df_reg, df_ibge_macro, df_ibge_micro

geojson_estados, geojson_mun, df_mun_regiao, _df_ibge_macro, _df_ibge_micro = _carregar_geojsons_v3()

# Garantir cache atualizado com ibge_norm
if (geojson_mun is not None and
        geojson_mun["features"] and
        "ibge_norm" not in geojson_mun["features"][0].get("properties",{})):
    _carregar_geojsons_v3.clear()
    geojson_estados, geojson_mun, df_mun_regiao, _df_ibge_macro, _df_ibge_micro = _carregar_geojsons_v3()
if df_mun_regiao is not None and "ibge_norm" not in df_mun_regiao.columns:
    _carregar_geojsons_v3.clear()
    geojson_estados, geojson_mun, df_mun_regiao, _df_ibge_macro, _df_ibge_micro = _carregar_geojsons_v3()

# Diagnóstico de carregamento — visível na sidebar
with st.sidebar:
    st.divider()
    _status_geo = []
    _status_geo.append("✅ GeoJSON estados" if geojson_estados else "❌ GeoJSON estados")
    _path_mun_check = _BASE_DIR / "assets" / "municipios_br.json"
    _src_mun = "local" if _path_mun_check.exists() else "remoto"
    _status_geo.append(f"✅ GeoJSON municípios ({_src_mun})" if geojson_mun else
                       "❌ GeoJSON municípios — execute baixar_geojson_municipios.py")
    _status_geo.append("✅ Tabela regiões" if df_mun_regiao is not None else "❌ Tabela regiões")
    for _s in _status_geo:
        st.caption(_s)


# ── Recorte e simplificação do GeoJSON ────────────────────────────────────────
# O Plotly serializa o geojson INTEIRO em CADA trace. Com um trace por região e 5.570
# municípios, o mesmo mapa de 6,7 MB ia dezenas de vezes para o navegador — foi o que
# estourou o limite de mensagem do Streamlit (506 MB).
# Duas medidas, nesta ordem de impacto:
#   1. cada trace recebe SÓ os municípios dele. A soma de todos os traces passa a ser,
#      no máximo, um geojson inteiro em vez de N cópias;
#   2. coordenadas arredondadas para 3 casas (~110 m), resolução mais que suficiente para
#      um mapa de município em escala nacional.
_PREC_COORD = 3


def _arredondar_coords(c):
    if isinstance(c, (int, float)):
        return round(c, _PREC_COORD)
    return [_arredondar_coords(x) for x in c]


@st.cache_data(ttl=3600, show_spinner=False)
def _indexar_geojson(_gj_key):
    """Devolve {ibge_norm: feature simplificada}. Cacheado: roda uma vez por sessão."""
    if geojson_mun is None:
        return {}
    idx = {}
    for feat in geojson_mun.get("features", []):
        _ib = feat.get("properties", {}).get("ibge_norm", "")
        if not _ib:
            continue
        idx[_ib] = {
            "type": "Feature",
            "properties": {"ibge_norm": _ib},
            "geometry": {"type": feat["geometry"]["type"],
                         "coordinates": _arredondar_coords(feat["geometry"]["coordinates"])},
        }
    return idx


_IDX_MUN = _indexar_geojson("v1")


def _recortar(ids):
    """FeatureCollection só com os municípios pedidos."""
    return {"type": "FeatureCollection",
            "features": [_IDX_MUN[i] for i in ids if i in _IDX_MUN]}


# ── Centroides dos estados ────────────────────────────────────────────────────
_CENTRO_ESTADOS = {
    "AC":(-9.02,-70.81),"AL":(-9.57,-36.78),"AM":(-4.14,-65.34),"AP":(1.41,-51.77),
    "BA":(-12.97,-41.73),"CE":(-5.20,-39.53),"DF":(-15.78,-47.93),"ES":(-19.57,-40.68),
    "GO":(-15.83,-49.62),"MA":(-5.43,-45.29),"MG":(-18.10,-44.38),"MS":(-20.51,-54.93),
    "MT":(-12.64,-55.42),"PA":(-3.79,-52.48),"PB":(-7.12,-36.72),"PE":(-8.38,-37.86),
    "PI":(-7.72,-42.73),"PR":(-24.89,-51.55),"RJ":(-22.25,-43.45),"RN":(-5.81,-36.59),
    "RO":(-11.43,-62.84),"RR":(2.05,-61.38),"RS":(-30.17,-53.50),"SC":(-27.25,-50.22),
    "SE":(-10.57,-37.45),"SP":(-22.25,-48.65),"TO":(-10.18,-48.33),
}

# Paleta para estados — 27 cores distintas
_CORES_ESTADOS = [
    "#1F4E79","#C0392B","#27AE60","#8E44AD","#E67E22",
    "#16A085","#884EA0","#2980B9","#D35400","#1ABC9C",
    "#7D3C98","#CB4335","#1A5276","#117A65","#784212",
    "#1B4F72","#6C3483","#0E6655","#922B21","#4A235A",
    "#1D8348","#2E4057","#A93226","#196F3D","#4D5656",
    "#1A252F","#873600",
]
# Paleta Macro — 5 regiões, cores maximamente distintas
# MACRO I=azul, II=verde, III=vermelho, IV=laranja, V=roxo
_CORES_MACRO = [
    "#1565C0",  # azul forte
    "#2E7D32",  # verde escuro
    "#C62828",  # vermelho escuro
    "#E65100",  # laranja escuro
    "#6A1B9A",  # roxo escuro
    "#00838F",  # ciano escuro
    "#F9A825",  # amarelo escuro
    "#37474F",  # cinza azulado
    "#AD1457",  # rosa escuro
    "#558B2F",  # verde oliva
]
_CORES_CAT = _CORES_MACRO  # fallback genérico

_GEO_LAYOUT = dict(
    scope="south america",
    showland=True,  landcolor="#FFFFFF",
    showocean=True, oceancolor="#FFFFFF",
    showlakes=False, showrivers=False,
    showcountries=True, countrycolor="#D1D5DB",
    showsubunits=True,  subunitcolor="#E5E7EB",
    center=dict(lat=-15, lon=-52),
    bgcolor="#FFFFFF",
)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        '<p style="font-size:11px;font-weight:600;color:#6B7280;text-transform:uppercase;'
        'letter-spacing:0.05em;padding:0.5rem;">Filtros</p>',
        unsafe_allow_html=True,
    )
    if st.button("🔄 Limpar filtros", use_container_width=True, key="mp_btn_limpar"):
        for k in list(st.session_state.keys()):
            if str(k).startswith("mp_") and not str(k).endswith("_btn_limpar"):
                del st.session_state[k]
        st.rerun()

    # Tipo de ensaio — só existe no milho. Faixa e Densidade são tabelas distintas e
    # NÃO se misturam: a de Densidade tem várias populações do mesmo híbrido no mesmo
    # local, o que puxaria a média do mapa sem que se veja o motivo.
    _tipos_ok = [k for k, v in _TA.items() if v is not None and not v.empty]
    _rot_tipo = {"faixa": "Faixa", "densidade": "Densidade"}
    _tipo_sel = st.radio("Tipo de ensaio", options=_tipos_ok,
                         format_func=lambda t: _rot_tipo.get(t, t),
                         key="mp_tipo", horizontal=True)
    ta = _TA[_tipo_sel].copy()
    ta = ta[ta["sc_ha"].notna() & (ta["sc_ha"] > 0)].copy()

    def _chk(opcoes, prefix, defaults=None):
        """Nada marcado = filtro desligado (padrão do painel de milho)."""
        sel = []
        for o in opcoes:
            _key = f"{prefix}_{o}"
            if _key in st.session_state:
                marcado = st.checkbox(str(o), key=_key)
            else:
                marcado = st.checkbox(str(o), value=(defaults is None or o in defaults),
                                      key=_key)
            if marcado:
                sel.append(o)
        return sel

    def _busca(opcoes, prefix):
        """Busca textual para listas longas; a seleção sobrevive ao filtro da busca."""
        if f"{prefix}_reset" not in st.session_state:
            st.session_state[f"{prefix}_reset"] = 0
        r = st.session_state[f"{prefix}_reset"]
        _b = st.text_input("Buscar", value="", key=f"mp_busca_{prefix}",
                           placeholder="Digite para filtrar...")
        _vis = [c for c in opcoes if _b.strip().lower() in str(c).lower()] if _b.strip() else opcoes
        if st.button("Limpar seleção", key=f"{prefix}_limpar", use_container_width=True):
            for o in opcoes:
                st.session_state.pop(f"{prefix}_chk_{r}_{o}", None)
            st.session_state[f"{prefix}_reset"] = r + 1
            st.rerun()
        for c in _vis:
            st.checkbox(str(c), key=f"{prefix}_chk_{r}_{c}")
        sel = [o for o in opcoes if st.session_state.get(f"{prefix}_chk_{r}_{o}", False)]
        return sel or opcoes

    # Safra — padrão na mais recente
    def _ano(x):
        try:
            return int(str(x).split("/")[0])
        except Exception:
            return -1
    safras_all = sorted(ta["safra"].dropna().unique().tolist(), key=_ano, reverse=True) \
        if "safra" in ta.columns else []
    safra_def = safras_all[:1]
    with st.expander("📅 Safra", expanded=True):
        safras_sel = _chk(safras_all, "mp_safra", defaults=safra_def)
    ta_f = ta[ta["safra"].isin(safras_sel)] if safras_sel else ta

    # Demais filtros, encadeados. Seleção vazia NÃO zera a base — só não filtra.
    for _col, _lab, _pref, _usa_busca in [
        ("status_material", "🏷️ Status", "mp_status", False),
        ("estado_sigla", "🏛️ Estado", "mp_estado", False),
        ("regiao_macro", "🗺️ Macro", "mp_macro", False),
        ("regiao_micro", "📍 Micro", "mp_micro", True),
        ("cidade_nome", "🏙️ Cidade", "mp_cid", True),
        ("nomeFazenda", "🚜 Fazenda", "mp_faz", True),
    ]:
        if _col not in ta_f.columns:
            continue
        _ops = sorted(ta_f[_col].dropna().unique().tolist())
        if not _ops:
            continue
        with st.expander(_lab, expanded=False):
            _sel = _busca(_ops, _pref) if _usa_busca else _chk(_ops, _pref)
        _sel_ok = [v for v in (_sel or []) if v in _ops]
        if _sel_ok:
            ta_f = ta_f[ta_f[_col].isin(_sel_ok)]

    # Híbrido — seleção única, é o eixo da página
    with st.expander("🌽 Híbrido", expanded=True):
        cults_all = sorted(ta_f["dePara"].dropna().unique().tolist()) \
            if "dePara" in ta_f.columns else []
        if not cults_all:
            st.warning("Nenhum híbrido disponível nos filtros ativos.")
            cultivar_sel = None
        else:
            cultivar_sel = st.selectbox("Híbrido", cults_all, key="mp_cult",
                                        label_visibility="collapsed")

    st.divider()
    # o restante do código já sabia lidar com as três métricas (ver col_val); faltava o
    # widget. Mediana protege contra região com poucas parcelas e um valor extremo;
    # N parcelas responde "onde temos cobertura", que é outra pergunta legítima do mapa.
    metrica = st.radio("Métrica", ["Média sc/ha", "Mediana sc/ha", "N parcelas"],
                       key="mp_metrica",
                       help="Mediana é mais resistente a uma parcela fora da curva. "
                            "N parcelas mostra a cobertura do ensaio, não o desempenho.")

if cultivar_sel is None:
    st.info("Selecione um híbrido.")
    st.stop()

df_cult = ta_f[ta_f["dePara"] == cultivar_sel].copy()
if df_cult.empty:
    st.info("Nenhum dado para os filtros selecionados.")
    st.stop()

col_val   = {"Média sc/ha":"media","Mediana sc/ha":"mediana","N parcelas":"n"}[metrica]
label_val = metrica

_N_MIN_REGIAO = 5   # abaixo disso a região não pode ser anunciada como a melhor


def _melhor_confiavel(df_agg, col):
    """Região com o maior valor, ignorando as apoiadas em poucas parcelas.

    Sem isso, uma região com duas parcelas lidera por acaso e vira manchete no título da
    aba — o mesmo problema que já apareceu na tabela de grupos e no mapa de calor.
    Devolve (regiao, fragil): fragil=True quando TODAS estão abaixo do mínimo e não houve
    como escolher com segurança.
    """
    _ok = df_agg[df_agg["n"] >= _N_MIN_REGIAO]
    if not _ok.empty:
        return _ok.loc[_ok[col].idxmax(), "regiao"], False
    return df_agg.loc[df_agg[col].idxmax(), "regiao"], True


def _agg(df, col_grupo):
    r = (df.groupby(col_grupo, dropna=True)["sc_ha"]
         .agg(media="mean", mediana="median", n="count")
         .reset_index()
         .rename(columns={col_grupo:"regiao"}))
    r["media"]   = r["media"].round(1)
    r["mediana"] = r["mediana"].round(1)
    return r

def _fmt(v):
    if pd.isna(v): return "—"
    return f"{v:.0f}" if col_val != "n" else str(int(v))

def _cor_texto(hex_cor):
    """Retorna preto ou branco dependendo da luminância da cor de fundo (WCAG)."""
    hex_cor = hex_cor.lstrip("#")
    if len(hex_cor) != 6:
        return "#FFFFFF"
    r, g, b = int(hex_cor[0:2],16), int(hex_cor[2:4],16), int(hex_cor[4:6],16)
    def _lin(c):
        c = c / 255
        return c/12.92 if c <= 0.03928 else ((c+0.055)/1.055)**2.4
    lum = 0.2126*_lin(r) + 0.7152*_lin(g) + 0.0722*_lin(b)
    return "#1A1A1A" if lum > 0.35 else "#FFFFFF"


def _lollipop(df_agg, cor_map, melhor, label_nivel):
    """Lollipop horizontal ordenado por média sc/ha."""
    _df = df_agg[df_agg["media"].notna()].sort_values("media", ascending=True).reset_index(drop=True)
    if _df.empty:
        return
    _fig = go.Figure()
    for i, row in _df.iterrows():
        _cor  = cor_map.get(row["regiao"], "#2976B6")
        _dest = row["regiao"] == melhor
        # região com poucas parcelas: valor em cinza e ponto oco. Sem isso, uma região de
        # UMA parcela aparece no topo do gráfico com o mesmo peso visual de uma com trinta.
        _frag = row["n"] < _N_MIN_REGIAO
        _abr  = str(row["regiao"]).replace("MACRO ","M ").replace("REC ","")
        _cor_num = "#9CA3AF" if _frag else "#1A1A1A"
        _lbl  = (f"<b style='font-size:14px;color:{_cor_num}'>{row['media']:.1f}</b>"
                 f"<b style='font-size:13px;color:{_cor_num}'> ({int(row['n'])})</b>"
                 + ("<span style='font-size:12px;color:#B45309'> ⚠</span>" if _frag else ""))
        _fig.add_shape(type="line", x0=0, x1=row["media"], y0=i, y1=i,
                       line=dict(color=_cor, width=3 if _dest else 1.5,
                                 dash="dot" if _frag else None))
        _fig.add_trace(go.Scatter(
            x=[row["media"]], y=[i], mode="markers",
            marker=dict(color="#FFFFFF" if _frag else _cor,
                        size=14 if _dest else 10,
                        line=dict(color=_cor, width=2)),
            hovertemplate=f"<b>{row['regiao']}</b><br>{row['media']:.1f} sc/ha<br>{int(row['n'])} parcelas<extra></extra>",
            showlegend=False,
        ))
        _fig.add_annotation(
            x=row["media"], y=i, text=_lbl,
            xanchor="left", yanchor="middle",
            showarrow=False, xshift=10,
            font=dict(size=14, color="#1A1A1A", weight="bold"),
        )
    # Linha de média geral
    _media_geral = _df["media"].mean()
    _fig.add_shape(type="line",
        x0=_media_geral, x1=_media_geral,
        y0=-0.5, y1=len(_df)-0.5,
        line=dict(color="#374151", width=1.5, dash="dot"),
    )
    _fig.add_annotation(
        x=_media_geral, y=len(_df)-0.5,
        text=f"<b>Média: {_media_geral:.1f}</b>",
        xanchor="left", yanchor="bottom",
        showarrow=False, xshift=5,
        font=dict(size=12, color="#374151", weight="bold"),
    )

    _fig.update_layout(
        height=max(280, len(_df) * 34),
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
        font=dict(family="Helvetica Neue, sans-serif", color="#1A1A1A"),
        showlegend=False,
        margin=dict(t=10, b=40, l=10, r=130),   # espaço para o rótulo do maior valor
        xaxis=dict(
            # 1,18 em vez de 1,3: a folga antiga deixava um terço do gráfico vazio e
            # mesmo assim o rótulo do maior valor era cortado (ver margem r abaixo)
            range=[0, _df["media"].max() * 1.18],
            title=dict(text="<b>sc/ha médio</b>", font=dict(size=14, color="#1A1A1A", weight="bold")),
            tickfont=dict(size=12, color="#1A1A1A", weight="bold"),
            showgrid=True, gridcolor="#F0F0F0", zeroline=False,
        ),
        yaxis=dict(
            tickmode="array", tickvals=list(range(len(_df))),
            ticktext=[
                f"<b>{str(r).replace('MACRO ','M ').replace('REC ','')}</b>"
                if r == melhor
                else f"<b>{str(r).replace('MACRO ','M ').replace('REC ','')}</b>"
                for r in _df["regiao"]
            ],
            tickfont=dict(size=12, color="#1A1A1A", weight="bold"),
            showgrid=False, zeroline=False,
        ),
    )
    st.plotly_chart(_fig, use_container_width=True)

def _exportar(df_agg, nivel, key):
    """Exporta o que está na tela: uma linha por região, com as três métricas."""
    import io
    _out = (df_agg.rename(columns={"regiao": nivel, "media": "Média sc/ha",
                                   "mediana": "Mediana sc/ha", "n": "N parcelas"})
            .sort_values("Média sc/ha", ascending=False))
    _out.insert(1, "Híbrido", cultivar_sel)
    _out.insert(2, "Tipo de ensaio", _rot_tipo.get(_tipo_sel, _tipo_sel))
    _out.insert(3, "Safras", ", ".join(safras_sel) if safras_sel else "todas")
    _buf = io.BytesIO()
    with pd.ExcelWriter(_buf, engine="openpyxl") as _w:
        _out.to_excel(_w, index=False, sheet_name=nivel[:28])
    st.download_button(f"⬇️ Exportar {nivel.lower()}", data=_buf.getvalue(),
                       file_name=f"mapa_{nivel.lower()}_{cultivar_sel}.xlsx",
                       mime=("application/vnd.openxmlformats-officedocument."
                             "spreadsheetml.sheet"), key=key)


def _legenda(df_agg, cor_map, regioes_todas, melhor_reg, abrev=False):
    """Renderiza legenda colorida com valor."""
    regioes_com_dado = set(df_agg["regiao"].dropna())
    df_ord = df_agg.sort_values(col_val, ascending=False)
    for _, row in df_ord.iterrows():
        cor  = cor_map.get(row["regiao"], "#94A3B8")
        best = row["regiao"] == melhor_reg
        nome = str(row["regiao"]).replace("MACRO ","M ").replace("REC ","") if abrev else row["regiao"]
        val  = _fmt(row[col_val])
        borda = "2px solid #1A1A1A" if best else "1px solid #E5E7EB"
        fw    = "700" if best else "400"
        st.markdown(f"""
<div style="display:flex;align-items:center;gap:8px;margin-bottom:5px;
            padding:5px 8px;border-radius:6px;background:#F8FAFC;border:{borda}">
  <div style="width:12px;height:12px;border-radius:2px;background:{cor};flex-shrink:0;"></div>
  <div style="flex:1;font-size:11px;color:#1A1A1A;font-weight:{fw};">{nome}</div>
  <div style="font-size:12px;font-weight:700;color:{cor};">{val}</div>
</div>""", unsafe_allow_html=True)
    # Regiões sem dado
    for r in sorted(set(regioes_todas) - regioes_com_dado):
        nome = str(r).replace("MACRO ","M ").replace("REC ","") if abrev else r
        st.markdown(f"""
<div style="display:flex;align-items:center;gap:8px;margin-bottom:5px;
            padding:5px 8px;border-radius:6px;background:#F8FAFC;border:1px solid #E5E7EB">
  <div style="width:12px;height:12px;border-radius:2px;background:#E5E7EB;flex-shrink:0;"></div>
  <div style="flex:1;font-size:11px;color:#9CA3AF;">{nome}</div>
  <div style="font-size:12px;color:#9CA3AF;">—</div>
</div>""", unsafe_allow_html=True)

with st.popover("ℹ️ Como ler esta página", use_container_width=False):
    st.markdown("""
**O mapa mostra UM híbrido por vez.** Cada aba agrega a produtividade dele num nível
geográfico diferente: por estado, por macrorregião ou por microrregião.

**Os controles da barra lateral**

- **Tipo de ensaio** — Faixa ou Densidade. São tabelas separadas e **não se misturam**: a de
  Densidade tem várias populações do mesmo híbrido no mesmo local, o que mexeria na média do
  mapa sem que se veja o motivo.
- **Safra** — abre na mais recente. Marcar mais de uma junta as safras na mesma média.
- **Status, Estado, Macro, Micro, Cidade e Fazenda** são encadeados: escolher um estado reduz
  as cidades disponíveis. **Filtro sem nada marcado está desligado** e mostra tudo.
- **Híbrido** — seleção única, é o eixo da página.

**Como o número é calculado.** Para cada região, a métrica escolhida na barra lateral sobre as
parcelas daquele híbrido:

- **Média sc/ha** — a média simples. É o padrão.
- **Mediana sc/ha** — o valor do meio. Resiste a uma parcela fora da curva, o que importa em
  região com poucas parcelas.
- **N parcelas** — quantas parcelas existem ali. Responde "onde temos cobertura", não desempenho.

O número entre parênteses no gráfico é sempre quantas parcelas sustentam aquele valor.

**"Melhor em X" ignora regiões com menos de 5 parcelas.** Sem essa trava, uma região com duas
parcelas lidera por acaso e vira manchete. Quando nenhuma região atinge o mínimo, o subtítulo diz
isso em vez de eleger uma, e aparece um aviso abaixo do gráfico.

**Como ler o mapa.** A cor identifica a **região**, não o desempenho — regiões diferentes têm
cores diferentes para você localizar onde está cada uma. Quem mostra o desempenho é o gráfico
ao lado, ordenado da menor para a maior produtividade, com uma linha pontilhada na média geral.
Regiões sem resultado para aquele híbrido aparecem em tom desbotado.

**Dois cuidados**

**Média de poucas parcelas oscila muito.** Uma região com duas parcelas pode liderar por acaso.
Confira sempre o número entre parênteses antes de concluir que uma região é melhor.

**As regiões de milho não são as mesmas da soja.** O mapa usa as colunas `macroMilho` e
`microMilho` do cadastro de municípios. Comparar diretamente com o mapa da soja não faz sentido.
""")

# ══════════════════════════════════════════════════════════════════════════════
# ABAS — Estado | Macro | Micro
# ══════════════════════════════════════════════════════════════════════════════
tab_estado, tab_macro, tab_micro = st.tabs(["Estado", "Macro", "Micro"])

# ── ABA ESTADO ────────────────────────────────────────────────────────────────
with tab_estado:
    df_est = _agg(df_cult, "estado_sigla")
    if df_est.empty:
        st.info("Sem dados por estado.")
    else:
        _melhor_est, _est_fragil = _melhor_confiavel(df_est, col_val)
        _todos_est  = list(_CENTRO_ESTADOS.keys())
        # Cores distintas para cada estado
        _cor_est    = {r: _CORES_ESTADOS[i % len(_CORES_ESTADOS)]
                       for i, r in enumerate(sorted(_todos_est))}

        _sub_est = (f"{cultivar_sel} — melhor em {_melhor_est}" if not _est_fragil
                    else f"{cultivar_sel} — nenhum estado com {_N_MIN_REGIAO}+ parcelas")
        secao_titulo("Estado", _sub_est,
                     f"{metrica} · {_rot_tipo.get(_tipo_sel, _tipo_sel)} · "
                 f"{', '.join(safras_sel) if safras_sel else 'todas as safras'}")

        col_m, _dummy = st.columns([4, 0.001])
        with col_m:
            if geojson_estados:
                # Choropleth com todos os estados
                _df_full = pd.DataFrame([{"sigla":s} for s in _todos_est])
                _df_full = _df_full.merge(df_est.rename(columns={"regiao":"sigla"}),
                                          on="sigla", how="left")
                _fig = go.Figure()
                # Estados com dado
                for _, row in df_est.iterrows():
                    _cor = _cor_est.get(row["regiao"], "#E5E7EB")
                    _fig.add_trace(go.Choropleth(
                        geojson=geojson_estados,
                        locations=[row["regiao"]],
                        z=[1], featureidkey="properties.sigla",
                        colorscale=[[0,_cor],[1,_cor]],
                        showscale=False, zmin=0, zmax=1,
                        marker_line_color="#FFFFFF", marker_line_width=0.8,
                        hovertemplate=f"<b>{row['regiao']}</b><br>{label_val}: {_fmt(row[col_val])}<br>N: {int(row['n'])}<extra></extra>",
                        showlegend=False,
                    ))
                # Estados sem dado
                _sem_est = [s for s in _todos_est if s not in df_est["regiao"].values]
                if _sem_est:
                    _fig.add_trace(go.Choropleth(
                        geojson=geojson_estados, locations=_sem_est, z=[1]*len(_sem_est),
                        featureidkey="properties.sigla",
                        colorscale=[[0,"#E5E7EB"],[1,"#E5E7EB"]],
                        showscale=False, zmin=0, zmax=1,
                        marker_line_color="#FFFFFF", marker_line_width=0.8,
                        hoverinfo="skip", showlegend=False,
                    ))
                # Siglas + valores
                _lats = [_CENTRO_ESTADOS[s][0] for s in _df_full["sigla"]]
                _lons = [_CENTRO_ESTADOS[s][1] for s in _df_full["sigla"]]
                # Valor sobre estados com dado
                _df_com = _df_full[_df_full[col_val].notna()]
                # Cor do texto baseada no contraste com a cor do estado
                _texts_est = []
                _lats_est  = []
                _lons_est  = []
                for _, _er in _df_com.iterrows():
                    _cor_e = _cor_est.get(_er["sigla"], "#E5E7EB")
                    _txt_c = _cor_texto(_cor_e)
                    _fig.add_trace(go.Scattergeo(
                        lat=[_CENTRO_ESTADOS[_er["sigla"]][0]],
                        lon=[_CENTRO_ESTADOS[_er["sigla"]][1]],
                        mode="text",
                        text=[f"<b>{_fmt(_er[col_val])} sc</b>"],
                        textfont=dict(size=22, color=_txt_c,
                                      family="Helvetica Neue, sans-serif"),
                        hoverinfo="skip", showlegend=False,
                    ))
                # Legenda no mapa — traces invisíveis com showlegend=True
                _df_leg_est = df_est.sort_values(col_val, ascending=False)
                for _i, _lr in _df_leg_est.iterrows():
                    _c = _cor_est.get(_lr["regiao"], "#94A3B8")
                    _lbl = f"{_lr['regiao']}  {_fmt(_lr[col_val])} sc"
                    _fig.add_trace(go.Scattergeo(
                        lat=[None], lon=[None], mode="markers",
                        marker=dict(color=_c, size=10, symbol="square"),
                        name=_lbl, showlegend=True,
                    ))
                _fig.update_geos(fitbounds="geojson", visible=False, bgcolor="#FFFFFF")
                _fig.update_layout(
                    height=680, margin=dict(t=0,b=0,l=0,r=120),
                    paper_bgcolor="#FFFFFF",
                    font=dict(family="Helvetica Neue, sans-serif"),
                    legend=dict(
                        x=1.01, y=1, xanchor="left", yanchor="top",
                        bgcolor="rgba(255,255,255,0.9)",
                        bordercolor="#E5E7EB", borderwidth=1,
                        font=dict(size=10), itemsizing="constant",
                        tracegroupgap=2,
                    ),
                )
                st.plotly_chart(_fig, use_container_width=True)

        # Lollipop Estado
        st.divider()
        st.markdown('<p style="font-size:12px;font-weight:600;color:#6B7280;text-transform:uppercase;'
                    'letter-spacing:0.07em;">Produtividade Média por Estado</p>', unsafe_allow_html=True)
        _lollipop(df_est, _cor_est, _melhor_est, "Estado")
        if _est_fragil:
            st.warning(f"Nenhum estado tem {_N_MIN_REGIAO} parcelas ou mais para "
                       f"**{cultivar_sel}** neste recorte. As médias oscilam muito com poucas "
                       f"parcelas — leia os números, mas não eleja um estado por eles.")
        _exportar(df_est, "Estado", key="mp_xlsx_estado")


# ── Helpers de geocódigo IBGE — usados em _mapa_regiao ───────────────────────
_IBGE_UF_MAP = {
    "11":"RO","12":"AC","13":"AM","14":"RR","15":"PA","16":"AP","17":"TO",
    "21":"MA","22":"PI","23":"CE","24":"RN","25":"PB","26":"PE","27":"AL",
    "28":"SE","29":"BA","31":"MG","32":"ES","33":"RJ","35":"SP","41":"PR",
    "42":"SC","43":"RS","50":"MS","51":"MT","52":"GO","53":"DF",
}

def _uf_do_geocod(p):
    """Extrai sigla do estado a partir das propriedades do GeoJSON."""
    # abbrev_state está disponível neste GeoJSON
    if p.get("abbrev_state"):
        return str(p["abbrev_state"]).upper()
    _cod = str(p.get("codarea") or p.get("geocodigo") or p.get("code_muni") or
               p.get("id") or p.get("CD_MUN") or p.get("GEOCODIGO") or
               p.get("sigla_uf","")).split(".")[0]
    if len(_cod) >= 2 and _cod[:2].isdigit():
        return _IBGE_UF_MAP.get(_cod[:2], "")
    if len(_cod) == 2 and _cod.isalpha():
        return _cod.upper()
    return ""

def _geocod_ibge(p):
    """Extrai geocódigo IBGE de 7 dígitos das propriedades do GeoJSON.
    Suporta formato float ('3500105.0') e inteiro ('3500105').
    """
    raw = (p.get("codarea") or p.get("geocodigo") or p.get("id") or
           p.get("CD_MUN") or p.get("GEOCODIGO") or p.get("code_muni") or "")
    # Remove parte decimal se vier como float string (ex: '3500105.0' → '3500105')
    s = str(raw).strip().split(".")[0].zfill(7)
    return s if s.isdigit() and len(s) >= 6 else ""

# ── FUNÇÃO MAPA MACRO/MICRO ───────────────────────────────────────────────────

def _processar_geojson_mun(_cache_key, col_grupo):
    """Retorna (df_coords, df_mun_map) — usa dados pré-computados no carregamento."""
    df_pre = _df_ibge_macro if col_grupo == "regiao_macro" else _df_ibge_micro
    if df_pre.empty:
        return pd.DataFrame(), pd.DataFrame()
    df_coords = pd.DataFrame()
    if "lat" in df_pre.columns:
        df_coords = (df_pre[df_pre["lat"].notna()]
                     .groupby("regiao")[["lat","lon"]].mean().reset_index())
    return df_coords, df_pre[["ibge_norm","regiao"]]

def _mapa_regiao(col_grupo, label_nivel, fator_cor=1.0):
    df_reg = _agg(df_cult, col_grupo)
    if df_reg.empty:
        st.info(f"Sem dados por {label_nivel}.")
        return

    _melhor_reg, _reg_fragil = _melhor_confiavel(df_reg, col_val)
    # Regiões para colorir = TODAS da tabela de municípios (não só as com resultado)
    # Assim municípios de regiões sem resultado ficam com cor desbotada em vez de cinza
    if df_mun_regiao is not None and col_grupo in df_mun_regiao.columns:
        _regioes_all = sorted(df_mun_regiao[col_grupo].dropna().unique().tolist())
        _regioes_all = [r for r in _regioes_all if r and r != "Não Identificado"]
    else:
        _regioes_all = sorted(ta[col_grupo].dropna().unique().tolist())
    _regioes_com = set(df_reg["regiao"].dropna())
    _cor_map     = {r: _CORES_MACRO[i % len(_CORES_MACRO)]
                    for i, r in enumerate(sorted(_regioes_all))}

    _sub_reg = (f"{cultivar_sel} — melhor em {_melhor_reg}" if not _reg_fragil
                else f"{cultivar_sel} — nenhuma região com {_N_MIN_REGIAO}+ parcelas")
    secao_titulo(label_nivel, _sub_reg,
                 f"{metrica} · {_rot_tipo.get(_tipo_sel, _tipo_sel)} · "
                 f"{', '.join(safras_sel) if safras_sel else 'todas as safras'}")

    # ── Centroides e mapeamento IBGE — cacheados ─────────────────────────────
    _cache_result  = _processar_geojson_mun(col_grupo, col_grupo)
    df_coords      = _cache_result[0]
    _df_mun_map    = _cache_result[1]

    # fallback: média das fazendas
    if df_coords.empty and "latitude" in df_cult.columns and "longitude" in df_cult.columns:
        _df_valid = df_cult[
            df_cult["latitude"].between(-35, 6) &
            df_cult["longitude"].between(-74, -28)
        ]
        df_coords = (
            _df_valid.groupby(col_grupo, dropna=True)
            .agg(lat=("latitude", "mean"), lon=("longitude", "mean"))
            .reset_index()
            .rename(columns={col_grupo: "regiao"})
        )

    col_m, _dummy2 = st.columns([4, 0.001])
    with col_m:
        _fig = go.Figure()

        if geojson_mun is not None and df_mun_regiao is not None:
            # Converter DataFrame em dicionário reg→lista de ibge
            if not _df_mun_map.empty:
                _ibge_por_reg = (_df_mun_map[_df_mun_map["regiao"].notna()]
                                 .groupby("regiao")["ibge_norm"].apply(list).to_dict())
                _ibge_sem = _df_mun_map[_df_mun_map["regiao"].isna()]["ibge_norm"].tolist()
            else:
                _ibge_por_reg, _ibge_sem = {}, []

            def _desbotar(hex_cor, fator=0.6):
                hex_cor = hex_cor.lstrip("#")
                if len(hex_cor) != 6:
                    return "#E5E7EB"
                r = int(int(hex_cor[0:2], 16) * fator + 255 * (1 - fator))
                g = int(int(hex_cor[2:4], 16) * fator + 255 * (1 - fator))
                b = int(int(hex_cor[4:6], 16) * fator + 255 * (1 - fator))
                return f"#{r:02X}{g:02X}{b:02X}"

            _regioes_com_resultado = set(df_reg["regiao"].dropna())
            _regioes_validas = [r for r in _regioes_all if r and r != "Não Identificado"]

            for _reg in _regioes_validas:
                _ids = _ibge_por_reg.get(_reg, [])
                if not _ids:
                    continue

                _cor_base = _cor_map.get(_reg, "#E5E7EB")
                _cor = _desbotar(_cor_base, fator=fator_cor) if _reg in _regioes_com_resultado else _desbotar(_cor_base, fator=fator_cor * 0.6)

                _val_row = df_reg[df_reg["regiao"] == _reg]
                _val_txt = _fmt(_val_row.iloc[0][col_val]) if not _val_row.empty else "—"
                _n_txt = str(int(_val_row.iloc[0]["n"])) if not _val_row.empty else "0"

                _fig.add_trace(go.Choropleth(
                    geojson=_recortar(_ids),
                    locations=_ids,
                    z=[1] * len(_ids),
                    featureidkey="properties.ibge_norm",
                    colorscale=[[0, _cor], [1, _cor]],
                    showscale=False,
                    zmin=0,
                    zmax=1,
                    marker_line_color="#FFFFFF",
                    marker_line_width=0.2,
                    hovertemplate=f"<b>{_reg}</b><br>{label_val}: {_val_txt}<br>N: {_n_txt}<extra></extra>",
                    showlegend=False,
                ))

            # municípios sem região
            if _ibge_sem:
                _fig.add_trace(go.Choropleth(
                    geojson=_recortar(_ibge_sem),
                    locations=_ibge_sem,
                    z=[1] * len(_ibge_sem),
                    featureidkey="properties.ibge_norm",
                    colorscale=[[0, "#E5E7EB"], [1, "#E5E7EB"]],
                    showscale=False,
                    zmin=0,
                    zmax=1,
                    marker_line_color="#E5E7EB",
                    marker_line_width=0.1,
                    hoverinfo="skip",
                    showlegend=False,
                ))

        elif geojson_estados:
            # ── fallback por estado ─────────────────────────────────────────
            _est_reg = (
                ta[ta["estado_sigla"].notna() & ta[col_grupo].notna()]
                .groupby("estado_sigla")[col_grupo]
                .agg(lambda x: x.value_counts().index[0])
                .reset_index()
                .rename(columns={col_grupo: "regiao"})
            )

            for _reg in _regioes_all:
                _ests = _est_reg[_est_reg["regiao"] == _reg]["estado_sigla"].tolist()
                if not _ests:
                    continue

                _cor = _cor_map.get(_reg, "#E5E7EB")
                _val_row = df_reg[df_reg["regiao"] == _reg]
                _val_txt = _fmt(_val_row.iloc[0][col_val]) if not _val_row.empty else "—"
                _n_txt = str(int(_val_row.iloc[0]["n"])) if not _val_row.empty else "0"

                _fig.add_trace(go.Choropleth(
                    geojson=geojson_estados,
                    locations=_ests,
                    z=[1] * len(_ests),
                    featureidkey="properties.sigla",
                    colorscale=[[0, _cor], [1, _cor]],
                    showscale=False,
                    zmin=0,
                    zmax=1,
                    marker_line_color="#FFFFFF",
                    marker_line_width=0.8,
                    hovertemplate=f"<b>{_reg}</b><br>{label_val}: {_val_txt}<br>N: {_n_txt}<extra></extra>",
                    showlegend=False,
                ))

        # contorno dos estados — só borda transparente, sem pintar cinza
        if geojson_estados:
            _todos_siglas = list(_CENTRO_ESTADOS.keys())

            _fig.add_trace(go.Choropleth(
                geojson=geojson_estados,
                locations=_todos_siglas,
                z=[1] * len(_todos_siglas),
                featureidkey="properties.sigla",
                colorscale=[[0, "rgba(255,255,255,0)"], [1, "rgba(255,255,255,0)"]],
                showscale=False,
                zmin=0,
                zmax=1,
                marker_line_color="#FFFFFF",
                marker_line_width=2.0,
                hoverinfo="skip",
                showlegend=False,
            ))

        # ── Rótulos no mapa ──────────────────────────────────────────────────
        # Rotular TODAS as regiões torna a aba Micro ilegível: com 14 microrregiões
        # concentradas no Centro-Oeste, os nomes se empilham e nenhum se lê. A regra:
        #   - só entram regiões com dado e com n suficiente (as frágeis viram só hover);
        #   - no máximo _MAX_LABELS, escolhidas pelo valor, mais a melhor;
        #   - fonte menor e halo branco, para o texto sobreviver sobre o mapa colorido.
        # O que sai do mapa não some: continua no hover e no gráfico ao lado.
        _MAX_LABELS = 6
        if not df_coords.empty:
            _df_lbl = (df_reg.merge(df_coords, on="regiao", how="left")
                       .dropna(subset=["lat", "lon"]))
            _df_lbl = _df_lbl[_df_lbl["n"] >= _N_MIN_REGIAO]
            _df_lbl = _df_lbl.sort_values(col_val, ascending=False)

            # Limitar a quantidade não basta: no Centro-Oeste as microrregiões são vizinhas e
            # os centroides caem a poucos graus um do outro, então três rótulos válidos ainda
            # se sobrepõem. Aqui entra a distância mínima: percorre do maior valor para o menor
            # e só aceita um rótulo se ele estiver longe o bastante de todos os já aceitos.
            # Em área dispersa isso não tira nada; em área densa tira o suficiente para ler.
            # A melhor região entra sempre, e é a âncora — os outros cedem lugar a ela.
            _DIST_MIN_GRAUS = 4.5

            _aceitos = []
            _linha_melhor = _df_lbl[_df_lbl["regiao"] == _melhor_reg]
            if not _linha_melhor.empty:
                _aceitos.append(_linha_melhor.iloc[0])

            for _, _cand in _df_lbl.iterrows():
                if len(_aceitos) >= _MAX_LABELS:
                    break
                if any(_cand["regiao"] == _a["regiao"] for _a in _aceitos):
                    continue
                _longe = all(
                    ((_cand["lat"] - _a["lat"]) ** 2 + (_cand["lon"] - _a["lon"]) ** 2) ** 0.5
                    >= _DIST_MIN_GRAUS for _a in _aceitos)
                if _longe:
                    _aceitos.append(_cand)

            _df_lbl = (pd.DataFrame(_aceitos) if _aceitos
                       else _df_lbl.iloc[0:0])

            for _, _row in _df_lbl.iterrows():
                _abr = str(_row["regiao"]).replace("MACRO ", "M ").replace("REC ", "")
                # nomes de micro são longos (TB13_Norte_Mato_Grosso): quebra em duas linhas
                if len(_abr) > 16 and "_" in _abr:
                    _p = _abr.split("_")
                    _meio = len(_p) // 2 + len(_p) % 2
                    _abr = " ".join(_p[:_meio]) + "<br>" + " ".join(_p[_meio:])
                else:
                    _abr = _abr.replace("_", " ")
                _val = _fmt(_row[col_val])
                _dest = _row["regiao"] == _melhor_reg

                _fig.add_trace(go.Scattergeo(
                    lat=[_row["lat"]],
                    lon=[_row["lon"]],
                    mode="text",
                    text=[f"<b>{_abr}</b><br><b>{_val} sc</b> ({int(_row['n'])})"],
                    textfont=dict(
                        size=14 if _dest else 12,
                        color="#1A1A1A",
                        family="Helvetica Neue, sans-serif",
                    ),
                    hoverinfo="skip",
                    showlegend=False,
                ))

            _ocultas = len(df_reg) - len(_df_lbl)
            if _ocultas > 0:
                _fig.add_annotation(
                    x=0.01, y=0.01, xref="paper", yref="paper", xanchor="left",
                    text=f"<i>{_ocultas} região(ões) sem rótulo, por proximidade ou por terem "
                         f"menos de {_N_MIN_REGIAO} parcelas — passe o mouse ou veja o "
                         f"gráfico abaixo</i>",
                    showarrow=False, font=dict(size=11, color="#6B7280"))

        # legenda
        _df_leg_r = df_reg.sort_values(col_val, ascending=False)
        for _, _lr in _df_leg_r.iterrows():
            _c = _cor_map.get(_lr["regiao"], "#94A3B8")
            _abr = str(_lr["regiao"]).replace("MACRO ", "M ").replace("REC ", "")
            _lbl = f"{_abr}  {_fmt(_lr[col_val])} sc"

            _fig.add_trace(go.Scattergeo(
                lat=[None],
                lon=[None],
                mode="markers",
                marker=dict(color=_c, size=10, symbol="square"),
                name=_lbl,
                showlegend=True,
            ))

        for _r in sorted(set(_regioes_all) - set(df_reg["regiao"].dropna())):
            _c = _cor_map.get(_r, "#E5E7EB")
            _abr = str(_r).replace("MACRO ", "M ").replace("REC ", "")
            _fig.add_trace(go.Scattergeo(
                lat=[None],
                lon=[None],
                mode="markers",
                marker=dict(color=_c, size=10, symbol="square"),
                name=f"{_abr}  —",
                showlegend=True,
            ))

        _fig.update_geos(fitbounds="geojson", visible=False, bgcolor="#FFFFFF")
        _fig.update_layout(
            height=680,
            margin=dict(t=0, b=0, l=0, r=130),
            paper_bgcolor="#FFFFFF",
            font=dict(family="Helvetica Neue, sans-serif"),
            legend=dict(
                x=1.01, y=1, xanchor="left", yanchor="top",
                bgcolor="rgba(255,255,255,0.9)",
                bordercolor="#E5E7EB", borderwidth=1,
                font=dict(size=10),
                itemsizing="constant",
                tracegroupgap=2,
            ),
        )
        st.plotly_chart(_fig, use_container_width=True)

    st.divider()
    st.markdown(
        f'<p style="font-size:12px;font-weight:600;color:#6B7280;text-transform:uppercase;'
        f'letter-spacing:0.07em;">Produtividade Média por {label_nivel}</p>',
        unsafe_allow_html=True
    )
    _lollipop(df_reg, _cor_map, _melhor_reg, label_nivel)
    if _reg_fragil:
        st.warning(f"Nenhuma região tem {_N_MIN_REGIAO} parcelas ou mais para "
                   f"**{cultivar_sel}** neste recorte. As médias oscilam muito com poucas "
                   f"parcelas — leia os números, mas não eleja uma região por eles.")
    _exportar(df_reg, label_nivel, key=f"mp_xlsx_{label_nivel.lower()}")
with tab_macro:
    _mapa_regiao("regiao_macro", "Macro", fator_cor=0.65)

with tab_micro:
    _mapa_regiao("regiao_micro", "Micro", fator_cor=0.65)

# ── Rodapé ────────────────────────────────────────────────────────────────────
# usa o rodape() compartilhado, como as demais páginas do painel de milho, em vez de
# repetir o HTML aqui — assim uma mudança de assinatura vale para todas de uma vez
st.divider()
rodape()
