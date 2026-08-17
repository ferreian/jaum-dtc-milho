"""
pipeline_milho_2025.py
Pipeline da safra 2025 de milho (JAUM DTC) — refatorado do notebook jaum_dtc_milho_2025.
Adaptado para Streamlit: usa st.secrets (banco 2025) e st.cache_data.

Do banco cru ao modelo dimensional (base_plots + analíticas Faixa/Densidade + base_detalhe).
Preserva as réguas validadas da safra 2025:
  - peso em KG (sem ÷1000), PMG correto (sem ÷10), metragem 10m
  - estande final = média dos 8 pontos; perda usa estande final como denominador
  - população METROS_CONTAGEM=10; produtividade corrigida a 13,5% umidade
"""

import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from supabase import create_client

warnings.filterwarnings("ignore")

BASE_DIR = Path(__file__).parent
CONFIG_DIR = BASE_DIR / "config"

# ── Constantes de cálculo (réguas 2025) ───────────────────────────────────────
UMID_PADRAO = 13.5          # base de correção de umidade do grão
UMID_MIN, UMID_MAX = 10, 40 # faixa sã de umidade de colheita
PROD_TETO = 16000           # kg/ha acima disso é impossível (flag)
METROS_CONTAGEM = 10        # metragem da contagem de estande (2025 = 10)
TETO_FILEIRAS = 30
TETO_UMID_AMOSTRA = 100

# ── Detecção de teste — o DE-PARA USERS é a regra (fonte única) ────────────────
# Um registro é de teste se, e só se, rastreia até um usuário 'nao_fica' do de-para:
#   - fazenda de teste = dtcResponsavelRef é usuário 'nao_fica' → cai em cascata
#     (e, via join inner do enriquecimento, as avaliações dela caem junto);
#   - a Fazenda Teste nominal tem como dono o STINE Admin ('nao_fica'), então a
#     cascata já a remove — não precisa de lista de UUID chumbada;
#   - o material "TESTE MILHO" não está no de-para de materiais (config), então cai
#     no join com dePara=NaN, e seus plots vivem só em fazendas de teste (removidas).
# Não há UUID de fazenda/material chumbado: o de-para users é a única régua de teste.


# ── Conexão Supabase (projeto 2025) ───────────────────────────────────────────
@st.cache_resource
def get_supabase_2025():
    return create_client(
        st.secrets["supabase_2025"]["url"],
        st.secrets["supabase_2025"]["key"],
    )


def _extrair(supabase, nome: str) -> pd.DataFrame:
    response = supabase.table(nome).select("*").execute()
    return pd.DataFrame(response.data)


# Faixa plausível de timestamp Unix em SEGUNDOS: ~1970 a ~2100.
# Fora disso é lixo (ex.: data em milissegundos = ~1e12, que estoura o to_datetime).
_TS_MIN_SEG = 0
_TS_MAX_SEG = 4_102_444_800   # 2100-01-01


def _ts_para_datetime(serie: pd.Series) -> pd.Series:
    """Converte timestamp Unix (segundos) em datetime, à prova de overflow.
    O to_datetime(unit='s') multiplica internamente por 1e9 (→ nanossegundos); um valor em
    milissegundos ou lixo gigante estoura o int64 ANTES do errors='coerce' agir. Então:
      1) força numérico (strings/vazios → NaN);
      2) descarta o que está fora da faixa plausível de segundos (→ NaN, sem estourar);
      3) só então converte. Valores implausíveis viram NaT com segurança."""
    s = pd.to_numeric(serie, errors="coerce")
    s = s.where((s >= _TS_MIN_SEG) & (s <= _TS_MAX_SEG))
    return pd.to_datetime(s, unit="s", errors="coerce")


TABELAS = [
    "av1TratamentoMilho", "av2TratamentoMilho",
    "av3TratamentoMilho", "av4TratamentoMilho",
    "av1DetalheTratamentoMilho", "av2DetalheTratamentoMilho",
    "av3DetalheTratamentoMilho", "av4DetalheTratamentoMilho",
    "avaliacao", "fazenda", "cidade", "estado", "pais",
    "tratamentoBase", "users",
]


# ── Bases de referência (config/) ─────────────────────────────────────────────
@st.cache_data
def carregar_base_regioes():
    """Municípios com regiões macro/micro de milho + lat/long (de-para geográfico)."""
    return pd.read_excel(CONFIG_DIR / "base_municipios_regioes_soja_milho.xlsx")


@st.cache_data
def carregar_depara_users():
    """De-para de users (fica/nao_fica) — classifica quem é real e quem é teste."""
    path = CONFIG_DIR / "depara_users.csv"
    if path.exists():
        return pd.read_csv(path)
    return None


def uuids_teste_do_depara(df_users_banco):
    """
    Retorna o set de uuids a REMOVER (status 'nao_fica' no de-para).
    Opção C: users no banco que NÃO estão no de-para são MANTIDOS (tratados como reais),
    mas sinalizados para classificação posterior.
    """
    dep = carregar_depara_users()
    if dep is None:
        return set(), []  # sem de-para: não remove nada

    dep = dep.copy()
    dep["uuid"] = dep["uuid"].astype(str).str.strip()
    dep["status"] = dep["status"].astype(str).str.strip().str.lower()

    uuids_remover = set(dep.loc[dep["status"] == "nao_fica", "uuid"])

    # users no banco não classificados no de-para (Opção C: manter + avisar)
    uuids_banco = set(df_users_banco["uuid"].astype(str))
    uuids_no_depara = set(dep["uuid"])
    nao_classificados = sorted(uuids_banco - uuids_no_depara)

    return uuids_remover, nao_classificados


@st.cache_data
def carregar_depara_materiais_2025():
    """De-para de materiais da safra 2025: nome (cru) -> dePara (canônico) -> status.
    É o primeiro elo (nome->canônico da safra). O depara_mestre reconcilia entre safras.

    Duas colunas OPCIONAIS sustentam materiais que compartilham o mesmo nome no banco:
      - `indexTratamento`: preenchido SÓ nas exceções. Vazio = casa por nome (regra geral,
        que é o caso de 42 dos 43 materiais de 2025).
      - `tratamento_semente`: 'padrao' na maioria; identifica o tratamento industrial
        ('victrato'), permitindo a análise pareada do mesmo genótipo com e sem tratamento.
    O `usecols` por lambda tolera um CSV que ainda não tenha as colunas — assim as duas
    safras podem migrar em momentos diferentes sem quebrar. A coluna `n_tratamentos` do
    CSV é informativa (conferência do de-para) e continua fora do pipeline.
    """
    path = CONFIG_DIR / "depara_materiais_2025.csv"
    _cols = ["nome", "dePara", "status_material", "indexTratamento", "tratamento_semente"]
    dep = pd.read_csv(path, usecols=lambda c: c in _cols)
    dep["nome"] = dep["nome"].astype(str).str.strip()
    for col in ["dePara", "status_material", "tratamento_semente"]:
        if col in dep.columns:
            dep[col] = dep[col].astype(str).str.strip().replace(
                {"": np.nan, "nan": np.nan, "None": np.nan})
    if "indexTratamento" not in dep.columns:
        dep["indexTratamento"] = np.nan
    # numérico: o índice vem do banco como número e o CSV pode trazer vazio/texto
    dep["indexTratamento"] = pd.to_numeric(dep["indexTratamento"], errors="coerce")
    return dep


@st.cache_data
def carregar_depara_mestre():
    """Reconciliação de materiais entre safras (dePara_safra -> dePara_mestre)."""
    path = CONFIG_DIR / "depara_mestre.csv"
    if path.exists():
        return pd.read_csv(path)
    return None


# ══════════════════════════════════════════════════════════════════════════════
# BLOCO 2 — SILVER
# Transforma o cru: remove colunas de controle, ajusta tipos, corrige erros de
# origem, embute remoção de teste. Mantém nomes da fonte (rename canônico é gold).
# ══════════════════════════════════════════════════════════════════════════════

def _silver_apoio(dfs: dict) -> dict:
    """Silver das tabelas de apoio (pais, estado, cidade, fazenda, avaliacao,
    tratamentoBase, users). Remoção de teste embutida, com o DE-PARA DE USERS como
    régua única: os users 'nao_fica' derrubam (a) os próprios users e (b) EM CASCATA
    as fazendas cujo dtcResponsavelRef é um user de teste — e, via join inner do
    enriquecimento, as avaliações dessas fazendas. Não há UUID de fazenda/material
    chumbado: um registro só é teste se rastreia até um usuário 'nao_fica'."""

    # uuids de teste (de-para users): calculados no CRU, antes de filtrar a fazenda.
    uuids_user_teste, nao_classificados = uuids_teste_do_depara(dfs["users"])

    # ── pais ──
    df_pais = dfs["pais"].drop(columns=["firebase", "dataSync", "acao"], errors="ignore")

    # ── estado (corrige erro de origem: codigoEstado/nomeEstado trocados no banco) ──
    df_estado = dfs["estado"].drop(columns=["firebase", "dataSync", "acao"], errors="ignore")
    df_estado = df_estado.rename(columns={
        "codigoEstado": "nomeEstado",   # guardava o nome completo
        "nomeEstado":   "siglaEstado",  # guardava a sigla
    })

    # ── cidade ──
    df_cidade = dfs["cidade"].drop(columns=["firebase", "dataSync", "acao"], errors="ignore")

    # ── fazenda (remoção de teste em cascata via responsável nao_fica) ──
    df_fazenda = dfs["fazenda"].drop(
        columns=["firebase", "dataSync", "acao", "rcResponsavel", "hide",
                 "dataPlantio", "dataColheita"], errors="ignore")
    for col in ["criadoEm", "modificadoEm", "dataPlantioMilho", "dataColheitaMilho"]:
        df_fazenda[col] = _ts_para_datetime(df_fazenda[col])
    for col in ["dataPlantioMilho", "dataColheitaMilho"]:
        df_fazenda[col] = df_fazenda[col].where(df_fazenda[col].dt.year > 1970)
        df_fazenda[col] = df_fazenda[col].dt.date
    for col in ["isMilho", "isSoja"]:
        df_fazenda[col] = df_fazenda[col].astype(bool)

    # Remoção de teste em CASCATA — o de-para users é a régua única:
    # remove as fazendas cujo responsável (dtcResponsavelRef) é usuário 'nao_fica'.
    # Via join inner do enriquecimento, as avaliações dessas fazendas caem junto.
    resp_teste = df_fazenda["dtcResponsavelRef"].astype(str).isin(uuids_user_teste)
    if int(resp_teste.sum()):
        print(f"[de-para users] {int(resp_teste.sum())} fazenda(s) removida(s) "
              f"por responsável 'nao_fica'.")
    df_fazenda = df_fazenda[~resp_teste].reset_index(drop=True)

    # ── avaliacao ──
    df_avaliacao = dfs["avaliacao"].drop(
        columns=["firebase", "dataSync", "acao", "rcResponsavel"], errors="ignore")
    df_avaliacao["modificadoEm"] = _ts_para_datetime(df_avaliacao["modificadoEm"])
    df_avaliacao["dataAgendamento"] = _ts_para_datetime(df_avaliacao["dataAgendamento"])
    df_avaliacao["dataAgendamento"] = df_avaliacao["dataAgendamento"].where(
        df_avaliacao["dataAgendamento"].dt.year > 1970).dt.date

    # ── tratamentoBase (filtra milho) ──
    # Sem remoção de material chumbado: "TESTE MILHO" não está no de-para de materiais
    # (dePara=NaN no join do Bloco 3) e seus plots vivem só em fazendas de teste (removidas
    # pela cascata do de-para users). O de-para é a régua.
    df_tb = dfs["tratamentoBase"].drop(columns=["firebase", "dataSync", "acao", "gm"], errors="ignore")
    df_tb = df_tb[df_tb["cultura"] == "milho"].drop(columns=["cultura"]).reset_index(drop=True)

    # ── users (remoção de teste via DE-PARA; Opção C: não classificado = mantém + avisa) ──
    df_users = dfs["users"].drop(
        columns=["firebase", "dataSync", "acao", "forceUpdate", "time", "photoUrl", "phoneNumber"],
        errors="ignore")
    for col in ["createdTime", "lastUpdate"]:
        df_users[col] = _ts_para_datetime(df_users[col])
    for col in ["isGerente", "isAdmin", "ativo"]:
        if col in df_users.columns:
            df_users[col] = df_users[col].astype(bool)

    df_users = df_users[~df_users["uuid"].astype(str).isin(uuids_user_teste)].reset_index(drop=True)
    if nao_classificados:
        print(f"[de-para users] {len(nao_classificados)} user(s) no banco não classificados "
              f"(mantidos como reais): {nao_classificados[:5]}{'...' if len(nao_classificados) > 5 else ''}")

    return {
        "pais": df_pais, "estado": df_estado, "cidade": df_cidade,
        "fazenda": df_fazenda, "avaliacao": df_avaliacao,
        "tratamentoBase": df_tb, "users": df_users,
    }


def _silver_avaliacoes(dfs: dict) -> dict:
    """Silver das avaliações av1-av4: dropa controle (dataSync, acao, cultivar),
    converte datas da av3. NADA de cálculo/rename (é gold). Colunas mortas ficam cruas."""
    df_av1 = dfs["av1TratamentoMilho"].drop(columns=["dataSync", "acao", "cultivar"], errors="ignore")
    df_av2 = dfs["av2TratamentoMilho"].drop(columns=["dataSync", "acao", "cultivar"], errors="ignore")

    df_av3 = dfs["av3TratamentoMilho"].drop(columns=["dataSync", "acao", "cultivar"], errors="ignore")
    for col in ["dataFlorescimentoFeminina", "dataFlorescimentoMasculina"]:
        df_av3[col] = _ts_para_datetime(df_av3[col])
        df_av3[col] = df_av3[col].where(df_av3[col].dt.year > 1970).dt.date

    df_av4 = dfs["av4TratamentoMilho"].drop(columns=["dataSync", "acao", "cultivar"], errors="ignore")

    return {"av1": df_av1, "av2": df_av2, "av3": df_av3, "av4": df_av4}


def _silver_detalhe(dfs: dict) -> pd.DataFrame:
    """Empilha as 4 tabelas de detalhe numa só, coluna `avaliacao` marca a origem.
    Dropa fotoBase64 (morto em 2025), converte dataCriacao, normaliza vazios."""
    def _limpa_texto(s):
        return (s.astype(str).str.strip()
                 .replace({"": np.nan, "None": np.nan, "null": np.nan, "nan": np.nan}))

    tabelas_detalhe = ["av1DetalheTratamentoMilho", "av2DetalheTratamentoMilho",
                       "av3DetalheTratamentoMilho", "av4DetalheTratamentoMilho"]
    partes = []
    for tabela in tabelas_detalhe:
        df = dfs[tabela].drop(columns=["dataSync", "acao", "fotoBase64"], errors="ignore")
        df["dataCriacao"] = _ts_para_datetime(df["dataCriacao"])
        df["nota"]     = _limpa_texto(df["nota"])
        df["photoUrl"] = _limpa_texto(df["photoUrl"])
        df["avaliacao"] = tabela.replace("DetalheTratamentoMilho", "")
        partes.append(df[["uuid", "avaliacao", "tratamentoRef", "fazendaRef",
                          "dataCriacao", "nota", "photoUrl"]])
    return pd.concat(partes, ignore_index=True)


def _enriquecer_detalhe_fotos(df_det, avs_enriquecidas, df_fazenda=None):
    """Dá contexto às fotos e comentários: local, cultivar, status e responsável.

    ATENÇÃO À CHAVE. O detalhe traz `tratamentoRef`, que aponta para a linha da avaliação
    (av1TratamentoMilho.uuid), NÃO para o tratamentoBase — este é alcançado por `idBaseRef`.
    Ligar o detalhe direto no tratamentoBase por tratamentoRef não casa nada: todas as fotos
    saem sem material e os filtros de híbrido e status ficam vazios.

    O caminho certo é o detalhe cair nas tabelas de avaliação JÁ ENRIQUECIDAS, que trazem o
    contexto completo de uma vez — e com isso as fotos usam exatamente o mesmo cultivar e o
    mesmo status que o resto do painel, sem risco de divergir.

    avs_enriquecidas: dict {"av1": tb_av1, ...} — as tabelas depois do enriquecer_av.
    df_fazenda: opcional; usado só como reserva para os registros cujo tratamentoRef não casa,
                para a foto ao menos manter local e safra.
    """
    _COLS_CTX = ["cod_fazenda", "nomeFazenda", "cidade_nome", "estado_sigla",
                 "regiao_macro", "regiao_micro", "safra", "epoca", "nomeResponsavel",
                 "nome", "dePara", "status_material", "tratamento_semente",
                 "tipoTeste", "indexTratamento", "pop_tratamento"]

    _lk = []
    for _av, _tb in (avs_enriquecidas or {}).items():
        if _tb is None or _tb.empty or "uuid" not in _tb.columns:
            continue
        _cols = ["uuid"] + [c for c in _COLS_CTX if c in _tb.columns]
        _lk.append(_tb[_cols].copy())
    lookup = (pd.concat(_lk, ignore_index=True).drop_duplicates(subset="uuid")
              if _lk else pd.DataFrame(columns=["uuid"]))

    df = df_det.copy()
    df = df.merge(lookup.rename(columns={"uuid": "tratamentoRef"}),
                  on="tratamentoRef", how="left")

    # reserva: registro sem tratamento casado ainda tem fazendaRef, então preserva o local
    if df_fazenda is not None and "fazendaRef" in df.columns:
        _cols_faz = [c for c in ["uuid", "cod_fazenda", "nomeFazenda", "cidade_nome",
                                 "estado_sigla", "regiao_macro", "regiao_micro",
                                 "safra", "epoca"] if c in df_fazenda.columns]
        _faz = df_fazenda[_cols_faz].rename(
            columns={"uuid": "fazendaRef",
                     **{c: f"{c}__faz" for c in _cols_faz if c != "uuid"}})
        df = df.merge(_faz, on="fazendaRef", how="left")
        for _c in [c for c in _cols_faz if c != "uuid"]:
            if _c in df.columns and f"{_c}__faz" in df.columns:
                df[_c] = df[_c].fillna(df[f"{_c}__faz"])
        df = df.drop(columns=[c for c in df.columns if c.endswith("__faz")])

    # normaliza para av1..av4, seja qual for o formato que vier do silver
    _av_col = df["avaliacao"].astype(str).str.extract(r"(av\d)", expand=False)
    df["avaliacao"] = _av_col.fillna(df["avaliacao"].astype(str))
    return df


# ══════════════════════════════════════════════════════════════════════════════
# BLOCO 3 — ENRIQUECIMENTO
# Regiões (macro/micro milho), localidade + época/safra, cod_fazenda (com
# desambiguação), join do de-para de materiais, enriquecer_av (4 joins).
# ══════════════════════════════════════════════════════════════════════════════

def _prep_regioes():
    """Base de regiões, mantendo colunas de milho (microMilho/macroMilho) + ibge."""
    df = carregar_base_regioes()
    return df.drop(columns=["macroSoja", "microSoja", "latitude", "longitude"], errors="ignore")


def gerar_cod_cidade(nome_cidade, sigla_estado, n_letras=3):
    """Iniciais de cada palavra da cidade (sem acento, upper) + UF. Estende letras
    para desambiguar colisões (Uberaba/Uberlândia -> UBE)."""
    import unicodedata
    if pd.isna(nome_cidade) or pd.isna(sigla_estado):
        return np.nan
    nome_norm = unicodedata.normalize("NFKD", str(nome_cidade)).encode("ascii", "ignore").decode("ascii").upper()
    stopwords = {"DE", "DO", "DA", "DOS", "DAS", "E", "O", "A"}
    palavras = [p for p in nome_norm.replace("-", " ").split() if p not in stopwords]
    iniciais = "".join([p[0] for p in palavras])
    if len(iniciais) < n_letras and palavras:
        iniciais = (iniciais + palavras[0][1:])[:n_letras]
    return f"{iniciais}-{sigla_estado}"


def _enriquecer_fazenda(df_fazenda, df_cidade, df_estado, safra="25/26"):
    """Filtra milho, padroniza época/safra, junta cidade/estado/região, gera cod_fazenda.
    `safra`: '25/26' (default) ou '24/25' para o pipeline 2024. Época é sempre Safrinha."""
    df = df_fazenda[df_fazenda["isMilho"]].copy()
    df["epoca"] = "Safrinha"   # todo o milho safrinha (plantios fev-jun)
    df["safra"] = safra

    # lookup cidade + estado
    cidade_lookup = df_cidade.merge(
        df_estado[["uuid", "siglaEstado", "nomeEstado"]].rename(columns={"uuid": "estadoRef"}),
        on="estadoRef", how="left"
    )[["uuid", "nomeCidade", "siglaEstado", "nomeEstado"]].rename(columns={"uuid": "cidadeRef"})
    cidade_lookup["cidade_siglaEstado"] = cidade_lookup["nomeCidade"] + "_" + cidade_lookup["siglaEstado"]

    df = df.merge(cidade_lookup, on="cidadeRef", how="left")

    # join regiões de milho + ibge
    df_regioes = _prep_regioes()
    df = df.merge(
        df_regioes[["cidade_siglaEstado", "ibge", "microMilho", "macroMilho"]],
        on="cidade_siglaEstado", how="left"
    ).drop(columns=["cidade_siglaEstado"])

    df = df.rename(columns={
        "nomeCidade": "cidade_nome", "siglaEstado": "estado_sigla", "nomeEstado": "estado_nome",
        "macroMilho": "regiao_macro", "microMilho": "regiao_micro",
    })

    # cod_fazenda com desambiguação de colisão
    df["_base"] = df.apply(lambda r: gerar_cod_cidade(r["cidade_nome"], r["estado_sigla"]), axis=1)
    colidem = df.groupby("_base")["cidade_nome"].transform("nunique") > 1
    for n in range(4, 8):
        if not colidem.any():
            break
        df.loc[colidem, "_base"] = df.loc[colidem].apply(
            lambda r: gerar_cod_cidade(r["cidade_nome"], r["estado_sigla"], n), axis=1)
        colidem = df.groupby("_base")["cidade_nome"].transform("nunique") > 1

    df["_ano_safra"] = df["safra"].apply(lambda s: str(s).split("/")[0][-2:] if pd.notna(s) else np.nan)
    df["_chave"] = df["_base"] + "_" + df["_ano_safra"]
    contagem = df["_chave"].map(df["_chave"].value_counts())
    df["_sufixo"] = df.groupby("_chave").cumcount() + 1
    df["cod_fazenda"] = df.apply(
        lambda r: r["_chave"] if contagem[r.name] == 1 else f"{r['_chave']}_{int(r['_sufixo'])}", axis=1)
    df = df.drop(columns=["_base", "_ano_safra", "_chave", "_sufixo"])
    return df


def _enriquecer_tratamento(df_tb, dep_materiais=None):
    """Join do tratamentoBase com o de-para de materiais da safra (nome -> dePara -> status).
    `dep_materiais`: DataFrame do de-para; se None, usa o de 2025 (retrocompat).

    DUAS CAMADAS DE CASAMENTO:
      1. REGRA GERAL — por `nome`. Vale para as linhas do de-para com `indexTratamento` vazio,
         que é a maioria (42 dos 43 materiais de 2025).
      2. EXCEÇÕES — linhas com `indexTratamento` preenchido casam por nome + índice e
         SOBRESCREVEM a camada 1.

    A camada 2 existe porque o 9505PRO4 VICTRATO tem o MESMO nome do 9505PRO4 padrão no banco:
    o que os separa é `indexTratamento = 26`. Sem ela os dois viram um único material e, como o
    Victrato rende menos (10,4 sc/ha em 32 locais pareados, p = 0,0003, com população real ~4.500
    plantas/ha menor), a média do padrão sai contaminada — em MT, 142,6 em vez de 149,1.

    POR QUE NÃO trocar a chave para nome + índice: doze materiais têm dois números de tratamento
    apenas porque a numeração varia por ensaio (AS1868PRO4 10/11, SYD8124ZL 16/17, DKB360PRO3
    13/14, e outros nove) — nunca no mesmo local, e NÃO são materiais distintos. Com chave
    composta cada um precisaria de duas linhas no CSV e qualquer índice novo deixaria o material
    sem de-para. Aqui, índice desconhecido continua casando pela regra geral.

    O casamento acontece no tratamentoBase (que já traz `indexTratamento` da fonte), então tudo
    a jusante — av1..av4, base_plots, analíticas, detalhe, fotos — herda pelo `idBaseRef`.
    """
    dep = dep_materiais if dep_materiais is not None else carregar_depara_materiais_2025()
    df = df_tb.copy()
    if "pop" in df.columns and "pop_tratamento" not in df.columns:
        df = df.rename(columns={"pop": "pop_tratamento"})   # população-alvo do tratamento (Densidade)
    df["nome"] = df["nome"].astype(str).str.strip()

    if "indexTratamento" not in dep.columns:      # de-para antigo: só a regra geral existe
        return df.merge(dep, on="nome", how="left")

    _cols_sobrescreve = [c for c in ["dePara", "status_material", "tratamento_semente"]
                         if c in dep.columns]
    _geral = dep[dep["indexTratamento"].isna()].drop(columns=["indexTratamento"])
    _exc = (dep[dep["indexTratamento"].notna()]
            .rename(columns={"indexTratamento": "_idx",
                             **{c: f"_{c}_exc" for c in _cols_sobrescreve}}))

    # CUIDADO: o merge tem de usar SÓ `_geral`. Passar o de-para inteiro por `nome` duplicaria
    # a linha de todo material com exceção (o nome aparece duas vezes no CSV) — fan-out silencioso
    # que multiplicaria parcelas rede afora. Sem índice no tratamentoBase não há como desempatar,
    # então a exceção não é aplicada e o aviso sai no log em vez de virar dado errado.
    if _exc.empty or "indexTratamento" not in df.columns:
        if not _exc.empty:
            print(f"[de-para materiais] {len(_exc)} exceção(ões) por indexTratamento não "
                  f"aplicada(s): o tratamentoBase não trouxe a coluna `indexTratamento`.")
        return df.merge(_geral, on="nome", how="left")

    out = df.merge(_geral, on="nome", how="left")
    out["_idx"] = pd.to_numeric(out["indexTratamento"], errors="coerce")
    out = out.merge(_exc[["nome", "_idx"] + [f"_{c}_exc" for c in _cols_sobrescreve]],
                    on=["nome", "_idx"], how="left")

    _casou = out["_dePara_exc"].notna()
    for c in _cols_sobrescreve:
        out.loc[_casou, c] = out.loc[_casou, f"_{c}_exc"]
    return out.drop(columns=["_idx"] + [f"_{c}_exc" for c in _cols_sobrescreve])


def _fazer_enriquecer_av(df_fazenda, df_users, df_avaliacao, df_tb):
    """Retorna a função enriquecer_av com os lookups fechados (contexto via 4 joins)."""
    user_lookup = df_users[["uuid", "displayName"]].rename(
        columns={"uuid": "dtcResponsavelRef", "displayName": "nomeResponsavel"})
    av_lookup = df_avaliacao[["uuid", "fazendaRef"]].rename(columns={"uuid": "avaliacaoRef"})

    def enriquecer_av(df_av):
        df = df_av.copy()
        df = df.merge(av_lookup, on="avaliacaoRef", how="left")
        df = df.merge(
            df_fazenda[["uuid", "cod_fazenda", "nomeFazenda", "cidade_nome", "estado_sigla",
                        "regiao_macro", "regiao_micro", "safra", "epoca", "dtcResponsavelRef",
                        "dataPlantioMilho", "dataColheitaMilho",
                        "latitude", "longitude", "altitude"]].rename(columns={"uuid": "fazendaRef"}),
            on="fazendaRef", how="inner"
        )
        df = df.merge(user_lookup, on="dtcResponsavelRef", how="left")
        # inclui a população-alvo do tratamento (pop -> pop_tratamento): é o que define o ensaio
        # de Densidade (planejado). Diferente de populacao_real_plantas_ha (contada no campo, av4).
        _cols_tb = ["uuid", "dePara", "status_material", "regional"]
        # tratamento industrial de semente ('padrao'/'victrato'): vem do de-para da safra e
        # precisa viajar junto, senão a análise pareada não existe fora do tratamentoBase
        if "tratamento_semente" in df_tb.columns:
            _cols_tb.append("tratamento_semente")
        _ren_tb = {"uuid": "idBaseRef"}
        if "pop_tratamento" in df_tb.columns:
            _cols_tb.append("pop_tratamento")
        elif "pop" in df_tb.columns:
            _cols_tb.append("pop")
            _ren_tb["pop"] = "pop_tratamento"
        df = df.merge(
            df_tb[_cols_tb].rename(columns=_ren_tb),
            on="idBaseRef", how="left"
        )
        cols_ids = ["uuid", "avaliacaoRef", "idBaseRef", "fazendaRef", "dtcResponsavelRef"]
        cols_contexto = [
            "cod_fazenda", "nomeFazenda", "cidade_nome", "estado_sigla",
            "regiao_macro", "regiao_micro", "regional", "safra", "epoca",
            "dataPlantioMilho", "dataColheitaMilho", "latitude", "longitude", "altitude",
            "nomeResponsavel", "nome", "dePara", "status_material", "tratamento_semente",
            "tipoTeste", "indexTratamento", "pop_tratamento",
        ]
        cols_contexto = [c for c in cols_contexto if c in df.columns]
        cols_metricas = [c for c in df.columns if c not in cols_ids + cols_contexto]
        return df[cols_ids + cols_contexto + cols_metricas]

    return enriquecer_av


# ══════════════════════════════════════════════════════════════════════════════
# BLOCO 4 — GOLD av1 + av2
# Renomeia para canônico e calcula as métricas por plot. NADA de I/O nem prints:
# cada função recebe o tabelão enriquecido (saída de enriquecer_av) e devolve o
# gold. As réguas são idênticas ao notebook jaum_dtc_milho_2025 (validadas).
# ══════════════════════════════════════════════════════════════════════════════

# ── av1: qualidade inicial (8 notas + nota geral do técnico) ──────────────────
RENAME_AV1 = {
    "uniformidadeEmergencia":    "nota_uniformidade",
    "densidadePlantas":          "nota_densidade",
    "vigorPlantas":              "nota_vigor",
    "presencaPlantasDaninhas":   "nota_daninhas",
    "presencaPragas":            "nota_pragas",
    "presencaDoecas":            "nota_doencas",         # typo da fonte, preservado no de-para
    "homogenidadeCrescimento":   "nota_homogeneidade",   # typo da fonte, preservado no de-para
    "estadoGeralSolo":           "nota_solo",
    "nota0QualidadeInicialPlot": "qualidade_plot_inicial",  # NOVO 2025 (nota geral do técnico)
}

NOTAS_AV1 = ["nota_uniformidade", "nota_densidade", "nota_vigor", "nota_daninhas",
             "nota_pragas", "nota_doencas", "nota_homogeneidade", "nota_solo"]


def _gold_av1(tb_av1: pd.DataFrame) -> pd.DataFrame:
    """Gold av1: rename canônico + media_categorias (média das 8 notas, escala 1-5).

    Régua (2025):
      - 0 → NaN (não avaliado) antes da média;
      - plot com QUALQUER nota >=6 (resíduo de escala 1-9) → media_categorias NaN
        no plot inteiro. Em 2025 não há nenhum (proteção herdada de 2024);
      - qualidade_plot_inicial é a nota geral do técnico (nova em 2025), mantida crua.
    """
    df = tb_av1.copy()
    df = df.rename(columns=RENAME_AV1)

    notas = df[NOTAS_AV1].apply(pd.to_numeric, errors="coerce")
    escala_errada = (notas >= 6).any(axis=1)
    notas_para_media = notas.where(notas >= 1)   # 0 → NaN (não avaliado)

    df["media_categorias"] = notas_para_media.mean(axis=1).round(1)
    df.loc[escala_errada, "media_categorias"] = pd.NA
    return df


# ── av2: sanidade (nota + incidência + classe de reação por doença) ───────────
RENAME_AV2 = {
    "manchaTurcicum":   "nota_turcicum",
    "manchaCercospora": "nota_cercospora",
    "manchaBranca":     "nota_mancha_branca",
    "manchaBipolaris":  "nota_bipolaris",
    "ferrugemTropical": "nota_ferrugem_tropical",
    "enfezamento":      "nota_enfezamento",
    "tombamentoVerde":  "nota_tombamento_verde",   # MORTO 2025 (virou contagem na av4)
    "graosArdidos":     "graos_ardidos_pct",        # ATIVO (ver nota abaixo)
    "empalhamento":     "nota_empalhamento",        # NOVO 2025 (qualidade da espiga, não doença)
}

DOENCAS_AV2 = ["nota_turcicum", "nota_cercospora", "nota_mancha_branca",
               "nota_bipolaris", "nota_ferrugem_tropical", "nota_enfezamento"]

_COLS_EXTRAS_AV2 = ["nota_empalhamento", "nota_tombamento_verde", "graos_ardidos_pct"]


def _classificar_doenca(nota):
    """Reação da planta na escala 1-9 (9 = mais resistente).
    1-2 AS (altamente suscetível), 3-4 S, 5-6 MT, 7-8 T, 9 R (resistente)."""
    if pd.isna(nota):
        return np.nan
    if nota <= 2:
        return "AS"
    if nota <= 4:
        return "S"
    if nota <= 6:
        return "MT"
    if nota <= 8:
        return "T"
    return "R"


def _gold_av2(tb_av2: pd.DataFrame) -> pd.DataFrame:
    """Gold av2: rename canônico + por doença nota/inc_/class_.

    Régua (2025):
      - 0 → NaN (não avaliado);
      - inc_ (incidência): 1 = presente (nota 1-5), 0 = ausente (6-9), NA = não avaliado;
      - class_ (reação): AS/S/MT/T/R;
      - empalhamento é campo à parte (não entra em sanidade);
      - tombamentoVerde MORTO em 2025 (virou a contagem de green snap na av4, sai vazio);
      - graosArdidos CONTINUA SENDO COLETADO: a auditoria de 25/26 tem 887 parcelas avaliadas
        e 36 com ocorrência (1,0% a 24,0%), em 30 dos 40 locais. O comentário anterior dizia
        que saía sempre vazio e estava errado — corrigido em 08/2026.
    Ordena: meta + (nota→inc→class por doença) + campos à parte.
    """
    df = tb_av2.copy()
    df = df.rename(columns=RENAME_AV2)

    for nota in DOENCAS_AV2:
        n = pd.to_numeric(df[nota], errors="coerce").where(lambda x: x > 0)  # 0 → NaN
        df[f"inc_{nota}"]   = n.between(1, 5).where(n.notna()).astype("Int64")
        df[f"class_{nota}"] = n.apply(_classificar_doenca)

    cols_meta = [c for c in df.columns
                 if c not in DOENCAS_AV2
                 and not c.startswith("inc_") and not c.startswith("class_")
                 and c not in _COLS_EXTRAS_AV2]

    cols_doencas = []
    for nota in DOENCAS_AV2:
        cols_doencas += [nota, f"inc_{nota}", f"class_{nota}"]

    cols_extras = [c for c in _COLS_EXTRAS_AV2 if c in df.columns]

    return df[cols_meta + cols_doencas + cols_extras]


# ══════════════════════════════════════════════════════════════════════════════
# BLOCO 5 — GOLD av3 + av4
# av3 (arquitetura de planta + florescimento): SAFRA-AGNÓSTICO — as réguas de altura
#   e florescimento são idênticas entre safras, então o 2024 importa estas funções.
# av4 (produtividade): a ESTRUTURA é comum (produtividade Caminho A → detalhe das
#   subamostras → consolidada + população + Caminho B), mas as réguas mudam por safra
#   (peso g×kg, metragem, nº de pontos de estande). O que muda vem por
#   parâmetro/constante; a leitura de colunas de estande é explícita por safra.
# ══════════════════════════════════════════════════════════════════════════════

# ── av3: alturas (5 plantas) + florescimento ──────────────────────────────────
COLS_ALT_PLANTA = [f"planta{i}AlturaPlanta" for i in range(1, 6)]
COLS_ALT_ESPIGA = [f"planta{i}AlturaEspiga" for i in range(1, 6)]
FLOR_MIN, FLOR_MAX = 40, 80   # dias plantio→florescimento sãos (safrinha ~45-75)


def padronizar_altura_cm(s: pd.Series) -> pd.Series:
    """Altura de planta/espiga em cm. Blinda tipo misto (int/float), 0 = não medido → NaN,
    outlier >350 → NaN, valor <10 (digitado em metros) → ×100. Régua igual nas duas safras."""
    s = pd.to_numeric(s, errors="coerce")
    s = s.where(s > 0)          # 0 = não medido
    s = s.where(s <= 350)       # outlier de digitação
    s = s.mask(s < 10, s * 100) # metros → cm
    return s


def _gold_av3(tb_av3: pd.DataFrame) -> pd.DataFrame:
    """Gold av3 consolidada: altura_planta_cm / altura_espiga_cm (média das 5 subamostras)
    e dias de florescimento (masc/fem) com validação de plausibilidade fisiológica.
    Remove as 10 subamostras cruas (migram para a detalhe). Safra-agnóstico."""
    df = tb_av3.copy()

    alt_planta = df[COLS_ALT_PLANTA].apply(padronizar_altura_cm)
    alt_espiga = df[COLS_ALT_ESPIGA].apply(padronizar_altura_cm)
    df["altura_planta_cm"] = alt_planta.mean(axis=1).round(1)
    df["altura_espiga_cm"] = alt_espiga.mean(axis=1).round(1)

    plantio = pd.to_datetime(df["dataPlantioMilho"], errors="coerce")
    flor_m  = pd.to_datetime(df["dataFlorescimentoMasculina"], errors="coerce")
    flor_f  = pd.to_datetime(df["dataFlorescimentoFeminina"], errors="coerce")
    df["dias_flor_masculino"] = (flor_m - plantio).dt.days
    df["dias_flor_feminino"]  = (flor_f - plantio).dt.days

    for col in ["dias_flor_masculino", "dias_flor_feminino"]:
        bruto = df[col]
        fora = bruto.notna() & ((bruto < FLOR_MIN) | (bruto > FLOR_MAX))
        df[f"{col}_valido"] = bruto.where(~fora)
        df[f"obs_{col}"] = fora.map({True: "data_implausivel", False: ""})

    return df.drop(columns=COLS_ALT_PLANTA + COLS_ALT_ESPIGA)


def _gold_av3_detalhe(tb_av3: pd.DataFrame) -> pd.DataFrame:
    """Detalhe av3 (long puro): uma linha por (plot × planta × métrica). Mesma régua da
    consolidada, então a média por (plot, métrica) reproduz a consolidada. Safra-agnóstico."""
    cols_contexto = [c for c in tb_av3.columns if c not in COLS_ALT_PLANTA + COLS_ALT_ESPIGA]

    blocos = []
    for i in range(1, 6):
        b = tb_av3[cols_contexto].copy()
        b["planta"] = i
        b["altura_planta_cm"] = padronizar_altura_cm(tb_av3[f"planta{i}AlturaPlanta"])
        b["altura_espiga_cm"] = padronizar_altura_cm(tb_av3[f"planta{i}AlturaEspiga"])
        blocos.append(b)
    wide = pd.concat(blocos, ignore_index=True)

    return wide.melt(
        id_vars=cols_contexto + ["planta"],
        value_vars=["altura_planta_cm", "altura_espiga_cm"],
        var_name="metrica", value_name="valor",
    ).sort_values(["uuid", "planta", "metrica"]).reset_index(drop=True)


# ── av4: produtividade (Caminho A) + estande/detalhe + consolidada (Caminho B) ─
# Réguas que MUDAM por safra vêm por função (não flag), porque são CONDICIONAIS:
# 2024 aplica rede de segurança ao peso (≥100→÷1000); o PMG NÃO tem régua (ver regua_pmg_2024),
# enquanto 2025 usa o valor direto. Defaults = 2025 (identidade). O 2024 passa suas réguas.
MAPA_PERDAS = {"acamadas": "NumPlantasAcamadas", "quebradas": "NumPlantasQuebradas",
               "dominadas": "NumPlantasDominadas", "colmo_podre": "ColmoPodre"}

# Fenômenos da colheita (av4) — NOVOS em 2025 (ajuste de protocolo; não existem em 2024).
# Contagem por subamostra (4 em 2025), agregação PLOT-LEVEL: MÉDIA das contagens ÷ estande
# final (8 pontos) × 100 — a MESMA régua das perdas, e no app é a mesma subamostra.
# Antes era a SOMA das 4, o que devolvia a taxa multiplicada por 4 (não era taxa).
# Vive só na consolidada, não no detalhe long (4 subamostras ≠ 8 pontos de estande).
# Em 2024 as colunas não existem → não entram (NaN no concat).
FENOMENOS_AV4 = {
    "green_snap":        "GreenSnap",
    "morte_prematura":   "MortePrematura",
    "ma_formacao_espigas": "MaformacaoEspigas",
    "enfezamento":       "Enfezamento",
}

COLS_ESTANDE_8P = [f"numeroPlantas10Metros{i}aFinal" for i in range(1, 9)]      # 2025: 8 pontos
COLS_ESTANDE_5SUB = [f"planta{i}NumPlantas10metros" for i in range(1, 6)]       # 2024: 5 subamostras

# Quantas subamostras o PROTOCOLO de cada safra coleta. Régua de safra, como peso e PMG.
# O banco tem slots planta1..planta5 nas duas safras, mas em 2025 o protocolo passou a 4 —
# a 5ª não é preenchida. Como a contagem 0 agora ENTRA na média, ler um slot fora do
# protocolo que esteja gravado como 0 diluiria toda a taxa; por isso o limite é explícito
# aqui e não depende de o banco estar vazio no lugar certo.
N_SUBAMOSTRAS_2025 = 4
N_SUBAMOSTRAS_2024 = 5
MAX_SLOTS_SUBAMOSTRA = 5   # slots existentes no banco (para excluir do contexto)


def regua_peso_2025(peso_kg: pd.Series) -> pd.Series:
    """2025: peso já vem em kg — usa direto (identidade)."""
    return peso_kg


def regua_peso_2024(peso: pd.Series) -> pd.Series:
    """2024: rede de segurança — peso ≥100 provavelmente veio em gramas → ÷1000.
    Valores <100 já estão em kg e ficam intactos. Condicional, não incondicional."""
    return peso.where(peso < 100, peso / 1000.0)


def regua_pmg_2025(pmg: pd.Series) -> pd.Series:
    """2025: PMG já vem correto (mediana ~360 g) — usa direto (identidade)."""
    return pmg


def regua_pmg_2024(pmg: pd.Series) -> pd.Series:
    """2024: PMG já vem correto em gramas (mediana ~352 g) — identidade, igual a 2025.

    CORRIGIDO (era `pmg.where(pmg <= 40, pmg / 10.0)`). A régua antiga supunha vírgula
    deslocada ("231 = 23,1") e dividia por 10 todo valor >40. A premissa estava errada:
    231 g é PMG de milho perfeitamente plausível — é grão de milho, não de soja. O efeito
    era um PMG dez vezes menor na safra 24/25 (mediana 35 g, faixa 19–51).

    Três evidências de que os valores crus já estão certos:
      1. ordem de grandeza — 2025, sem régua nenhuma, dá mediana 341 g; 2024 cru dá 352 g;
      2. os componentes do Caminho B (fileiras 16,5 e grãos/fileira 33,6) são idênticos entre
         as safras, então só o PMG estava fora de escala;
      3. o Caminho B fecha: com a régua antiga a estimativa saía em 11% da produtividade
         medida (razão 0,115); sem ela, 1,15 — mesma família da razão de 2025 (1,34).
    Todos os 1.159 valores de 2024, multiplicados por 10, caem na faixa plausível 150–600 g.

    A função fica (mesmo sendo identidade) porque o pipeline_milho_2024 a importa por nome e
    porque o par régua-por-safra é o contrato do núcleo compartilhado.
    """
    return pmg


def _gold_av4_produtividade(tb_av4: pd.DataFrame, *,
                            regua_peso=regua_peso_2025,
                            umid_padrao: float = UMID_PADRAO,
                            umid_min: float = UMID_MIN, umid_max: float = UMID_MAX,
                            prod_teto: float = PROD_TETO) -> pd.DataFrame:
    """Caminho A — produtividade medida pelo peso da parcela, corrigida a `umid_padrao`.

    `regua_peso`: função que normaliza o peso para kg. 2025=identidade; 2024=≥100→÷1000.
    Área com teto de espaçamento ≤2m. Flags: nao_colhido / sem_geometria /
    umidade_baixa/alta / prod_impossivel. Bloqueantes zeram a produtividade válida."""
    df = tb_av4.copy()

    # Blindagem contra overflow int64: força numérico (→ float64) toda coluna do cálculo.
    # Colunas vindas do banco podem chegar como int64 ou texto; float64 satura (inf) em vez
    # de estourar, e o pd.to_numeric neutraliza valores-texto.
    for c in ["pesoParcela", "humidade", "numeroLinhas", "comprimentoLinha", "espacamento"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    peso = df["pesoParcela"].where(df["pesoParcela"] > 0)
    peso_kg = regua_peso(peso)   # 2025: direto; 2024: ≥100 → ÷1000

    umidade = df["humidade"].where(df["humidade"] > 0)
    df["umidade_pct"] = umidade.round(1)

    n_linhas = df["numeroLinhas"].where(df["numeroLinhas"] > 0)
    compr    = df["comprimentoLinha"].where(df["comprimentoLinha"] > 0)
    espac    = df["espacamento"].where((df["espacamento"] > 0) & (df["espacamento"] <= 2))
    df["area_parcela_m2"] = (n_linhas * compr * espac).round(2)

    peso_corrigido = peso_kg * (100 - umidade) / (100 - umid_padrao)
    df["produtividade_kg_ha"] = (peso_corrigido * (10000 / df["area_parcela_m2"])).round(1)
    df["produtividade_sacas_ha"] = (df["produtividade_kg_ha"] / 60).round(1)

    flags = pd.Series([[] for _ in range(len(df))], index=df.index)
    sem_peso = peso_kg.isna()
    sem_area = df["area_parcela_m2"].isna()
    flags[sem_peso] = flags[sem_peso].apply(lambda x: x + ["nao_colhido"])
    flags[~sem_peso & sem_area] = flags[~sem_peso & sem_area].apply(lambda x: x + ["sem_geometria"])
    u = df["umidade_pct"]
    flags[u.notna() & (u < umid_min)] = flags[u.notna() & (u < umid_min)].apply(lambda x: x + ["umidade_baixa"])
    flags[u.notna() & (u > umid_max)] = flags[u.notna() & (u > umid_max)].apply(lambda x: x + ["umidade_alta"])
    pa = df["produtividade_kg_ha"] > prod_teto
    flags[pa.fillna(False)] = flags[pa.fillna(False)].apply(lambda x: x + ["prod_impossivel"])

    df["flags_produtividade"] = flags.apply(lambda x: "; ".join(x))
    bloqueantes = {"nao_colhido", "sem_geometria", "prod_impossivel"}
    bloqueante = flags.apply(lambda x: bool(bloqueantes & set(x)))
    df["produtividade_valida_kg_ha"]    = df["produtividade_kg_ha"].where(~bloqueante)
    df["produtividade_valida_sacas_ha"] = df["produtividade_sacas_ha"].where(~bloqueante)
    return df


def _gold_av4_detalhe(tb_av4: pd.DataFrame, *,
                      cols_estande: list = COLS_ESTANDE_8P,
                      n_subamostras: int = N_SUBAMOSTRAS_2025,
                      umid_padrao: float = UMID_PADRAO,
                      regua_pmg=regua_pmg_2025,
                      teto_fileiras: float = TETO_FILEIRAS,
                      teto_umid_amostra: float = TETO_UMID_AMOSTRA) -> tuple:
    """Detalhe av4 (long): subamostras × componentes + perdas, pct sobre o estande FINAL.

    Réguas que MUDAM por safra:
      - `cols_estande`: em 2025 são os 8 pontos (…{1..8}aFinal); em 2024, o estande 5-sub;
      - `n_subamostras`: quantas o protocolo coleta — 2025=4, 2024=5. Os slots além disso
        existem no banco mas não são preenchidos, e ler um deles gravado como 0 diluiria
        a média (a contagem 0 conta);
      - `regua_pmg`: função que normaliza o PMG. 2025=identidade; 2024=>40→÷10 (vírgula).
    Denominador da perda = estande final do plot (média dos pontos válidos), NUNCA a soma.
    Contagem 0 é dado (contou e não havia) e ENTRA na média; só a subamostra vazia fica fora.
    Retorna (tb_av4_detalhe_long, estande_final_series) — a série alimenta a consolidada."""
    est = tb_av4[cols_estande].apply(lambda s: pd.to_numeric(s, errors="coerce")).where(lambda x: x > 0)
    estande_final = est.mean(axis=1)

    # exclui do contexto TODOS os slots do banco, mesmo os fora do protocolo da safra
    cols_sub = [f"planta{i}{m}" for i in range(1, MAX_SLOTS_SUBAMOSTRA + 1)
                for m in ["NumFileiras", "NumGraosPorFileira", "PesoMilGraos",
                          "UmidadeAmostraMilGraos"] + list(MAPA_PERDAS.values())]
    cols_contexto = [c for c in tb_av4.columns if c not in cols_sub and c not in cols_estande]

    blocos = []
    for i in range(1, n_subamostras + 1):   # só as subamostras do protocolo da safra
        b = tb_av4[cols_contexto].copy()
        b["planta"] = i
        b["estande_final_plot"] = estande_final.round(1)  # mesmo nas 5 sub (auditável)

        fil = pd.to_numeric(tb_av4[f"planta{i}NumFileiras"], errors="coerce").where(lambda x: x > 0)
        b["fileiras"] = fil.where(fil <= teto_fileiras)
        b["graos_fileira"] = pd.to_numeric(tb_av4[f"planta{i}NumGraosPorFileira"], errors="coerce").where(lambda x: x > 0)

        pmg = pd.to_numeric(tb_av4[f"planta{i}PesoMilGraos"], errors="coerce").where(lambda x: x > 0)
        pmg = regua_pmg(pmg)   # 2025: direto; 2024: >40 → ÷10
        umid_pmg = pd.to_numeric(tb_av4[f"planta{i}UmidadeAmostraMilGraos"], errors="coerce").where(lambda x: x > 0)
        umid_pmg = umid_pmg.where(umid_pmg <= teto_umid_amostra)
        b["pmg_bruto_g"] = pmg
        b["umidade_pmg_pct"] = umid_pmg.round(1)
        b["pmg_corrigido_g"] = (pmg * (100 - umid_pmg) / (100 - umid_padrao)).round(1)

        # Perdas E fenômenos vêm da MESMA subamostra no app (mesma aba, mesmos 10 m), então
        # são calculados juntos, com a mesma régua. 0 é DADO, não ausência: o app registra 0
        # quando o avaliador percorreu e não encontrou. Só o vazio (não avaliada) sai.
        for nome, col in list(MAPA_PERDAS.items()) + list(FENOMENOS_AV4.items()):
            col_i = f"planta{i}{col}"
            if col_i not in tb_av4.columns:
                continue          # fenômeno em 2024: coluna não existe → não cria métrica
            cont = pd.to_numeric(tb_av4[col_i], errors="coerce")
            b[f"{nome}_n"] = cont
            b[f"pct_{nome}"] = ((cont / estande_final) * 100).round(1)  # denominador = estande final
        blocos.append(b)

    wide = pd.concat(blocos, ignore_index=True)
    metricas_valor = ["estande_final_plot", "fileiras", "graos_fileira",
                      "pmg_bruto_g", "umidade_pmg_pct", "pmg_corrigido_g",
                      "acamadas_n", "quebradas_n", "dominadas_n", "colmo_podre_n",
                      "pct_acamadas", "pct_quebradas", "pct_dominadas", "pct_colmo_podre"]
    # fenômenos (2025): entram só se as colunas existirem — em 2024 nem chegam a ser criadas
    metricas_valor += [m for f in FENOMENOS_AV4
                       for m in (f"{f}_n", f"pct_{f}") if m in wide.columns]
    detalhe = wide.melt(
        id_vars=cols_contexto + ["planta"], value_vars=metricas_valor,
        var_name="metrica", value_name="valor",
    ).sort_values(["uuid", "planta", "metrica"]).reset_index(drop=True)
    return detalhe, estande_final




def _gold_av4_fenomenos(df_consolidada: pd.DataFrame, tb_av4_detalhe: pd.DataFrame) -> pd.DataFrame:
    """Traz os fenômenos da colheita do DETALHE para a consolidada.

    Não recalcula nada: as contagens e os percentuais por subamostra já são produzidos em
    `_gold_av4_detalhe`, na mesma passagem das perdas (no app é a mesma subamostra). Aqui só
    se agrega por plot:
      - `pct_<fen>`   = MÉDIA dos percentuais das subamostras avaliadas;
      - `<fen>_plantas` = soma bruta das contagens, para auditoria.

    Fonte única de fórmula — antes esta função repetia o cálculo por conta própria, e foi assim
    que perda e fenômeno acabaram com réguas diferentes. Em 2024 os fenômenos não existem no
    detalhe e a função não faz nada (schema aditivo)."""
    if tb_av4_detalhe is None or tb_av4_detalhe.empty or "metrica" not in tb_av4_detalhe.columns:
        return df_consolidada
    metricas = set(tb_av4_detalhe["metrica"].unique())
    presentes = [c for c in FENOMENOS_AV4 if f"pct_{c}" in metricas]
    if not presentes:
        return df_consolidada

    df = df_consolidada.set_index("uuid")
    for canon in presentes:
        _n = tb_av4_detalhe[tb_av4_detalhe["metrica"] == f"{canon}_n"]
        _p = tb_av4_detalhe[tb_av4_detalhe["metrica"] == f"pct_{canon}"]
        df[f"{canon}_plantas"] = _n.groupby("uuid")["valor"].sum(min_count=1).reindex(df.index).round(0)
        df[f"pct_{canon}"] = _p.groupby("uuid")["valor"].mean().reindex(df.index).round(1)
    return df.reset_index()


def _gold_av4_consolidar(tb_av4_prod: pd.DataFrame, tb_av4_detalhe: pd.DataFrame, *,
                         metros_contagem: int = METROS_CONTAGEM) -> pd.DataFrame:
    """Consolida o av4: junta a produtividade (Caminho A) com as médias das subamostras
    (do detalhe), calcula população real e o Caminho B (validação) + divergência.

    Régua que MUDA por safra: `metros_contagem` (2025=10, 2024=5).
    - pct_perda_total = SOMA das 4 pct médias (réplicas → média por subamostra, depois soma);
    - populacao = plantas_10m_media / metros_contagem / espacamento × 10000;
    - Caminho B = pop × fileiras × graos × (pmg/1000) / 1000 (não oficial, valida o A)."""
    medias = tb_av4_detalhe.pivot_table(index="uuid", columns="metrica", values="valor", aggfunc="mean")
    # Blinda: se uma métrica for inteiramente NaN (ex.: perda ausente numa safra parcial),
    # o pivot_table a descarta. Reindexamos para garantir todas as colunas esperadas.
    esperadas = ["estande_final_plot", "fileiras", "graos_fileira", "pmg_corrigido_g",
                 "pct_acamadas", "pct_quebradas", "pct_dominadas", "pct_colmo_podre"]
    medias = medias.reindex(columns=medias.columns.union(esperadas))

    df = tb_av4_prod.set_index("uuid")
    df["plantas_10m_media"]   = medias["estande_final_plot"].round(1)
    df["fileiras_media"]      = medias["fileiras"].round(1)
    df["graos_fileira_media"] = medias["graos_fileira"].round(1)
    df["pmg_corrigido_g"]     = medias["pmg_corrigido_g"].round(1)
    for p in ["acamadas", "quebradas", "dominadas", "colmo_podre"]:
        df[f"pct_{p}"] = medias[f"pct_{p}"].round(1)

    cols_perda = ["pct_acamadas", "pct_quebradas", "pct_dominadas", "pct_colmo_podre"]
    # sum() trata NaN como 0; preservamos NaN só quando TODAS as 4 pct estão ausentes
    # (perda não medida ≠ perda 0%). Com ao menos uma medida, soma normal.
    df["pct_perda_total"] = df[cols_perda].sum(axis=1, min_count=1).round(1)

    espac = df["espacamento"].where((df["espacamento"] > 0) & (df["espacamento"] <= 2))
    df["populacao_real_plantas_ha"] = (
        (df["plantas_10m_media"] / metros_contagem / espac) * 10000).round(0)

    pop = df["populacao_real_plantas_ha"]
    prod_b = pop * df["fileiras_media"] * df["graos_fileira_media"] \
        * (df["pmg_corrigido_g"] / 1000) / 1000
    df["prod_estimada_kg_ha"]    = prod_b.round(1)
    df["prod_estimada_sacas_ha"] = (prod_b / 60).round(1)
    df["divergencia_prod_pct"] = (
        (df["prod_estimada_kg_ha"] - df["produtividade_valida_kg_ha"])
        / df["produtividade_valida_kg_ha"] * 100).round(1)

    return df.reset_index()


def _gold_av4(tb_av4: pd.DataFrame) -> tuple:
    """Orquestra o av4 da safra 2025 (réguas 2025 por default). Retorna (consolidada, detalhe).
    O 2024 terá o seu próprio wrapper chamando as mesmas funções com suas réguas."""
    prod = _gold_av4_produtividade(tb_av4, regua_peso=regua_peso_2025)       # 2025: kg direto
    detalhe, estande_final = _gold_av4_detalhe(tb_av4, cols_estande=COLS_ESTANDE_8P,
                                               regua_pmg=regua_pmg_2025)      # 2025: 8 pontos, PMG direto
    consolidada = _gold_av4_consolidar(prod, detalhe, metros_contagem=METROS_CONTAGEM)  # 2025: 10m
    consolidada = _gold_av4_fenomenos(consolidada, detalhe)                  # 2025: fenômenos (do detalhe)
    return consolidada, detalhe


# ══════════════════════════════════════════════════════════════════════════════
# BLOCO 6 — CONSOLIDAÇÃO + ORQUESTRAÇÃO
# Modelo dimensional: base_plots (1 linha/plot com contexto) + analíticas por
# tipoTeste (Faixa, Densidade, com as métricas das 4 av) + base_detalhe (long de
# av3+av4). rodar_pipeline() encadeia extração→silver→enriquecimento→gold→consolidação
# e devolve um dict de DataFrames, tudo em cache. Nada de I/O em disco.
# ══════════════════════════════════════════════════════════════════════════════

CHAVE = ["fazendaRef", "idBaseRef", "tipoTeste", "indexTratamento"]
CONTEXTO = ["cod_fazenda", "nomeFazenda", "nomeProdutor", "cidade_nome", "estado_sigla",
            "regiao_macro", "regiao_micro", "safra", "epoca", "nome", "dePara",
            "status_material", "tratamento_semente", "pop_tratamento", "nomeResponsavel",
            "dataPlantioMilho", "dataColheitaMilho", "latitude", "longitude"]

# Métricas por avaliação levadas às analíticas (nomes canônicos de 2025).
_DOENCAS = ["turcicum", "cercospora", "mancha_branca", "bipolaris", "ferrugem_tropical", "enfezamento"]
MET_AV1 = ["media_categorias", "qualidade_plot_inicial"]
MET_AV2 = ([f"nota_{d}" for d in _DOENCAS] + [f"class_nota_{d}" for d in _DOENCAS]
           + ["nota_empalhamento", "nota_tombamento_verde", "graos_ardidos_pct"])
MET_AV3 = ["altura_planta_cm", "altura_espiga_cm",
           "dias_flor_masculino_valido", "dias_flor_feminino_valido"]
MET_AV4 = ["produtividade_valida_kg_ha", "produtividade_valida_sacas_ha",
           "produtividade_kg_ha", "flags_produtividade", "umidade_pct",
           "populacao_real_plantas_ha", "plantas_10m_media",
           "fileiras_media", "graos_fileira_media", "pmg_corrigido_g",
           "prod_estimada_kg_ha", "prod_estimada_sacas_ha", "divergencia_prod_pct",
           "pct_acamadas", "pct_quebradas", "pct_dominadas", "pct_colmo_podre", "pct_perda_total",
           # fenômenos da colheita (novos 2025; ausentes em 2024 → NaN no empilhamento)
           "green_snap_plantas", "pct_green_snap",
           "morte_prematura_plantas", "pct_morte_prematura",
           "ma_formacao_espigas_plantas", "pct_ma_formacao_espigas",
           "enfezamento_plantas", "pct_enfezamento"]

COLS_DETALHE = ["uuid", "fazendaRef", "idBaseRef", "tipoTeste", "indexTratamento",
                "cod_fazenda", "nomeFazenda", "cidade_nome", "estado_sigla",
                "regiao_macro", "regiao_micro", "safra", "epoca",
                "nome", "dePara", "status_material", "tratamento_semente",
                "pop_tratamento", "nomeResponsavel",
                "planta", "metrica", "valor"]


def _montar_base_plots(golds: dict) -> pd.DataFrame:
    """Dimensão central: 1 linha por plot (CHAVE) com o contexto de qualquer av que o tenha."""
    blocos = []
    for df in golds.values():
        cols = [c for c in CHAVE + CONTEXTO if c in df.columns]
        blocos.append(df[cols])
    return (pd.concat(blocos, ignore_index=True)
            .groupby(CHAVE, as_index=False).first())


def _consolidar_tipo(tipo_teste: str, base_plots: pd.DataFrame, golds: dict) -> pd.DataFrame:
    """Analítica de um tipoTeste (Faixa/Densidade): base_plots + métricas das 4 av por plot."""
    tab = base_plots[base_plots["tipoTeste"] == tipo_teste].copy()
    for chave_av, met in [("av4", MET_AV4), ("av1", MET_AV1), ("av2", MET_AV2), ("av3", MET_AV3)]:
        d = golds[chave_av]
        d = d[d["tipoTeste"] == tipo_teste]
        cols = CHAVE + [c for c in met if c in d.columns]
        tab = tab.merge(d[cols], on=CHAVE, how="left")
    return tab.reset_index(drop=True)


def _unificar_detalhe(detalhes: dict) -> pd.DataFrame:
    """Empilha os detalhes long (av3 alturas + av4 componentes) numa base única (planta×metrica×valor)."""
    blocos = []
    for aval, df in detalhes.items():
        cols = [c for c in COLS_DETALHE if c in df.columns]
        b = df[cols].copy()
        b.insert(0, "avaliacao", aval)
        blocos.append(b)
    return pd.concat(blocos, ignore_index=True)


@st.cache_data(show_spinner="Carregando dados de milho 2025...")
def rodar_pipeline() -> dict:
    """Orquestra o pipeline 2025 ao vivo (Supabase) e devolve o modelo dimensional em cache.

    Retorna dict com:
      base_plots, tabela_analitica_faixa, tabela_analitica_densidade, base_detalhe,
      detalhe_fotos (fotos/comentários de campo por avaliação),
      e os golds por avaliação (av1..av4) para páginas que precisem do grão fino.
    """
    supabase = get_supabase_2025()

    # 1) Extração (bronze) + silver
    dfs_cru = {t: _extrair(supabase, t) for t in TABELAS}
    apoio = _silver_apoio(dfs_cru)
    avs = _silver_avaliacoes(dfs_cru)

    # 2) Enriquecimento (contexto: regiões, localidade, cod_fazenda, materiais)
    df_fazenda = _enriquecer_fazenda(apoio["fazenda"], apoio["cidade"], apoio["estado"], safra="25/26")
    df_tb = _enriquecer_tratamento(apoio["tratamentoBase"], dep_materiais=carregar_depara_materiais_2025())
    enriquecer_av = _fazer_enriquecer_av(df_fazenda, apoio["users"], apoio["avaliacao"], df_tb)

    tb_av1 = enriquecer_av(avs["av1"])
    tb_av2 = enriquecer_av(avs["av2"])
    tb_av3 = enriquecer_av(avs["av3"])
    tb_av4 = enriquecer_av(avs["av4"])

    # fotos e comentários de campo (av1..av4). Precisa vir DEPOIS do enriquecer_av: o detalhe
    # liga pela linha da avaliação (tratamentoRef -> av.uuid), não pelo tratamentoBase.
    detalhe_fotos = _enriquecer_detalhe_fotos(
        _silver_detalhe(dfs_cru),
        {"av1": tb_av1, "av2": tb_av2, "av3": tb_av3, "av4": tb_av4},
        df_fazenda)

    # 3) Gold por avaliação
    tb_av1_gold = _gold_av1(tb_av1)
    tb_av2_gold = _gold_av2(tb_av2)
    tb_av3_gold = _gold_av3(tb_av3)
    tb_av3_detalhe = _gold_av3_detalhe(tb_av3)
    tb_av4_gold, tb_av4_detalhe = _gold_av4(tb_av4)

    golds = {"av1": tb_av1_gold, "av2": tb_av2_gold, "av3": tb_av3_gold, "av4": tb_av4_gold}

    # 4) Consolidação (modelo dimensional)
    base_plots = _montar_base_plots(golds)
    tab_faixa     = _consolidar_tipo("Faixa", base_plots, golds)
    tab_densidade = _consolidar_tipo("Densidade", base_plots, golds)
    base_detalhe  = _unificar_detalhe({"av3": tb_av3_detalhe, "av4": tb_av4_detalhe})

    return {
        "base_plots": base_plots,
        "tabela_analitica_faixa": tab_faixa,
        "tabela_analitica_densidade": tab_densidade,
        "base_detalhe": base_detalhe,
        "detalhe_fotos": detalhe_fotos,
        "av1_gold": tb_av1_gold, "av2_gold": tb_av2_gold,
        "av3_gold": tb_av3_gold, "av4_gold": tb_av4_gold,
        "av3_detalhe": tb_av3_detalhe, "av4_detalhe": tb_av4_detalhe,
        # dados de apoio para o Diagnóstico (integridade estrutural):
        "tratamento_base": df_tb.assign(safra="25/26"),   # catálogo de tratamentos + safra
        "fazendas": df_fazenda,        # fazendas enriquecidas (responsável, cidade, região)
    }
