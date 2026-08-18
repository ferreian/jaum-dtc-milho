"""
pipeline_milho_2024.py
Pipeline da safra 2024 de milho (JAUM DTC) — REUSA o núcleo do pipeline_milho_2025.

Filosofia (decidida): não duplicar. O silver, o enriquecimento e o gold av1/av2/av3 são
idênticos entre safras e vêm por import do pipeline_milho_2025. Só o que MUDA por safra vive
aqui: a conexão (banco 2024), o de-para de materiais 2024, e o wrapper do gold av4 com as
RÉGUAS 2024 (peso ≥100→÷1000, PMG >40→÷10, estande das 5 subamostras; metragem 10 m — igual
a 2025). Os 3 retrofits (cod_fazenda, cobertura, pct de perda) já vivem no núcleo compartilhado,
então o 2024 os herda automaticamente.

Réguas 2024 (confirmadas no PRINCIPIOS_PIPELINE):
  - peso: rede de segurança ≥100 → ÷1000 (às vezes vinha em gramas); <100 já é kg;
  - PMG: rede de segurança >40 → ÷10 (vírgula deslocada, 231 = 23,1); ≤40 já correto;
  - estande: média das 5 subamostras planta{1..5}NumPlantas10metros;
  - subamostras: 5 (o protocolo mudou para 4 em 2025 — ver N_SUBAMOSTRAS_* no núcleo);
  - METROS_CONTAGEM = 10 (linear total; igual a 2025 — a diferença é nº de pontos, não a metragem).
"""

import pandas as pd
import streamlit as st
from supabase import create_client

# Núcleo compartilhado (funções puras de silver/enriquecimento/gold + réguas 2024)
from pipeline_milho_2025 import (
    _extrair, TABELAS,
    _silver_apoio, _silver_avaliacoes, _silver_detalhe, _enriquecer_detalhe_fotos,
    _enriquecer_fazenda, _enriquecer_tratamento, _fazer_enriquecer_av,
    _gold_av1, _gold_av2, _gold_av3, _gold_av3_detalhe,
    _gold_av4_produtividade, _gold_av4_detalhe, _gold_av4_consolidar,
    _montar_base_plots, _consolidar_tipo, _unificar_detalhe,
    _RELATORIO_DUPLICATAS, CHAVE_LOGICA,      # colapso de plots com cadastro duplicado
    regua_peso_2024, regua_pmg_2024, COLS_ESTANDE_5SUB,
    METROS_CONTAGEM, N_SUBAMOSTRAS_2024,
    CONFIG_DIR,
)

import numpy as np


# ── Conexão Supabase (projeto 2024) ───────────────────────────────────────────
@st.cache_resource
def get_supabase_2024():
    return create_client(
        st.secrets["supabase_2024"]["url"],
        st.secrets["supabase_2024"]["key"],
    )


# ── De-para de materiais 2024 (elo nome cru -> canônico da safra 24/25) ────────
@st.cache_data
def carregar_depara_materiais_2024():
    """De-para de materiais da safra 2024: nome (cru) -> dePara (canônico) -> status.

    Mesmo contrato do de-para 2025: `indexTratamento` marca EXCEÇÕES (materiais que dividem o
    nome no banco e só se distinguem pelo tratamento) e `tratamento_semente` guarda o tratamento
    industrial. O casamento em duas camadas vive no `_enriquecer_tratamento` do núcleo, então
    aqui basta ler as colunas. Em 2024 elas estão vazias/'padrao' até se confirmar se o
    9505VTPRO4 também tinha dois tratamentos no mesmo local.
    """
    path = CONFIG_DIR / "depara_materiais_2024.csv"
    _cols = ["nome", "dePara", "status_material", "indexTratamento", "tratamento_semente"]
    dep = pd.read_csv(path, usecols=lambda c: c in _cols)
    dep["nome"] = dep["nome"].astype(str).str.strip()
    for col in ["dePara", "status_material", "tratamento_semente"]:
        if col in dep.columns:
            dep[col] = dep[col].astype(str).str.strip().replace(
                {"": np.nan, "nan": np.nan, "None": np.nan})
    if "indexTratamento" not in dep.columns:
        dep["indexTratamento"] = np.nan
    dep["indexTratamento"] = pd.to_numeric(dep["indexTratamento"], errors="coerce")
    return dep


# ── Gold av4 2024 (mesmas 3 funções do núcleo, com as RÉGUAS 2024) ─────────────
def _gold_av4_2024(tb_av4: pd.DataFrame) -> tuple:
    """Orquestra o av4 da safra 2024. Retorna (consolidada, detalhe).
    Réguas 2024: peso ≥100→÷1000, PMG >40→÷10, estande das 5 subamostras, metragem 10 m."""
    prod = _gold_av4_produtividade(tb_av4, regua_peso=regua_peso_2024)         # 2024: ≥100→÷1000
    detalhe, _ = _gold_av4_detalhe(tb_av4, cols_estande=COLS_ESTANDE_5SUB,
                                   n_subamostras=N_SUBAMOSTRAS_2024,            # 2024: protocolo de 5
                                   regua_pmg=regua_pmg_2024)                    # 2024: PMG >40→÷10
    consolidada = _gold_av4_consolidar(prod, detalhe, metros_contagem=METROS_CONTAGEM)  # 10 m
    return consolidada, detalhe


# ── Orquestração (mesmo desenho do 2025, apontando para o banco/de-para 2024) ──
@st.cache_data(show_spinner="Carregando dados de milho 2024...")
def rodar_pipeline() -> dict:
    """Orquestra o pipeline 2024 ao vivo (Supabase 2024) e devolve o modelo dimensional em cache.

    Mesma saída do 2025 (schema aditivo): base_plots, tabela_analitica_faixa/densidade,
    base_detalhe, detalhe_fotos, e os golds por avaliação. As colunas novas de 2025 (qualidade_plot_inicial,
    nota_empalhamento) simplesmente não existem aqui — entram como ausentes no empilhamento."""
    supabase = get_supabase_2024()
    _RELATORIO_DUPLICATAS.clear()      # relatório desta execução (ver _colapsar_duplicatas)

    # 1) Extração (bronze) + silver — funções idênticas às de 2025
    dfs_cru = {t: _extrair(supabase, t) for t in TABELAS}
    apoio = _silver_apoio(dfs_cru)
    avs = _silver_avaliacoes(dfs_cru)

    # 2) Enriquecimento — safra 24/25 e de-para de materiais 2024
    df_fazenda = _enriquecer_fazenda(apoio["fazenda"], apoio["cidade"], apoio["estado"], safra="24/25")
    df_tb = _enriquecer_tratamento(apoio["tratamentoBase"], dep_materiais=carregar_depara_materiais_2024())
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

    # 3) Gold — av1/av2/av3 idênticos ao 2025; av4 com as réguas 2024
    tb_av1_gold = _gold_av1(tb_av1)   # sem qualidade_plot_inicial (não coletada em 2024) -> ausente
    tb_av2_gold = _gold_av2(tb_av2)   # sem empalhamento (novo em 2025) -> ausente
    tb_av3_gold = _gold_av3(tb_av3)
    tb_av3_detalhe = _gold_av3_detalhe(tb_av3)
    tb_av4_gold, tb_av4_detalhe = _gold_av4_2024(tb_av4)

    golds = {"av1": tb_av1_gold, "av2": tb_av2_gold, "av3": tb_av3_gold, "av4": tb_av4_gold}

    # 4) Consolidação (modelo dimensional) — mesmas funções do núcleo
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
        # plots que existiam em duplicidade e foram colapsados (cadastro repetido do ensaio):
        "duplicatas": (pd.concat(_RELATORIO_DUPLICATAS, ignore_index=True)
                       if _RELATORIO_DUPLICATAS else pd.DataFrame(
                           columns=CHAVE_LOGICA + ["cadastros", "linhas", "tabela"])),
        "tratamento_base": df_tb.assign(safra="24/25"),
        "fazendas": df_fazenda,
    }
