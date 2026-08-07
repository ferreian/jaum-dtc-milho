"""
utils/loader.py — amarra os dois pipelines de milho e empilha as safras.

Do zero (não copiado do soja), porque no milho AS DUAS SAFRAS RODAM AO VIVO (Supabase):
  - carregar_2024() → pipeline_milho_2024.rodar_pipeline()
  - carregar_2025() → pipeline_milho_2025.rodar_pipeline()
  - carregar_multisafra() → empilha as duas aplicando o depara_mestre.csv

Empilhamento (decisão de arquitetura): acontece AQUI, no painel, não no pipeline. Cada pipeline
produz as analíticas de UMA safra, fiel ao seu banco. O concat multi-safra é feito na hora de
exibir, com a coluna `safra` (que os pipelines já trazem). Schema aditivo: colunas só de 2025
(fenômenos da colheita, qualidade_plot_inicial, nota_empalhamento) viram NaN nas linhas de 2024.

De-para mestre (reconciliação de materiais entre safras):
  - dePara_original: preserva o canônico da safra (auditoria);
  - dePara: vira o canônico MESTRE (para os materiais casarem entre safras);
  - status_safra: preserva o status daquela safra (auditoria);
  - status_material: status_mestre (mais recente vence) — decisão do André.
"""

import pandas as pd
import streamlit as st

from pipeline_milho_2024 import rodar_pipeline as _rodar_2024
from pipeline_milho_2025 import rodar_pipeline as _rodar_2025
from pipeline_milho_2025 import carregar_depara_mestre

# Tabelas que empilham entre safras (todas têm coluna `safra` e `dePara`)
_TABELAS_MULTISAFRA = [
    "tabela_analitica_faixa",
    "tabela_analitica_densidade",
    "base_plots",
    "base_detalhe",
    # fotos e comentários de campo (av1..av4). Entra aqui porque tem `safra` e `dePara`,
    # então empilha e passa pelo de-para mestre como as demais — sem isso o mesmo híbrido
    # apareceria com nomes diferentes nas fotos de 24/25 e de 25/26.
    "detalhe_fotos",
]


# ── Carga por safra (cada uma roda o pipeline ao vivo, já cacheado) ────────────
@st.cache_data(show_spinner="Carregando milho 2024...")
def carregar_2024() -> dict:
    """Roda o pipeline 2024. Devolve {'ok': True, **dados} ou {'ok': False, 'erro': ...}."""
    try:
        return {"ok": True, **_rodar_2024()}
    except Exception as e:
        import traceback
        return {"ok": False, "erro": str(e), "traceback": traceback.format_exc()}


@st.cache_data(show_spinner="Carregando milho 2025...")
def carregar_2025() -> dict:
    """Roda o pipeline 2025. Devolve {'ok': True, **dados} ou {'ok': False, 'erro': ...}."""
    try:
        return {"ok": True, **_rodar_2025()}
    except Exception as e:
        import traceback
        return {"ok": False, "erro": str(e), "traceback": traceback.format_exc()}


# ── De-para mestre: mapas de reconciliação ────────────────────────────────────
@st.cache_data(show_spinner=False)
def _mapas_mestre() -> tuple:
    """Constrói (mapa_canonico, mapa_status) do depara_mestre.csv.
      - mapa_canonico: dePara_safra -> dePara_mestre (unifica grafias entre safras);
      - mapa_status:   dePara_mestre -> status_mestre (mais recente vence).
    Se o mestre não existir, devolve mapas vazios (empilha sem reconciliar)."""
    m = carregar_depara_mestre()
    if m is None or m.empty:
        return {}, {}
    m = m.copy()
    for col in ["dePara_safra", "dePara_mestre", "status_mestre"]:
        m[col] = m[col].astype(str).str.strip()
    mapa_canonico = dict(zip(m["dePara_safra"], m["dePara_mestre"]))
    mapa_status = dict(zip(m["dePara_mestre"], m["status_mestre"]))
    return mapa_canonico, mapa_status


def aplicar_mestre(df: pd.DataFrame, mapa_canonico: dict, mapa_status: dict) -> pd.DataFrame:
    """Aplica o de-para mestre a uma tabela com coluna `dePara`.
      - dePara_original: canônico da safra (auditoria);
      - dePara: canônico MESTRE (materiais casam entre safras);
      - status_safra: status da safra (auditoria);
      - status_material: status_mestre (mais recente vence).
    Materiais fora do mestre ficam intactos (canônico e status originais).

    Idempotente: se a tabela já passou pelo mestre (tem `dePara_original`), os valores
    ORIGINAIS da safra são preservados — reaplicar não sobrescreve o histórico."""
    if df is None or df.empty or "dePara" not in df.columns:
        return df
    out = df.copy()

    # só grava o original na PRIMEIRA aplicação; senão o nome da safra se perderia
    if "dePara_original" not in out.columns:
        out["dePara_original"] = out["dePara"]
    canon = out["dePara"].astype(str).str.strip()
    out["dePara"] = canon.map(mapa_canonico).fillna(out["dePara"])

    if "status_material" in out.columns:
        if "status_safra" not in out.columns:
            out["status_safra"] = out["status_material"]                   # auditoria (por safra)
        novo = out["dePara"].astype(str).str.strip().map(mapa_status)
        out["status_material"] = novo.fillna(out["status_material"])       # mestre; fora do mestre mantém
    return out


# ── Empilhamento multi-safra ──────────────────────────────────────────────────
@st.cache_data(show_spinner="Empilhando safras de milho...")
def carregar_multisafra() -> dict:
    """Carrega as duas safras, aplica o de-para mestre e empilha (concat + coluna `safra`).

    Retorna dict com:
      - ok: True se ao menos uma safra carregou;
      - safras_ok: lista das safras que carregaram ('24/25', '25/26');
      - erros: {safra: mensagem} das que falharam;
      - as 5 tabelas empilhadas (faixa, densidade, base_plots, base_detalhe, detalhe_fotos);
      - '2024'/'2025': os dicts brutos por safra (para páginas que precisem de uma safra só).
    """
    d24 = carregar_2024()
    d25 = carregar_2025()
    mapa_canonico, mapa_status = _mapas_mestre()

    def _empilhar(chave: str) -> pd.DataFrame:
        frames = []
        for d in (d24, d25):
            tab = d.get(chave) if d.get("ok") else None
            if isinstance(tab, pd.DataFrame) and not tab.empty:
                frames.append(aplicar_mestre(tab, mapa_canonico, mapa_status))
        # concat alinha por nome; colunas só-de-2025 viram NaN nas linhas de 2024
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    out = {
        "ok": bool(d24.get("ok") or d25.get("ok")),
        "safras_ok": [s for s, d in [("24/25", d24), ("25/26", d25)] if d.get("ok")],
        "erros": {s: d.get("erro") for s, d in [("24/25", d24), ("25/26", d25)] if not d.get("ok")},
        "2024": d24,
        "2025": d25,
    }
    for chave in _TABELAS_MULTISAFRA:
        out[chave] = _empilhar(chave)
    return out
