"""
pages/8_Fotos_Comentarios.py — Fotos e Comentários de Campo (MILHO)

Espelha a página de soja, com quatro diferenças de conteúdo:
  - o milho tem 4 avaliações (av1..av4), não 7;
  - o de-para de materiais vem do pipeline, então o card mostra o nome canônico;
  - os filtros seguem o padrão do painel de milho: nada marcado = filtro desligado
    (na soja, nada marcado também mostra tudo, mas por caminho diferente);
  - a fonte é `detalhe_fotos`, exposta pelo pipeline (ver patch: _silver_detalhe existia
    mas não era chamada, e as fotos nunca chegavam ao painel).
"""
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from utils.theme import aplicar_tema, page_header, secao_titulo, rodape
from utils.loader import carregar_multisafra
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

st.set_page_config(
    page_title="Fotos e Comentários · JAUM DTC Milho",
    page_icon="📷",
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

AG_CSS = {
    ".ag-header":                  {"background-color": "#4A4A4A !important"},
    ".ag-header-row":              {"background-color": "#4A4A4A !important"},
    ".ag-header-cell":             {"background-color": "#4A4A4A !important"},
    ".ag-header-cell-label":       {"color": "#FFFFFF !important", "font-weight": "700"},
    ".ag-header-cell-text":        {"color": "#FFFFFF !important", "font-size": "13px !important",
                                    "font-weight": "700 !important"},
    ".ag-icon":                    {"color": "#FFFFFF !important", "opacity": "1 !important"},
    ".ag-header-icon":             {"color": "#FFFFFF !important", "opacity": "1 !important"},
    ".ag-header-cell-menu-button": {"opacity": "1 !important", "visibility": "visible !important"},
    ".ag-icon-menu":               {"color": "#FFFFFF !important", "opacity": "1 !important"},
    ".ag-icon-filter":             {"color": "#FFFFFF !important", "opacity": "1 !important"},
    ".ag-cell":                    {"font-size": "13px !important"},
    ".ag-row":                     {"font-size": "13px !important"},
}

# cores de status, as mesmas do resto do painel de milho
COR_STATUS = {"CHECK": "#F4B184", "STINE": "#2976B6", "EXP": "#00FF00", "DP2": "#C4DFB4"}
COR_TEXTO_STATUS = {"CHECK": "#1A1A1A", "STINE": "#FFFFFF", "EXP": "#1A1A1A", "DP2": "#1A1A1A"}


def ag_table(df, height=400):
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        resizable=True, sortable=True, filter=True,
        cellStyle={"fontSize": "13px", "fontFamily": "Helvetica Neue, sans-serif",
                   "color": "#000000"},
        wrapText=True, autoHeight=True)
    gb.configure_grid_options(headerHeight=36, rowHeight=32, domLayout="normal",
                              suppressMenuHide=True)
    go = gb.build()
    go["defaultColDef"]["headerClass"] = "ag-header-black"
    AgGrid(df, gridOptions=go, height=height, update_mode=GridUpdateMode.NO_UPDATE,
           fit_columns_on_grid_load=False, allow_unsafe_jscode=True,
           enable_enterprise_modules=True, custom_css=AG_CSS,
           theme="streamlit", use_container_width=True)


# ── As 4 avaliações do milho ──────────────────────────────────────────────────
# O milho tem quatro, não sete como a soja. Os nomes seguem o que o pipeline documenta:
# av1 qualidade inicial, av2 sanidade, av3 arquitetura + florescimento, av4 colheita.
AV_NOMES = {
    "av1": "AV1 · Qualidade da Faixa Inicial",
    "av2": "AV2 · Sanidade",
    "av3": "AV3 · Arquitetura de Planta e Florescimento",
    "av4": "AV4 · Colheita",
}

page_header(
    "Fotos e Comentários de Campo",
    "Registros fotográficos e observações feitas pelos responsáveis DTC durante as avaliações "
    "de milho.",
    imagem="Taking notes-amico.png",
)

# ── Carregamento ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Carregando registros de campo...")
def carregar_fotos():
    d = carregar_multisafra()
    df = d.get("detalhe_fotos")
    return df.copy() if df is not None else pd.DataFrame()


df_all = carregar_fotos()

if df_all.empty:
    st.error(
        "Nenhum registro de foto ou comentário disponível.\n\n"
        "A tabela `detalhe_fotos` não veio do pipeline. No `pipeline_milho_2025.py` a função "
        "`_silver_detalhe` existe mas não é chamada dentro de `rodar_pipeline` — é ela que trata "
        "`photoUrl` e `nota`. Ver o patch entregue junto com esta página."
    )
    st.stop()

# só interessa o que tem foto OU comentário; o resto é linha vazia de detalhe
_tem_algo = pd.Series(False, index=df_all.index)
for _c in ["photoUrl", "nota"]:
    if _c in df_all.columns:
        _tem_algo |= df_all[_c].notna() & (df_all[_c].astype(str).str.strip() != "")
df_all = df_all[_tem_algo].copy()

if df_all.empty:
    st.warning("Nenhuma parcela com foto ou comentário registrado nas safras carregadas.")
    st.stop()

# ── Filtros (mesmo padrão das demais páginas de milho) ────────────────────────
with st.sidebar:
    st.markdown('<p style="font-size:11px;font-weight:600;color:#6B7280;text-transform:uppercase;'
                'letter-spacing:0.05em;padding:0.5rem;">Filtros</p>', unsafe_allow_html=True)

    if st.button("🔄 Limpar filtros", use_container_width=True, key="fc_btn_limpar"):
        for _k in list(st.session_state.keys()):
            if str(_k).startswith("fc_") and not str(_k).endswith("_btn_limpar"):
                del st.session_state[_k]
        st.rerun()

    def _checkboxes(opcoes, prefix, defaults=None):
        """Nada marcado = filtro desligado. Mesma regra do resto do painel de milho:
        zerar a base quando o usuário desmarca tudo só produz tela vazia sem motivo."""
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

    def _filtro_busca(opcoes, prefix):
        """Busca textual para listas longas; a seleção sobrevive à busca."""
        if f"{prefix}_reset" not in st.session_state:
            st.session_state[f"{prefix}_reset"] = 0
        r = st.session_state[f"{prefix}_reset"]
        busca = st.text_input("Buscar", value="", key=f"fc_busca_{prefix}",
                              placeholder="Digite para filtrar...")
        filtradas = ([c for c in opcoes if busca.strip().lower() in str(c).lower()]
                     if busca.strip() else opcoes)
        if st.button("Limpar seleção", key=f"{prefix}_limpar", use_container_width=True):
            for o in opcoes:
                st.session_state.pop(f"{prefix}_chk_{r}_{o}", None)
            st.session_state[f"{prefix}_reset"] = r + 1
            st.rerun()
        for c in filtradas:
            st.checkbox(str(c), key=f"{prefix}_chk_{r}_{c}")
        sel = [o for o in opcoes if st.session_state.get(f"{prefix}_chk_{r}_{o}", False)]
        return sel or opcoes

    d1 = df_all.copy()

    # Safra — padrão na mais recente, como nas outras páginas
    if "safra" in d1.columns:
        def _ano(s):
            try:
                return int(str(s).split("/")[0])
            except Exception:
                return -1
        _safras = sorted(d1["safra"].dropna().unique().tolist())
        _def = sorted(_safras, key=_ano)[-1:] if _safras else []
        with st.expander("📅 Safra", expanded=True):
            _sel = _checkboxes(_safras, "fc_safra", defaults=_def)
        if _sel:
            d1 = d1[d1["safra"].isin(_sel)]

    _CONFIG = [
        ("estado_sigla", "🏛️ Estado", "fc_estado", False),
        ("cidade_nome", "🏙️ Cidade", "fc_cidade", True),
        ("nomeFazenda", "🚜 Fazenda", "fc_fazenda", True),
        ("nomeResponsavel", "👤 Responsável", "fc_resp", True),
        ("tipoTeste", "🔬 Tipo de Teste", "fc_tipo", False),
        ("status_material", "🏷️ Status", "fc_status", False),
        ("dePara", "🌽 Híbrido", "fc_hib", True),
    ]
    for _col, _lab, _pref, _busca in _CONFIG:
        if _col not in d1.columns:
            continue
        _ops = sorted(d1[_col].dropna().unique().tolist())
        if not _ops:
            continue
        with st.expander(_lab, expanded=False):
            _sel = _filtro_busca(_ops, _pref) if _busca else _checkboxes(_ops, _pref)
        _sel_ok = [v for v in (_sel or []) if v in _ops]
        if _sel_ok:
            d1 = d1[d1[_col].isin(_sel_ok)]

    st.markdown("---")
    so_foto = st.checkbox("Somente com foto", value=False, key="fc_so_foto")
    so_nota = st.checkbox("Somente com comentário", value=False, key="fc_so_nota")

df_filtrado = d1.copy()

if df_filtrado.empty:
    st.warning("Nenhum registro nos filtros ativos.")
    if st.button("🔄 Limpar todos os filtros", key="fc_limpar_inline"):
        for _k in list(st.session_state.keys()):
            if str(_k).startswith("fc_"):
                del st.session_state[_k]
        st.rerun()
    st.stop()

# ── Contexto do recorte, no padrão das outras páginas ─────────────────────────
_ctx = []
if "safra" in df_filtrado.columns:
    _ctx.append("Safra: " + ", ".join(sorted(df_filtrado["safra"].dropna().unique().astype(str))))
for _c, _rot, _pl in [("estado_sigla", "UF", "UFs"), ("cidade_nome", "Cidade", "cidades"),
                      ("cod_fazenda", "Local", "locais")]:
    if _c not in df_filtrado.columns:
        continue
    _v = sorted(df_filtrado[_c].dropna().unique().astype(str))
    if not _v:
        continue
    _ctx.append(f"{_rot}: " + ", ".join(_v) if len(_v) <= 3
                else f"{len(_v)} {_pl}: " + ", ".join(_v))
contexto_str = " · ".join(_ctx)

_n_foto = int(df_filtrado["photoUrl"].notna().sum()) if "photoUrl" in df_filtrado else 0
_n_nota = int(df_filtrado["nota"].notna().sum()) if "nota" in df_filtrado else 0
st.caption(f"{len(df_filtrado)} registros no recorte · {_n_foto} com foto · "
           f"{_n_nota} com comentário")

with st.popover("ℹ️ Como usar esta página", use_container_width=False):
    st.markdown("""
**O que você vê aqui.** Cada foto e cada comentário registrado pelo responsável DTC no aplicativo
de campo, durante as avaliações. É material de consulta e de apresentação — não entra em conta
nenhuma das outras páginas.

**As abas são as quatro avaliações do milho**, na ordem do ciclo:

- **AV1 · Qualidade da Faixa Inicial** — logo depois da emergência
- **AV2 · Sanidade** — notas de doença
- **AV3 · Arquitetura de Planta e Florescimento**
- **AV4 · Colheita** — componentes de produção e perdas

O número ao lado do nome da aba é quantos registros ela tem no recorte atual, então dá para ver
de fora onde há material antes de entrar.

**Os filtros da barra lateral** são encadeados: escolher um estado reduz as cidades, escolher a
cidade reduz as fazendas. **Filtro sem nada marcado está desligado** e mostra tudo — não é preciso
marcar todos. A safra abre na mais recente.

Os dois últimos, **Somente com foto** e **Somente com comentário**, valem para todas as abas.

**Na galeria**, clique em qualquer foto para abrir em tamanho grande, com botão de baixar. `Esc`
ou clique fora fecham. O selo colorido no canto do card é o status do material, na mesma legenda
do resto do painel: laranja CHECK, azul STINE, verde EXP.

**Fotos que não carregam somem do mural.** Se o link tiver expirado no armazenamento, o card não
aparece em vez de mostrar um quadrado quebrado. Por isso a contagem no topo pode ser maior que o
número de cards visíveis.

**A tabela de comentários** aceita ordenação pelo cabeçalho e filtro pelo funil, e é o caminho
para procurar por palavra dentro dos textos.
""")

# ── Abas por avaliação ────────────────────────────────────────────────────────
_avs = [av for av in AV_NOMES if (df_filtrado["avaliacao"] == av).any()]
if not _avs:
    st.info("Nenhuma avaliação com registros no recorte.")
    st.stop()

_labels = [f"{AV_NOMES[av]}  ({int((df_filtrado['avaliacao'] == av).sum())})" for av in _avs]
_tabs = st.tabs(_labels)

for _tab, _av in zip(_tabs, _avs):
    with _tab:
        df_av = df_filtrado[df_filtrado["avaliacao"] == _av].copy()

        if so_foto and "photoUrl" in df_av.columns:
            df_av = df_av[df_av["photoUrl"].notna() & (df_av["photoUrl"].astype(str).str.strip() != "")]
        if so_nota and "nota" in df_av.columns:
            df_av = df_av[df_av["nota"].notna() & (df_av["nota"].astype(str).str.strip() != "")]

        if df_av.empty:
            st.info("Nenhum registro para os filtros ativos nesta avaliação.")
            continue

        # ── Galeria ───────────────────────────────────────────────────────────
        df_fotos = (df_av[df_av["photoUrl"].notna()
                          & (df_av["photoUrl"].astype(str).str.strip() != "")]
                    if "photoUrl" in df_av.columns else df_av.iloc[0:0])

        if not df_fotos.empty:
            secao_titulo("Galeria de Fotos",
                         f"{len(df_fotos)} registros fotográficos nesta avaliação", contexto_str)

            if "dataCriacao" in df_fotos.columns:
                df_fotos = df_fotos.sort_values("dataCriacao", ascending=False)

            # NaN do pandas é TRUTHY: `rec.get("dePara") or "—"` devolve o próprio NaN, que
            # vira a string "nan" no card. Por isso todo campo passa por _txt(), que trata
            # NaN, None, "" e o literal "nan" como vazio.
            def _txt(v, vazio=""):
                # pd.isna cobre NaN, None, NaT e pd.NA de uma vez; o teste de string pega o
                # literal "nan" que sobra quando algo virou texto antes de chegar aqui
                try:
                    if v is None or pd.isna(v):
                        return vazio
                except (TypeError, ValueError):
                    pass
                t = str(v).strip()
                return vazio if t.lower() in ("", "nan", "nat", "none", "null", "<na>") else t

            def _esc(v):
                return _txt(v).replace("'", "&#39;").replace('"', "&quot;")

            _cards = ""
            _sem_material = 0
            for _, rec in df_fotos.iterrows():
                # o material vem do join com tratamentoBase; quando a foto não tem tratamento
                # vinculado, o join não casa e não há cultivar a mostrar
                cultivar = _esc(rec.get("dePara")) or _esc(rec.get("nome"))
                if not cultivar:
                    cultivar = "Sem material vinculado"
                    _sem_material += 1
                fazenda = _esc(rec.get("nomeFazenda")) or "—"
                cod = _esc(rec.get("cod_fazenda"))
                status = _txt(rec.get("status_material"))
                _dt = pd.to_datetime(rec.get("dataCriacao"), errors="coerce")
                data_str = _dt.strftime("%d/%m/%Y") if pd.notna(_dt) else "—"
                nota_str = _esc(rec.get("nota"))
                url = _txt(rec.get("photoUrl")).replace("'", "%27")

                _badge = ""
                if status:
                    _badge = (f'<span class="status-badge" style="background:'
                              f'{COR_STATUS.get(status, "#E5E7EB")};color:'
                              f'{COR_TEXTO_STATUS.get(status, "#1A1A1A")};">{status}</span>')
                _nota_html = f'<div class="nota-badge">{nota_str}</div>' if nota_str else ""
                _local = fazenda                      # rodapé do card: só o nome
                _local_modal = fazenda + (f" · {cod}" if cod else "")   # modal: com o código
                _cls_cult = "" if cultivar != "Sem material vinculado" else " sem-mat"

                _cards += f"""
<div class="foto-card" onclick="abrir('{url}','{cultivar}','{_local_modal}','{data_str}')">
  <div class="img-wrap">
    <img src="{url}" alt="{cultivar}"
         onerror="this.closest('.foto-card').style.display='none'"/>
    {_badge}
  </div>
  <div class="foto-info">
    <div class="foto-cultivar{_cls_cult}">{cultivar}</div>
    <div class="foto-fazenda">{_local}</div>
    <div class="foto-data">{data_str}</div>
    {_nota_html}
  </div>
</div>"""

            _html = f"""
<!DOCTYPE html><html><head><style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:'Helvetica Neue',sans-serif;}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:14px;padding:4px;}}
.foto-card{{border:1px solid #E5E7EB;border-radius:10px;overflow:hidden;background:#fff;
  cursor:pointer;transition:transform .15s,box-shadow .15s;}}
.foto-card:hover{{transform:translateY(-2px);box-shadow:0 4px 16px rgba(0,0,0,.12);}}
.img-wrap{{position:relative;}}
.foto-card img{{width:100%;height:180px;object-fit:cover;display:block;background:#F3F4F6;}}
.status-badge{{position:absolute;top:8px;left:8px;padding:2px 8px;border-radius:4px;
  font-size:11px;font-weight:700;letter-spacing:.02em;}}
.foto-info{{padding:8px 10px;}}
.foto-cultivar{{font-weight:700;font-size:13px;color:#111827;margin-bottom:2px;}}
.foto-cultivar.sem-mat{{color:#9CA3AF;font-weight:500;font-style:italic;}}
.foto-fazenda{{font-size:12px;color:#6B7280;margin-bottom:2px;}}
.foto-data{{font-size:11px;color:#9CA3AF;}}
.nota-badge{{background:#F3F4F6;border-left:3px solid #2976B6;padding:4px 8px;
  border-radius:0 4px 4px 0;font-size:11px;color:#374151;margin-top:6px;font-style:italic;
  max-height:64px;overflow:auto;}}
.overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.88);z-index:9999;
  align-items:center;justify-content:center;flex-direction:column;}}
.overlay.on{{display:flex;}}
.overlay img{{max-width:90vw;max-height:78vh;object-fit:contain;border-radius:8px;
  box-shadow:0 8px 40px rgba(0,0,0,.5);}}
.ov-info{{color:#fff;text-align:center;margin-top:12px;}}
.ov-t{{font-size:18px;font-weight:700;}}
.ov-s{{font-size:13px;color:#D1D5DB;margin-top:4px;}}
.ov-x{{position:fixed;top:20px;right:28px;font-size:32px;color:#fff;cursor:pointer;
  line-height:1;opacity:.8;}}
.ov-x:hover{{opacity:1;}}
.ov-dl{{display:inline-block;margin-top:12px;padding:8px 20px;background:#2976B6;color:#fff;
  border-radius:6px;font-size:13px;font-weight:600;text-decoration:none;}}
</style></head><body>
<div class="grid">{_cards}</div>
<div class="overlay" id="ov" onclick="fechar(event)">
  <span class="ov-x" onclick="document.getElementById('ov').classList.remove('on')">✕</span>
  <img id="ov-img" src="" alt=""/>
  <div class="ov-info">
    <div class="ov-t" id="ov-t"></div>
    <div class="ov-s" id="ov-s"></div>
    <a class="ov-dl" id="ov-dl" href="" download="" target="_blank">⬇️ Baixar foto</a>
  </div>
</div>
<script>
function abrir(url,cultivar,local,data){{
  document.getElementById('ov-img').src=url;
  document.getElementById('ov-t').textContent=cultivar;
  document.getElementById('ov-s').textContent=local+' · '+data;
  var dl=document.getElementById('ov-dl');
  dl.href=url;
  dl.download=cultivar.replace(/[^a-zA-Z0-9]/g,'_')+'_'+data.replace(/[/]/g,'-')+'.jpg';
  document.getElementById('ov').classList.add('on');
}}
function fechar(e){{
  if(e.target===document.getElementById('ov'))
    document.getElementById('ov').classList.remove('on');
}}
document.addEventListener('keydown',function(e){{
  if(e.key==='Escape') document.getElementById('ov').classList.remove('on');
}});
</script>
</body></html>"""

            _linhas = (len(df_fotos) + 3) // 4
            components.html(_html, height=min(1400, max(280, _linhas * 300 + 40)),
                            scrolling=True)
            st.caption("Clique na foto para ampliar e baixar · `Esc` fecha · o selo no canto é o "
                       "status do material · fotos com link expirado somem do mural, por isso a "
                       "contagem pode ser maior que o número de cards.")
            if _sem_material:
                st.warning(
                    f"**{_sem_material} de {len(df_fotos)} fotos não têm material vinculado.** "
                    f"O `tratamentoRef` do registro não casa com nenhuma linha da avaliação — "
                    f"normalmente porque a foto foi tirada no nível do talhão, não da parcela. "
                    f"Ela continua utilizável pelo local e pela data, mas não dá para dizer de "
                    f"que híbrido é, e não aparece ao filtrar por híbrido.")

        # ── Comentários ───────────────────────────────────────────────────────
        df_notas = (df_av[df_av["nota"].notna()
                          & (df_av["nota"].astype(str).str.strip() != "")]
                    if "nota" in df_av.columns else df_av.iloc[0:0])

        if not df_notas.empty:
            secao_titulo("Comentários",
                         f"{len(df_notas)} observações escritas nesta avaliação", contexto_str)
            _map = {"safra": "Safra", "cod_fazenda": "Cód. Local", "nomeFazenda": "Fazenda",
                    "cidade_nome": "Cidade", "estado_sigla": "UF",
                    "nomeResponsavel": "Responsável", "dePara": "Híbrido",
                    "status_material": "Status", "tipoTeste": "Tipo",
                    "dataCriacao": "Data", "nota": "Comentário"}
            _cols = [c for c in _map if c in df_notas.columns]
            _show = df_notas[_cols].rename(columns=_map).copy()
            if "Data" in _show.columns:
                _show["Data"] = (pd.to_datetime(_show["Data"], errors="coerce")
                                 .dt.strftime("%d/%m/%Y").fillna(""))
                _show = _show.sort_values("Data", ascending=False)
            ag_table(_show, height=min(520, 36 + 42 * len(_show) + 20))

            import io
            _buf = io.BytesIO()
            with pd.ExcelWriter(_buf, engine="openpyxl") as _w:
                _show.to_excel(_w, index=False, sheet_name="Comentarios")
            st.download_button("⬇️ Exportar comentários", data=_buf.getvalue(),
                               file_name=f"fotos_comentarios_{_av}.xlsx",
                               mime=("application/vnd.openxmlformats-officedocument."
                                     "spreadsheetml.sheet"),
                               key=f"fc_xlsx_{_av}")

        if df_fotos.empty and df_notas.empty:
            st.info("Esta avaliação não tem foto nem comentário no recorte ativo.")

st.divider()
rodape()
