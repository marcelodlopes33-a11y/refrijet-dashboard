import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re

st.set_page_config(
    page_title="Refrijet — Desenvolvimento de Produtos",
    page_icon="❄️",
    layout="wide",
)

# ── Estilos ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; padding-bottom: 1rem; }
    .metric-card {
        background: #f8f9fa; border-radius: 10px;
        padding: 16px 20px; border-left: 4px solid;
    }
    .metric-title { font-size: 11px; color: #6c757d; text-transform: uppercase; letter-spacing: .05em; margin-bottom: 4px; }
    .metric-value { font-size: 28px; font-weight: 600; line-height: 1; }
    .metric-sub   { font-size: 11px; color: #6c757d; margin-top: 4px; }
    div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
    .stDownloadButton > button { width: 100%; }
</style>
""", unsafe_allow_html=True)

FAMILIA_CORES = {
    "21 Compressores":            "#1f77b4",
    "22 Condensadores":           "#f59e0b",
    "44 Evaporadores":            "#10b981",
    "33 Comp. p/ Compressores":   "#ef4444",
    "30 Válvulas e Filtros":      "#8b5cf6",
    "35 Gases":                   "#ec4899",
    "29 Eletroventiladores":      "#14b8a6",
}

# ── Leitura e limpeza ─────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def carregar(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes), sheet_name="Solicitações", header=1)
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["N° ORDEM"])
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df["QTD MENSAL"] = pd.to_numeric(df["QTD MENSAL"], errors="coerce").fillna(0).astype(int)
    df["PREÇO MERCADO"] = pd.to_numeric(df["PREÇO MERCADO"], errors="coerce")
    for col in ["FILIAL", "SOLICITANTE", "FAMILIA", "MONTADORA", "MODELO"]:
        if col in df.columns:
            df[col] = df[col].fillna("Não informado").str.strip()
    df["ANO"] = df["ANO"].fillna("—").astype(str).str.replace(r"\.0$", "", regex=True)
    df["COD. ROYCE"] = df["COD. ROYCE"].fillna("—").astype(str).str.replace(r"\.0$", "", regex=True)
    df["COD. HDS"]   = df["COD. HDS"].fillna("—").astype(str).str.strip()
    df["COD. OEM"]   = df["COD. OEM"].fillna("—").astype(str).str.strip()
    # família curta para exibição
    df["FAM_CURTA"] = df["FAMILIA"].apply(
        lambda x: re.sub(r"^\d+\s*", "", str(x)) if pd.notna(x) else x
    )
    return df


def kpi_card(titulo, valor, subtitulo, cor):
    return f"""
    <div class="metric-card" style="border-color:{cor}">
        <div class="metric-title">{titulo}</div>
        <div class="metric-value" style="color:{cor}">{valor}</div>
        <div class="metric-sub">{subtitulo}</div>
    </div>"""


def exportar_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Solicitações")
        wb  = writer.book
        ws  = writer.sheets["Solicitações"]
        hdr = wb.add_format({"bold": True, "bg_color": "#1f4e79", "font_color": "white",
                              "border": 1, "align": "center"})
        cel = wb.add_format({"border": 1, "align": "left"})
        num = wb.add_format({"border": 1, "align": "center"})
        for col_num, col_name in enumerate(df.columns):
            ws.write(0, col_num, col_name, hdr)
        for row_num, row in enumerate(df.itertuples(index=False), start=1):
            for col_num, val in enumerate(row):
                fmt = num if isinstance(val, (int, float)) else cel
                ws.write(row_num, col_num, val if pd.notna(val) else "", fmt)
        for i, col in enumerate(df.columns):
            w = max(len(str(col)), df[col].astype(str).str.len().max())
            ws.set_column(i, i, min(w + 2, 40))

        # aba resumo
        resumo = pd.DataFrame({
            "Família": df.groupby("FAMILIA")["QTD MENSAL"].sum().sort_values(ascending=False).index,
            "Qtd. Total": df.groupby("FAMILIA")["QTD MENSAL"].sum().sort_values(ascending=False).values,
            "Solicitações": df.groupby("FAMILIA")["N° ORDEM"].count().reindex(
                df.groupby("FAMILIA")["QTD MENSAL"].sum().sort_values(ascending=False).index).values,
        })
        resumo.to_excel(writer, index=False, sheet_name="Resumo por Família")
    return buf.getvalue()


# ── Interface ─────────────────────────────────────────────────────────────────
st.markdown("## ❄️ Refrijet — Desenvolvimento de Produtos")
st.markdown("Faça o upload da planilha de solicitações para gerar o dashboard automaticamente.")

uploaded = st.file_uploader(
    "Arraste ou selecione o arquivo Excel (.xlsx)",
    type=["xlsx"],
    label_visibility="collapsed",
)

if not uploaded:
    st.info("Aguardando o arquivo... Use o botão acima para selecionar o Excel de Solicitações.")
    st.stop()

with st.spinner("Carregando dados..."):
    df_raw = carregar(uploaded.read())

st.success(f"✅  {len(df_raw)} solicitações carregadas com sucesso.")
st.divider()

# ── Filtros laterais ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔍 Filtros")

    filiais = ["Todas"] + sorted(df_raw["FILIAL"].unique().tolist())
    sel_filial = st.selectbox("Filial", filiais)

    familias = ["Todas"] + sorted(df_raw["FAMILIA"].unique().tolist())
    sel_familia = st.selectbox("Família de produto", familias)

    montadoras = ["Todas"] + sorted(df_raw["MONTADORA"].unique().tolist())
    sel_montadora = st.selectbox("Montadora", montadoras)

    solicitantes = ["Todos"] + sorted(df_raw["SOLICITANTE"].unique().tolist())
    sel_solicitante = st.selectbox("Solicitante", solicitantes)

    datas = df_raw["DATA"].dropna()
    if not datas.empty:
        d_min, d_max = datas.min().date(), datas.max().date()
        intervalo = st.date_input("Período", value=(d_min, d_max),
                                  min_value=d_min, max_value=d_max)
    else:
        intervalo = None

    st.divider()
    st.markdown("### 📥 Exportar")
    excel_bytes = exportar_excel(df_raw)
    st.download_button(
        "⬇️ Baixar Excel filtrado",
        data=excel_bytes,
        file_name="refrijet_solicitacoes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ── Aplicar filtros ───────────────────────────────────────────────────────────
df = df_raw.copy()
if sel_filial    != "Todas":  df = df[df["FILIAL"]      == sel_filial]
if sel_familia   != "Todas":  df = df[df["FAMILIA"]     == sel_familia]
if sel_montadora != "Todas":  df = df[df["MONTADORA"]   == sel_montadora]
if sel_solicitante != "Todos": df = df[df["SOLICITANTE"] == sel_solicitante]
if intervalo and len(intervalo) == 2:
    df = df[df["DATA"].between(pd.Timestamp(intervalo[0]), pd.Timestamp(intervalo[1]))]

# ── KPIs ──────────────────────────────────────────────────────────────────────
total_sol = len(df)
total_qtd = int(df["QTD MENSAL"].sum())
n_familias = df["FAMILIA"].nunique()
n_solicitantes = df["SOLICITANTE"].nunique()

c1, c2, c3, c4 = st.columns(4)
c1.markdown(kpi_card("Total de solicitações", total_sol, "registros no período", "#1f77b4"), unsafe_allow_html=True)
c2.markdown(kpi_card("Possíveis vendas perdidas", f"{total_qtd:,}".replace(",", "."), "unidades/mês sem o produto", "#f59e0b"), unsafe_allow_html=True)
c3.markdown(kpi_card("Famílias solicitadas", n_familias, "categorias de produto", "#10b981"), unsafe_allow_html=True)
c4.markdown(kpi_card("Solicitantes ativos", n_solicitantes, "vendedores / filiais", "#8b5cf6"), unsafe_allow_html=True)

st.divider()

# ── Gráficos — linha 1 ────────────────────────────────────────────────────────
col_a, col_b = st.columns([3, 2])

with col_a:
    st.markdown("##### Qtd. mensal por família de produto")
    fam_qtd = (df.groupby("FAMILIA")["QTD MENSAL"].sum()
                 .reset_index().sort_values("QTD MENSAL", ascending=True))
    fam_qtd["cor"] = fam_qtd["FAMILIA"].map(FAMILIA_CORES).fillna("#94a3b8")
    fig_bar = px.bar(
        fam_qtd, x="QTD MENSAL", y="FAMILIA", orientation="h",
        color="FAMILIA", color_discrete_map=FAMILIA_CORES,
        text="QTD MENSAL",
    )
    fig_bar.update_traces(textposition="outside", marker_line_width=0)
    fig_bar.update_layout(showlegend=False, margin=dict(l=0, r=20, t=10, b=10),
                          height=280, yaxis_title="", xaxis_title="Unidades/mês")
    st.plotly_chart(fig_bar, use_container_width=True)

with col_b:
    st.markdown("##### Solicitações por filial")
    filial_cnt = df["FILIAL"].value_counts().reset_index()
    filial_cnt.columns = ["Filial", "Qtd"]
    fig_pie = px.pie(filial_cnt, names="Filial", values="Qtd",
                     color_discrete_sequence=px.colors.qualitative.Set2,
                     hole=0.4)
    fig_pie.update_traces(textposition="inside", textinfo="percent+label")
    fig_pie.update_layout(showlegend=False, margin=dict(l=0, r=0, t=10, b=10), height=280)
    st.plotly_chart(fig_pie, use_container_width=True)

# ── Gráficos — linha 2 ────────────────────────────────────────────────────────
col_c, col_d = st.columns([2, 3])

with col_c:
    st.markdown("##### Solicitantes — Nº de solicitações")
    sol_cnt = (df["SOLICITANTE"].value_counts().head(10)
                 .reset_index().rename(columns={"index": "Solicitante", "SOLICITANTE": "Qtd"}))
    sol_cnt.columns = ["Solicitante", "Qtd"]
    fig_sol = px.bar(sol_cnt.sort_values("Qtd"), x="Qtd", y="Solicitante",
                     orientation="h", text="Qtd",
                     color_discrete_sequence=["#6366f1"])
    fig_sol.update_traces(textposition="outside", marker_line_width=0)
    fig_sol.update_layout(showlegend=False, margin=dict(l=0, r=20, t=10, b=10),
                          height=320, yaxis_title="", xaxis_title="Solicitações")
    st.plotly_chart(fig_sol, use_container_width=True)

with col_d:
    st.markdown("##### Top 15 montadoras — Qtd. mensal acumulada")
    mo_qtd = (df.groupby("MONTADORA")["QTD MENSAL"].sum()
                .reset_index().sort_values("QTD MENSAL", ascending=False).head(15))
    fig_mo = px.bar(mo_qtd.sort_values("QTD MENSAL"), x="QTD MENSAL", y="MONTADORA",
                    orientation="h", text="QTD MENSAL",
                    color_discrete_sequence=["#0ea5e9"])
    fig_mo.update_traces(textposition="outside", marker_line_width=0)
    fig_mo.update_layout(showlegend=False, margin=dict(l=0, r=20, t=10, b=10),
                          height=320, yaxis_title="", xaxis_title="Unidades/mês")
    st.plotly_chart(fig_mo, use_container_width=True)

# ── Evolução temporal ─────────────────────────────────────────────────────────
if df["DATA"].notna().any():
    st.markdown("##### Evolução diária de solicitações")
    daily = (df.dropna(subset=["DATA"])
               .groupby(df["DATA"].dt.date)["QTD MENSAL"].sum()
               .reset_index().rename(columns={"DATA": "Data", "QTD MENSAL": "Qtd"}))
    fig_line = px.line(daily, x="Data", y="Qtd", markers=True,
                       color_discrete_sequence=["#1f77b4"])
    fig_line.update_layout(margin=dict(l=0, r=0, t=10, b=10), height=220,
                            yaxis_title="Qtd. mensal", xaxis_title="")
    st.plotly_chart(fig_line, use_container_width=True)

st.divider()

# ── Top 10 tabela ──────────────────────────────────────────────────────────────
st.markdown("##### 🏆 Top 10 — Itens mais solicitados")
top10 = (df.sort_values("QTD MENSAL", ascending=False)
           .head(10)
           [["N° ORDEM", "FAMILIA", "MONTADORA", "MODELO", "ANO",
             "COD. ROYCE", "COD. HDS", "QTD MENSAL", "PREÇO MERCADO"]]
           .rename(columns={"QTD MENSAL": "Qtd/mês", "PREÇO MERCADO": "Preço (R$)"}))
st.dataframe(top10, use_container_width=True, hide_index=True)

st.divider()

# ── Tabela completa ───────────────────────────────────────────────────────────
st.markdown(f"##### 📋 Todas as solicitações ({len(df)} registros)")

cols_exibir = ["N° ORDEM", "DATA", "FILIAL", "SOLICITANTE", "FAMILIA",
               "MONTADORA", "MODELO", "ANO", "COD. ROYCE", "COD. HDS",
               "COD. OEM", "QTD MENSAL", "PREÇO MERCADO", "OBS"]
cols_exibir = [c for c in cols_exibir if c in df.columns]

busca = st.text_input("🔎 Buscar em qualquer campo", placeholder="Ex: Renault, Compressor, RS...")
df_show = df[cols_exibir].copy()
if busca:
    mask = df_show.apply(lambda col: col.astype(str).str.contains(busca, case=False, na=False)).any(axis=1)
    df_show = df_show[mask]

st.dataframe(df_show, use_container_width=True, hide_index=True, height=380)

# download da seleção atual
excel_filtrado = exportar_excel(df[cols_exibir])
st.download_button(
    "⬇️  Exportar tabela atual para Excel",
    data=excel_filtrado,
    file_name="refrijet_filtrado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption("MDL Tech · Gestão e Inteligência de Dados")
