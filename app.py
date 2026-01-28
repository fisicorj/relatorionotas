import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import altair as alt

st.set_page_config(
    page_title="Evidência — Avaliação (Visualização Web)",
    layout="wide"
)

# =========================
# Utilidades
# =========================
EMAIL_RE = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")
WHITESPACE_RE = re.compile(r"\s+")

def guess_timestamp_col(cols):
    preferred = [
        "Carimbo de data/hora",
        "Timestamp",
        "Data/hora",
        "Data e hora"
    ]
    for p in preferred:
        for c in cols:
            if p.lower() in str(c).lower():
                return c
    return None

def is_email_col(colname: str) -> bool:
    c = str(colname).lower()
    return ("e-mail" in c) or ("email" in c)

def is_name_col(colname: str) -> bool:
    c = str(colname).lower()
    return c == "nome" or ("nome" in c and "seu" in c)

def is_comment_col(colname: str) -> bool:
    c = str(colname).lower()
    keys = [
        "coment",
        "sugest",
        "observa",
        "feedback",
        "melhoria",
        "deixe aqui"
    ]
    return any(k in c for k in keys)

def clean_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = EMAIL_RE.sub("[e-mail removido]", s)
    s = WHITESPACE_RE.sub(" ", s).strip()
    return s

def to_numeric_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    )

def nice_col(col: str) -> str:
    col = str(col)
    if len(col) <= 60:
        return col
    return col[:57] + "…"

# =========================
# Interface
# =========================
st.title("Evidência — Visualização Web da Avaliação")
st.caption(
    "Visualização agregada e anonimizada de dados exportados do Google Forms."
)

with st.sidebar:
    st.header("Entrada de dados")
    up = st.file_uploader(
        "Arquivo XLSX (exportação do Forms)",
        type=["xlsx"]
    )

    st.divider()
    st.header("Privacidade")
    hide_identifying = st.toggle(
        "Ocultar colunas identificáveis (nome/e-mails)",
        value=True
    )
    show_raw_table = st.toggle(
        "Mostrar tabela (dados não sensíveis)",
        value=False
    )

    st.divider()
    st.header("Comentários")
    max_comments = st.slider(
        "Máximo de comentários exibidos",
        5, 100, 30, 5
    )

if not up:
    st.info("Faça o upload do arquivo XLSX para iniciar a visualização.")
    st.stop()

# =========================
# Leitura dos dados
# =========================
df = pd.read_excel(up)

if df.empty:
    st.warning("O arquivo foi carregado, mas não contém dados.")
    st.stop()

cols = list(df.columns)
ts_col = guess_timestamp_col(cols)

df_work = df.copy()

# Remove colunas identificáveis
id_cols = [c for c in cols if is_email_col(c) or is_name_col(c)]
if hide_identifying and id_cols:
    df_work = df_work.drop(columns=id_cols, errors="ignore")

# =========================
# Resumo geral
# =========================
st.subheader("Resumo Geral")

c1, c2, c3 = st.columns(3)

c1.metric("Respondentes", len(df))

if ts_col:
    ts = pd.to_datetime(df[ts_col], errors="coerce")
    if ts.notna().any():
        c2.metric(
            "Período",
            f"{ts.min().date()} → {ts.max().date()}"
        )
    else:
        c2.metric("Período", "—")
else:
    c2.metric("Período", "—")

score_col = next(
    (c for c in df.columns if "pontuação" in str(c).lower()),
    None
)

if score_col:
    mean_score = to_numeric_series(df[score_col]).mean()
    c3.metric(
        "Média de Pontuação",
        f"{mean_score:.2f}" if pd.notna(mean_score) else "—"
    )
else:
    c3.metric("Média de Pontuação", "—")

# =========================
# Indicadores numéricos
# =========================
st.subheader("Indicadores de Avaliação (Médias)")

numeric_cols = []
for c in df_work.columns:
    s = to_numeric_series(df_work[c])
    if s.notna().mean() >= 0.5 and len(s.dropna().unique()) <= 20:
        numeric_cols.append(c)

if numeric_cols:
    means = pd.DataFrame({
        "Indicador": [nice_col(c) for c in numeric_cols],
        "Média": [to_numeric_series(df_work[c]).mean() for c in numeric_cols]
    }).sort_values("Média", ascending=False)

    chart = (
        alt.Chart(means)
        .mark_bar()
        .encode(
            x=alt.X("Média:Q", title="Média"),
            y=alt.Y("Indicador:N", sort="-x", title=""),
            tooltip=["Indicador:N", "Média:Q"]
        )
        .properties(height=25 * len(means) + 100)
    )

    st.altair_chart(chart, use_container_width=True)
    st.dataframe(means, use_container_width=True, hide_index=True)
else:
    st.info("Não foram encontrados indicadores numéricos suficientes.")

# =========================
# Questões conceituais
# =========================
st.subheader("Questões Conceituais (Frequência)")

textual_cols = [
    c for c in df.columns
    if c not in numeric_cols
    and c != ts_col
    and not is_comment_col(c)
    and not (hide_identifying and (is_email_col(c) or is_name_col(c)))
]

if textual_cols:
    q = st.selectbox(
        "Selecione a questão",
        textual_cols,
        format_func=nice_col
    )

    freq = (
        df[q]
        .dropna()
        .astype(str)
        .map(clean_text)
        .value_counts()
        .reset_index()
    )
    freq.columns = ["Resposta", "Quantidade"]

    chart2 = (
        alt.Chart(freq.head(15))
        .mark_bar()
        .encode(
            x="Quantidade:Q",
            y=alt.Y("Resposta:N", sort="-x"),
            tooltip=["Resposta", "Quantidade"]
        )
    )

    st.altair_chart(chart2, use_container_width=True)
    st.dataframe(freq, use_container_width=True, hide_index=True)
else:
    st.info("Nenhuma questão textual identificada.")

# =========================
# Comentários
# =========================
st.subheader("Comentários Qualitativos")

comment_cols = [c for c in df.columns if is_comment_col(c)]

if comment_cols:
    for c in comment_cols:
        st.markdown(f"**{nice_col(c)}**")
        comments = (
            df[c]
            .dropna()
            .astype(str)
            .map(clean_text)
            .tolist()
        )

        if not comments:
            st.write("_Sem comentários._")
        else:
            for txt in comments[:max_comments]:
                st.write(f"- {txt}")
        st.divider()
else:
    st.info("Nenhum campo de comentários encontrado.")

# =========================
# Tabela opcional
# =========================
if show_raw_table:
    st.subheader("Tabela de Dados (não sensíveis)")
    st.dataframe(df_work, use_container_width=True)

st.caption(
    "Visualização gerada automaticamente a partir do XLSX exportado do formulário."
)
