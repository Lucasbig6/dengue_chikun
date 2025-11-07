import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path
import json
import datetime
from components.mapa import render_mapa

# --- Cache de dados ---
@st.cache_data
def carregar_dados():
    return pd.read_csv("data/chikungunya.csv", encoding='utf-8-sig', low_memory=False)

@st.cache_data
def carregar_geojson():
    geo_path = Path(__file__).resolve().parent.parent / "data" / "geojs-22-mun-ajustado.json"
    with open(geo_path, "r", encoding="utf-8") as f:
        geojson = json.load(f)
    return geojson

def render_chikun():
    st.subheader("Dados Chikungunya", divider=True)
    st.write("Filtros")

    df = carregar_dados()
    geojson = carregar_geojson()
    df_macro = pd.read_csv("data/macrorregiões.csv", encoding='utf-8-sig')

    # Extrai nomes e códigos dos municípios
    municipios = [{"nome": f["properties"]["name"], "codigo": f["properties"]["id6"]} for f in geojson["features"]]
    municipios_df = pd.DataFrame(municipios).sort_values("nome")

    # Conversão de datas
    if "dt_notificacao" in df.columns:
        df["dt_notificacao"] = pd.to_datetime(df["dt_notificacao"], errors="coerce")
        df["Ano"] = df["dt_notificacao"].dt.year
    else:
        st.warning("⚠️ Coluna 'dt_notificacao' não encontrada no dataset.")
        return

    # Padroniza código do município
    df["co_municipio_residencia"] = df["co_municipio_residencia"].astype(str).str.replace(r"\.0+$", "", regex=True).str.zfill(6)

    # --- Filtros ---
    col1, col2 = st.columns(2)

    with col1.container(border=True):
        anos = sorted(df["Ano"].dropna().unique(), reverse=True)
        ano_sel = st.selectbox("📅 Selecione o Ano", anos, key="ano_chikun")

    with col2.container(border=True):
        municipios_opcoes = ["Todos"] + municipios_df["nome"].tolist()
        municipio_sel = st.selectbox("🏙️ Selecione o Município", municipios_opcoes, index=0, key="mun_chikun")

    # Aplica filtros
    df_filtrado = df[df["Ano"] == ano_sel]
    if municipio_sel != "Todos":
        cod_mun = municipios_df.loc[municipios_df["nome"] == municipio_sel, "codigo"].iloc[0]
        df_filtrado = df_filtrado[df_filtrado["co_municipio_residencia"] == cod_mun]

    # --- Definição de casos ---
    casos_confirmados_val = [10, 11, 12]
    casos_provaveis_val = [1, 2, 3, 4, 5, 8]

    df_confirmados = df_filtrado[df_filtrado["tp_classificacao_final"].isin(casos_confirmados_val)]
    df_provaveis = df_filtrado[df_filtrado["tp_classificacao_final"].isin(casos_provaveis_val)]

    qtd_provaveis = len(df_provaveis)
    qtd_confirmados = len(df_confirmados)
    obitos_invest = df_filtrado[df_filtrado["tp_evolucao_caso"] == 4]
    qtd_invest = len(obitos_invest)
    obitos_chikun = df_filtrado[df_filtrado["tp_evolucao_caso"] == 2]
    qtd_obitos = len(obitos_chikun)

    letalidade = (qtd_obitos / qtd_provaveis) * 100 if qtd_provaveis > 0 else 0

    casos_graves = df_provaveis[df_provaveis["tp_classificacao_final"] == 8]  # exemplo se houver casos graves
    obitos_graves = casos_graves[casos_graves["tp_evolucao_caso"] == 4]
    qtd_graves = len(casos_graves)
    qtd_obitos_graves = len(obitos_graves)
    letalidade_graves = (qtd_obitos_graves / qtd_graves) * 100 if qtd_graves > 0 else 0

    populacao = 3447200
    ci = (qtd_provaveis / populacao) * 100000

    # --- Cards ---
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Casos Prováveis", f"{qtd_provaveis:,}".replace(",", "."), border=True)
    col2.metric("Óbitos em Investigação", f"{qtd_invest:,}".replace(",", "."), border=True)
    col3.metric("Óbitos por Chikungunya", f"{qtd_obitos:,}".replace(",", "."), border=True)
    col4.metric("Coeficiente de Incidência", f"{ci:.2f}", border=True)

    st.divider()

    # --- Gráficos ---
    # Casos confirmados por ano
    casos_por_ano = df_confirmados.groupby("Ano").size().reset_index(name="Casos")
    chart_casos = alt.Chart(casos_por_ano).mark_bar(color="#4f46e5").encode(
        x="Ano:O",
        y="Casos:Q",
        tooltip=["Ano", "Casos"]
    ).properties(height=400, title="Casos Confirmados por Ano")
    st.altair_chart(chart_casos, use_container_width=True)

    # Casos prováveis por sexo
    map_sexo = {"F": "Feminino", "M": "Masculino", "I": "Indefinido"}
    df_provaveis["tp_sexo"] = df_provaveis["tp_sexo"].map(map_sexo).fillna("Indefinido")
    casos_por_sexo = df_provaveis.groupby("tp_sexo").size().reset_index(name="Casos")
    chart_sexo = alt.Chart(casos_por_sexo).mark_bar().encode(
        x="tp_sexo:N",
        y="Casos:Q",
        tooltip=["tp_sexo", "Casos"]
    ).properties(height=400, title="Casos Prováveis por Sexo")
    st.altair_chart(chart_sexo, use_container_width=True)

    # Exploração dinâmica de dados
    st.subheader("📊 Exploração Dinâmica de Dados")
    colunas_disponiveis = ["Ano", "co_municipio_residencia", "tp_classificacao_final", "tp_evolucao_caso", "tp_sexo", "ds_semana_sintoma"]
    colunas_selecionadas = st.multiselect(
        "Selecione as colunas para tabular", 
        options=colunas_disponiveis, 
        default=colunas_disponiveis
    )

    col1, col2 = st.columns(2)
    with col1:
        tabela_interativa = st.data_editor(
            df_filtrado[colunas_selecionadas],
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True
        )
    with col2:
        if tabela_interativa is not None and not tabela_interativa.empty:
            coluna_grafico = colunas_selecionadas[0]
            df_plot = tabela_interativa.groupby(coluna_grafico).size().reset_index(name="Contagem")
            chart = alt.Chart(df_plot).mark_bar(color="#4f46e5").encode(
                x=alt.X(f"{coluna_grafico}:N", title=coluna_grafico),
                y=alt.Y("Contagem:Q", title="Número de Casos"),
                tooltip=[coluna_grafico, "Contagem"]
            ).properties(height=370, title=f"Gráfico Dinâmico - {coluna_grafico}")
            st.container(border=True).altair_chart(chart, use_container_width=True)
        else:
            st.info("🔹 Selecione/filtre dados na tabela para gerar o gráfico.")
