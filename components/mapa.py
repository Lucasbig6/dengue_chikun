import streamlit as st
import pandas as pd
import json
import plotly.express as px
from pathlib import Path

@st.cache_data
def carregar_geojson():
    geo_path = Path(__file__).resolve().parent.parent / "data" / "geojs-22-mun-ajustado.json"
    with open(geo_path, "r", encoding="utf-8") as f:
        geojson = json.load(f)
    return geojson

df_macro = pd.read_csv("data/macrorregiões.csv", encoding='utf-8-sig')

def render_mapa(df: pd.DataFrame, df_macro: pd.DataFrame):
    st.subheader("🗺️ Mapa Interativo de Casos de Dengue")

    geojson = carregar_geojson()

    # Padroniza coluna de município
    df["co_municipio_residencia"] = (
        df["co_municipio_residencia"]
        .astype(str)
        .str.replace(r"\.0+$", "", regex=True)
        .str.zfill(6)
    )

    # --- Junta dados do CSV de macrorregiões ---
    df_macro_mun = df_macro[["Codigo Municipio", "Macrorregiao de Saude", "Regiao de Saude"]].copy()
    df_macro_mun["Codigo Municipio"] = df_macro_mun["Codigo Municipio"].astype(str).str.zfill(6)
    df = df.merge(df_macro_mun, left_on="co_municipio_residencia", right_on="Codigo Municipio", how="left")

    # --- Filtros ---
    col1, col2 = st.columns(2)
    with col1.container(border=True):
        macro_opcoes = ["Todos"] + sorted(df["Macrorregiao de Saude"].dropna().unique())
        macro_sel = st.selectbox("📍 Selecione a Macrorregião", macro_opcoes)
    with col2.container(border=True):
        if macro_sel != "Todos":
            regioes = ["Todos"] + sorted(df.loc[df["Macrorregiao de Saude"] == macro_sel, "Regiao de Saude"].dropna().unique())
        else:
            regioes = ["Todos"] + sorted(df["Regiao de Saude"].dropna().unique())
        regiao_sel = st.selectbox("📍 Selecione a Região de Saúde", regioes)

    # Aplica filtros
    df_filtrado = df.copy()
    if macro_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Macrorregiao de Saude"] == macro_sel]
    if regiao_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Regiao de Saude"] == regiao_sel]

    # --- Tabela resumo ---
    tabela_resumo = df_filtrado.groupby(["Macrorregiao de Saude", "Regiao de Saude"]).agg(
        Casos_Provaveis=pd.NamedAgg(column="tp_classificacao_final", aggfunc=lambda x: x.isin([10,11,12]).sum()),
        Casos_Confirmados=pd.NamedAgg(column="tp_classificacao_final", aggfunc=lambda x: (x==10).sum())
    ).reset_index()

    # --- Layout mapa + tabela ---
    col_map, col_table = st.columns(2)

    with col_map:
        tab1, tab2 = st.tabs(["🟥 Casos Prováveis", "🟦 Casos Confirmados"])

        with tab1.container(border=True):
            df_provaveis = df_filtrado[df_filtrado["tp_classificacao_final"].isin([10,11,12])]
            casos = df_provaveis.groupby("co_municipio_residencia").size().reset_index(name="Casos Prováveis")
            if "no_municipio_residencia" in df_filtrado.columns:
                casos = casos.merge(
                    df_filtrado[["co_municipio_residencia", "no_municipio_residencia"]].drop_duplicates(),
                    on="co_municipio_residencia",
                    how="left"
                )
            fig = px.choropleth_mapbox(
                casos,
                geojson=geojson,
                locations="co_municipio_residencia",
                featureidkey="properties.id6",
                color="Casos Prováveis",
                color_continuous_scale="Reds",
                hover_name="no_municipio_residencia" if "no_municipio_residencia" in casos.columns else "co_municipio_residencia",
                hover_data={"co_municipio_residencia": False, "Casos Prováveis": True},
                mapbox_style="carto-positron",
                zoom=5.0,
                center={"lat": -7.5, "lon": -42.5},
                opacity=0.8,
                title="Casos Prováveis de Dengue por Município - Piauí"
            )
            fig.update_layout(margin={"r":0,"t":50,"l":0,"b":0})
            st.plotly_chart(fig, use_container_width=True)

        with tab2:
            df_conf = df_filtrado[df_filtrado["tp_classificacao_final"]==10]
            casos_conf = df_conf.groupby("co_municipio_residencia").size().reset_index(name="Casos Confirmados")
            if "no_municipio_residencia" in df_filtrado.columns:
                casos_conf = casos_conf.merge(
                    df_filtrado[["co_municipio_residencia", "no_municipio_residencia"]].drop_duplicates(),
                    on="co_municipio_residencia",
                    how="left"
                )
            fig2 = px.choropleth_mapbox(
                casos_conf,
                geojson=geojson,
                locations="co_municipio_residencia",
                featureidkey="properties.id6",
                color="Casos Confirmados",
                color_continuous_scale="Blues",
                hover_name="no_municipio_residencia" if "no_municipio_residencia" in casos_conf.columns else "co_municipio_residencia",
                hover_data={"co_municipio_residencia": False, "Casos Confirmados": True},
                mapbox_style="carto-positron",
                zoom=5.0,
                center={"lat": -7.5, "lon": -42.5},
                opacity=0.8,
                title="Casos Confirmados de Dengue por Município - Piauí"
            )
            fig2.update_layout(margin={"r":0,"t":50,"l":0,"b":0})
            st.plotly_chart(fig2, use_container_width=True)

    with col_table.container(border=True):
        st.write("📊 Resumo de Casos por Macrorregião e Região de Saúde")
        st.divider()
        st.dataframe(tabela_resumo, use_container_width=True)
