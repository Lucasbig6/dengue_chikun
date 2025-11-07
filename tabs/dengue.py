import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path
import json
import datetime
from components.mapa import render_mapa

@st.cache_data
def carregar_dados():
    return pd.read_csv("data/dengue.csv", encoding='utf-8-sig', low_memory=False)

@st.cache_data
def carregar_geojson():
    geo_path = Path(__file__).resolve().parent.parent / "data" / "geojs-22-mun-ajustado.json"
    with open(geo_path, "r", encoding="utf-8") as f:
        geojson = json.load(f)
    return geojson

def render_dengue():
    st.subheader("Dados Dengue", divider=True)
    st.write("Filtros")

    df = carregar_dados()
    geojson = carregar_geojson()
    df_macro = pd.read_csv("data/macrorregiões.csv", encoding='utf-8-sig')

    # Extrai nomes e códigos dos municípios
    municipios = [{"nome": f["properties"]["name"], "codigo": f["properties"]["id6"]} for f in geojson["features"]]
    municipios_df = pd.DataFrame(municipios).sort_values("nome")

    if "dt_notificacao" in df.columns:
        df["dt_notificacao"] = pd.to_datetime(df["dt_notificacao"], errors="coerce")
        df["Ano"] = df["dt_notificacao"].dt.year
    else:
        st.warning("⚠️ Coluna 'dt_notificacao' não encontrada no dataset.")
        return

    df["co_municipio_residencia"] = df["co_municipio_residencia"].astype(str).str.replace(r"\.0+$", "", regex=True).str.zfill(6)

    # --- Filtros ---
    col1, col2  = st.columns(2)

    with col1.container(border=True):
        anos = sorted(df["Ano"].dropna().unique(), reverse=True)
        ano_sel = st.selectbox("📅 Selecione o Ano", anos)

    with col2.container(border=True):
        municipios_opcoes = ["Todos"] + municipios_df["nome"].tolist()
        municipio_sel = st.selectbox("🏙️ Selecione o Município", municipios_opcoes, index=0)

    # --- Aplica filtros ---
    df_filtrado = df[df["Ano"] == ano_sel]

    if municipio_sel != "Todos":
        cod_mun = municipios_df.loc[municipios_df["nome"] == municipio_sel, "codigo"].iloc[0]
        df_filtrado = df_filtrado[df_filtrado["co_municipio_residencia"] == cod_mun]

    # Aplica os filtros
    df_filtrado = df[df["Ano"] == ano_sel]
    if municipio_sel != "Todos":
        cod_mun = municipios_df.loc[municipios_df["nome"] == municipio_sel, "codigo"].iloc[0]
        df_filtrado = df_filtrado[df_filtrado["co_municipio_residencia"] == cod_mun]

    # Cálculos principais
    df_provaveis = df_filtrado[df_filtrado["tp_classificacao_final"].isin([10, 11, 12])]
    qtd_provaveis = len(df_provaveis)

    obitos_invest = df_filtrado[df_filtrado["tp_evolucao_caso"] == 4]
    qtd_invest = len(obitos_invest)

    obitos_dengue = df_filtrado[df_filtrado["tp_evolucao_caso"] == 2]
    qtd_obitos = len(obitos_dengue)

    letalidade = (qtd_obitos / qtd_provaveis) * 100 if qtd_provaveis > 0 else 0

    casos_graves = df_filtrado[df_filtrado["tp_classificacao_final"] == 12]
    obitos_graves = casos_graves[casos_graves["tp_evolucao_caso"] == 4]
    qtd_graves = len(casos_graves)
    qtd_obitos_graves = len(obitos_graves)
    letalidade_graves = (qtd_obitos_graves / qtd_graves) * 100 if qtd_graves > 0 else 0

    populacao = 3447200
    ci = (qtd_provaveis / populacao) * 100000

    # Cards
    col1, col2, col3 = st.columns(3)
    col1.metric("Casos Prováveis", f"{qtd_provaveis:,}".replace(",", "."), border=True)
    col1.metric("Óbitos em Investigação", f"{qtd_invest:,}".replace(",", "."), border=True)
    col2.metric("Óbitos por Dengue", f"{qtd_obitos:,}".replace(",", "."), border=True)
    col2.metric("Letalidade Casos Prováveis (%)", f"{letalidade:.2f}%", border=True)
    col3.metric("Letalidade Casos Graves (%)", f"{letalidade_graves:.2f}%", border=True)
    col3.metric("Coeficiente de Incidência", f"{ci:.2f}", border=True)

    st.divider()

    render_mapa(df_filtrado, df_macro)
    

    st.divider()

    # --- Casos prováveis por ano ---
    col1, col2 = st.columns(2)
    df_provaveis_grafico = df[df["tp_classificacao_final"].isin([10, 11, 12])]
    casos_por_ano = df_provaveis_grafico.groupby("Ano").size().reset_index(name="Casos")

    chart_casos = (
        alt.Chart(casos_por_ano)
        .mark_bar(color="#f87171")
        .encode(x="Ano:O", y="Casos:Q", tooltip=["Ano", "Casos"])
        .properties(width=700, height=400, title="Casos Prováveis de Dengue (2020 a 2025)")
    )
    text_casos = chart_casos.mark_text(align='center', baseline='bottom', dy=-2, color='black').encode(text='Casos:Q')
    chart_casos_final = chart_casos + text_casos
    
    with col1.container(border=True):
        st.altair_chart(chart_casos_final, use_container_width=True)

    # --- Óbitos por ano ---
    df_obitos = df[df["tp_evolucao_caso"] == 2]
    obitos_por_ano = df_obitos.groupby("Ano").size().reset_index(name="Óbitos")
    chart_obitos = (
        alt.Chart(obitos_por_ano)
        .mark_bar(color="#60a5fa")
        .encode(x="Ano:O", y="Óbitos:Q", tooltip=["Ano", "Óbitos"])
        .properties(height=400, title="Óbitos por Dengue (2020-2025)")
    )
    text_obitos = chart_obitos.mark_text(align='center', baseline='bottom', dy=-2, color='black').encode(text='Óbitos:Q')
    chart_obitos_final = chart_obitos + text_obitos
    
    with col2.container(border=True):
        st.altair_chart(chart_obitos_final, use_container_width=True)

    # --- Casos por sexo (pizza) ---
    col1, col2, col3 = st.columns(3)
    casos_por_sexo = df_provaveis.groupby("tp_sexo").size().reset_index(name="Casos")
    map_sexo = {"F": "Feminino", "M": "Masculino", "I": "Indefinido"}
    casos_por_sexo["tp_sexo"] = casos_por_sexo["tp_sexo"].map(map_sexo).fillna("Indefinido")
    total = casos_por_sexo["Casos"].sum()
    casos_por_sexo["Porcentagem"] = (casos_por_sexo["Casos"] / total * 100).round(1) if total > 0 else 0

    chart_roda = alt.Chart(casos_por_sexo).mark_arc(innerRadius=60, outerRadius=100).encode(
        theta="Porcentagem:Q",
        color=alt.Color("tp_sexo:N", title="Sexo"),
        tooltip=["tp_sexo:N", "Casos:Q", "Porcentagem:Q"]
    )
    text_roda = chart_roda.mark_text(radius=120, size=14, color='black').encode(text=alt.Text("Porcentagem:Q", format=".1f"))
    chart_roda_final = chart_roda + text_roda
    chart_roda_final = chart_roda_final.properties(title=f"Distribuição de Casos Prováveis por Sexo — {ano_sel}", width=400, height=400)
    
    with col1.container(border=True):
        st.altair_chart(chart_roda_final, use_container_width=True)

    # --- Casos por raça/cor ---
    mapa_raca_cor = {1: "Branca", 2: "Preta", 3: "Parda", 4: "Amarela", 5: "Indígena", 9: "Ignorado"}
    df_provaveis["Raça/Cor"] = df_provaveis["tp_raca_cor"].map(mapa_raca_cor).fillna("Ignorado")
    casos_por_raca = df_provaveis.groupby("Raça/Cor").size().reset_index(name="Casos Prováveis").sort_values(by="Casos Prováveis", ascending=False)
    chart_raca = alt.Chart(casos_por_raca).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
        x=alt.X("Casos Prováveis:Q", title="Quantidade de Casos"),
        y=alt.Y("Raça/Cor:N", sort="-x", title="Raça/Cor"),
        color=alt.Color("Raça/Cor:N", legend=None, scale=alt.Scale(scheme="set2")),
        tooltip=["Raça/Cor:N", "Casos Prováveis:Q"]
    ).properties(width="container", height=400, title="Casos Prováveis por Raça/Cor")
    text_raca = chart_raca.mark_text(align='left', baseline='middle', dx=3, color='black', fontWeight='bold').encode(text='Casos Prováveis:Q')
    chart_raca_final = chart_raca + text_raca
    
    with col2.container(border=True):
        st.altair_chart(chart_raca_final, use_container_width=True)

    # --- Casos por faixa etária e sexo ---
    df_provaveis["dt_nascimento"] = pd.to_datetime(df_provaveis["dt_nascimento"], errors="coerce")
    data_ref = df_provaveis["dt_notificacao"].fillna(datetime.datetime.now())
    df_provaveis["idade"] = ((data_ref - df_provaveis["dt_nascimento"]).dt.days // 365.25).astype('Int64')

    bins = [0, 1, 5, 10, 15, 20, 30, 40, 50, 60, 70, 80, float('inf')]
    labels = ["<1 ano", "01 a 04 anos", "05 a 09 anos", "10 a 14 anos", "15 a 19 anos", "20 a 29 anos",
              "30 a 39 anos", "40 a 49 anos", "50 a 59 anos", "60 a 69 anos", "70 a 79 anos", "80 anos e mais"]
    df_provaveis["faixa_etaria"] = pd.cut(df_provaveis["idade"], bins=bins, labels=labels, right=False, include_lowest=True)
    df_provaveis["sexo_label"] = df_provaveis["tp_sexo"].map(map_sexo).fillna("Indefinido")

    casos_faixa_sexo = df_provaveis.groupby(["faixa_etaria", "sexo_label"], observed=True).size().reset_index(name="Casos")
    chart_faixa_sexo = alt.Chart(casos_faixa_sexo).mark_bar(height=15).encode(
        y=alt.Y("faixa_etaria:N", title="Faixa Etária", sort=labels, axis=alt.Axis(labelFontSize=11)),
        x=alt.X("Casos:Q", title="Número de Casos"),
        color=alt.Color("sexo_label:N", title="Sexo", scale=alt.Scale(domain=["Masculino", "Feminino"], range=["#60a5fa", "#f87171"])),
        tooltip=["faixa_etaria:N", "sexo_label:N", alt.Tooltip("Casos:Q", format=",")]
    ).properties(title=f"Casos Prováveis por Faixa Etária e Sexo — {ano_sel}", width="container", height=400)
    text_faixa_sexo = chart_faixa_sexo.mark_text(align='left', baseline='middle', dx=3, color='black', fontWeight='bold').encode(text='Casos:Q')
    chart_faixa_sexo_final = chart_faixa_sexo + text_faixa_sexo
    
    with col3.container(border=True):
        st.altair_chart(chart_faixa_sexo_final, use_container_width=True)


    # --- Casos prováveis por semana epidemiológica ---
# Extrai apenas a semana (últimos dois dígitos da coluna ano+semana)
    df_provaveis["ds_semana_sintoma"] = df_provaveis["ds_semana_sintoma"].astype(str).str[-2:].astype(int)

    # Cria todas as semanas de 1 a 52
    semanas = pd.DataFrame({"ds_semana_sintoma": range(1, 53)})

    # Agrupa por semana
    casos_por_semana = df_provaveis.groupby("ds_semana_sintoma").size().reset_index(name="Casos")

    # Merge para incluir semanas sem casos
    casos_por_semana = semanas.merge(casos_por_semana, on="ds_semana_sintoma", how="left")
    casos_por_semana["Casos"] = casos_por_semana["Casos"].fillna(0).astype(int)

    # Ordena por semana
    casos_por_semana = casos_por_semana.sort_values("ds_semana_sintoma")

    # Gráfico de barras
    chart_semana = (
        alt.Chart(casos_por_semana)
        .mark_bar(color="#f87171")
        .encode(
            x=alt.X("ds_semana_sintoma:O", title="Semana Epidemiológica"),  # eixo categórico de 1 a 52
            y=alt.Y("Casos:Q", title="Número de Casos"),
            tooltip=["ds_semana_sintoma:N", "Casos:Q"]
        )
        .properties(
            width="container",
            height=400,
            title=f"Casos Prováveis por Semana Epidemiológica — {ano_sel}"
        )
    )

    # Adiciona rótulos no topo das barras
    text_semana = chart_semana.mark_text(
        align='center',
        baseline='bottom',
        dy=-2,
        color='black'
    ).encode(text='Casos:Q')

    chart_semana_final = chart_semana + text_semana

    with st.container(border=True):
        st.altair_chart(chart_semana_final, use_container_width=True)

    st.divider()

    st.subheader("📊 Exploração Dinâmica de Dados")

    # Copia dataframe filtrado do dashboard
    df_explorar = df_filtrado.copy()

    # Seleção de colunas para o usuário
    colunas_disponiveis = [
        "Ano", "co_municipio_residencia", "tp_classificacao_final", "tp_evolucao_caso", "tp_sexo",
        "ds_semana_sintoma"
    ]

    colunas_selecionadas = st.multiselect(
        "Selecione as colunas para tabular", 
        options=colunas_disponiveis, 
        default=colunas_disponiveis
    )

    # Cria layout lado a lado
    col1, col2 = st.columns(2)

    with col1:
        # Tabela interativa
        tabela_interativa = st.data_editor(
            df_explorar[colunas_selecionadas],
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True
        )

    with col2:
        if tabela_interativa is not None and not tabela_interativa.empty:
            # Exemplo de gráfico dinâmico: contagem por coluna selecionada
            # Para simplificar, usa a primeira coluna categórica escolhida
            coluna_grafico = colunas_selecionadas[0]

            df_plot = tabela_interativa.groupby(coluna_grafico).size().reset_index(name="Contagem")

            chart = alt.Chart(df_plot).mark_bar(color="#4f46e5").encode(
                x=alt.X(f"{coluna_grafico}:N", title=coluna_grafico),
                y=alt.Y("Contagem:Q", title="Número de Casos"),
                tooltip=[coluna_grafico, "Contagem"]
            ).properties(
                height=370, 
                title=f"Gráfico Dinâmico - {coluna_grafico}"
            )

            st.container(border=True).altair_chart(chart, use_container_width=True)

        else:
            st.info("🔹 Selecione/filtre dados na tabela para gerar o gráfico.")