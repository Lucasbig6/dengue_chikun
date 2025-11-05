import streamlit as st
from tabs.dengue import render_dengue 
from tabs.chikun import render_chikun
import pandas as pd

# Configuração da página
st.set_page_config(page_title="Painel Arboviroses", layout="wide", initial_sidebar_state="expanded", page_icon="🦟")

st.title("🦟 Painel Arboviroses")

tabs1, tabs2 = st.tabs(["Dengue", "Chikungunya"])

with tabs1:
    render_dengue()

with tabs2:
    render_chikun()    
