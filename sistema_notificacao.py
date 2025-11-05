import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import pandas as pd
from dbfread import DBF
import os
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.cm import ScalarMappable
from matplotlib.colors import Normalize
import seaborn as sns
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import PageTemplate, Frame
from reportlab.lib.units import inch
import io
from PIL import Image as PILImage
import geopandas as gpd
import geobr
import unicodedata
import statistics
import logging
import numpy as np

# Configurar logging para debug
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Dicionário completo com os 224 municípios do Piauí
MUNICIPIOS_PIAUI = {
    '220005': 'Acauã', '220010': 'Agricolândia', '220020': 'Água Branca', '220025': 'Alagoinha do Piauí',
    '220027': 'Alegrete do Piauí', '220030': 'Alto Longá', '220040': 'Altos', '220045': 'Alvorada do Gurguéia',
    '220050': 'Amarante', '220060': 'Angical do Piauí', '220070': 'Anísio de Abreu', '220080': 'Antônio Almeida',
    '220090': 'Aroazes', '220095': 'Aroeiras do Itaim', '220100': 'Arraial', '220105': 'Assunção do Piauí',
    '220110': 'Avelino Lopes', '220115': 'Baixa Grande do Ribeiro', '220117': 'Barra D\'Alcântara', '220120': 'Barras',
    '220130': 'Barreiras do Piauí', '220140': 'Barro Duro', '220150': 'Batalha', '220155': 'Bela Vista do Piauí',
    '220157': 'Belém do Piauí', '220160': 'Beneditinos', '220170': 'Bertolínia', '220173': 'Betânia do Piauí',
    '220177': 'Boa Hora', '220180': 'Bocaina', '220190': 'Bom Jesus', '220191': 'Bom Princípio do Piauí',
    '220192': 'Bonfim do Piauí', '220194': 'Boqueirão do Piauí', '220196': 'Brasileira', '220198': 'Brejo do Piauí',
    '220200': 'Buriti dos Lopes', '220202': 'Buriti dos Montes', '220205': 'Cabeceiras do Piauí', '220207': 'Cajazeiras do Piauí',
    '220208': 'Cajueiro da Praia', '220209': 'Caldeirão Grande do Piauí', '220210': 'Campinas do Piauí', '220211': 'Campo Alegre do Fidalgo',
    '220213': 'Campo Grande do Piauí', '220217': 'Campo Largo do Piauí', '220220': 'Campo Maior', '220225': 'Canavieira',
    '220230': 'Canto do Buriti', '220240': 'Capitão de Campos', '220245': 'Capitão Gervásio Oliveira', '220250': 'Caracol',
    '220253': 'Caraúbas do Piauí', '220255': 'Caridade do Piauí', '220260': 'Castelo do Piauí', '220265': 'Caxingó',
    '220270': 'Cocal', '220271': 'Cocal de Telha', '220272': 'Cocal dos Alves', '220273': 'Coivaras',
    '220275': 'Colônia do Gurguéia', '220277': 'Colônia do Piauí', '220280': 'Conceição do Canindé', '220285': 'Coronel José Dias',
    '220290': 'Corrente', '220300': 'Cristalândia do Piauí', '220310': 'Cristino Castro', '220320': 'Curimatá',
    '220323': 'Currais', '220325': 'Curral Novo do Piauí', '220327': 'Curralinhos', '220330': 'Demerval Lobão',
    '220335': 'Dirceu Arcoverde', '220340': 'Dom Expedito Lopes', '220342': 'Domingos Mourão', '220345': 'Dom Inocêncio',
    '220350': 'Elesbão Veloso', '220360': 'Eliseu Martins', '220370': 'Esperantina', '220375': 'Fartura do Piauí',
    '220380': 'Flores do Piauí', '220385': 'Floresta do Piauí', '220390': 'Floriano', '220400': 'Francinópolis',
    '220410': 'Francisco Ayres', '220415': 'Francisco Macedo', '220420': 'Francisco Santos', '220430': 'Fronteiras',
    '220435': 'Geminiano', '220440': 'Gilbués', '220450': 'Guadalupe', '220455': 'Guaribas',
    '220460': 'Hugo Napoleão', '220465': 'Ilha Grande', '220470': 'Inhuma', '220480': 'Ipiranga do Piauí',
    '220490': 'Isaías Coelho', '220500': 'Itainópolis', '220510': 'Itaueira', '220515': 'Jacobina do Piauí',
    '220520': 'Jaicós', '220525': 'Jardim do Mulato', '220527': 'Jatobá do Piauí', '220530': 'Jerumenha',
    '220535': 'João Costa', '220540': 'Joaquim Pires', '220545': 'Joca Marques', '220550': 'José de Freitas',
    '220551': 'Juazeiro do Piauí', '220552': 'Júlio Borges', '220553': 'Jurema', '220554': 'Lagoinha do Piauí',
    '220555': 'Lagoa Alegre', '220556': 'Lagoa do Barro do Piauí', '220557': 'Lagoa de São Francisco', '220558': 'Lagoa do Piauí',
    '220559': 'Lagoa do Sítio', '220560': 'Landri Sales', '220570': 'Luís Correia', '220580': 'Luzilândia',
    '220585': 'Madeiro', '220590': 'Manoel Emídio', '220595': 'Marcolândia', '220600': 'Marcos Parente',
    '220605': 'Massapê do Piauí', '220610': 'Matias Olímpio', '220620': 'Miguel Alves', '220630': 'Miguel Leão',
    '220635': 'Milton Brandão', '220640': 'Monsenhor Gil', '220650': 'Monsenhor Hipólito', '220660': 'Monte Alegre do Piauí',
    '220665': 'Morro Cabeça no Tempo', '220667': 'Morro do Chapéu do Piauí', '220669': 'Murici dos Portelas',
    '220670': 'Nazaré do Piauí', '220672': 'Nazária', '220675': 'Nossa Senhora de Nazaré', '220680': 'Nossa Senhora dos Remédios',
    '220690': 'Novo Oriente do Piauí', '220695': 'Novo Santo Antônio', '220700': 'Oeiras', '220710': 'Olho D\'Água do Piauí',
    '220720': 'Padre Marcos', '220730': 'Paes Landim', '220735': 'Pajeú do Piauí', '220740': 'Palmeira do Piauí',
    '220750': 'Palmeirais', '220755': 'Paquetá', '220760': 'Parnaguá', '220770': 'Parnaíba', '220775': 'Passagem Franca do Piauí',
    '220777': 'Patos do Piauí', '220779': 'Pau D\'Arco do Piauí', '220780': 'Paulistana', '220785': 'Pavussu',
    '220790': 'Pedro II', '220793': 'Pedro Laurentino', '220795': 'Nova Santa Rita', '220800': 'Picos',
    '220810': 'Pimenteiras', '220820': 'Pio IX', '220830': 'Piracuruca', '220840': 'Piripiri',
    '220850': 'Porto', '220855': 'Porto Alegre do Piauí', '220860': 'Prata do Piauí', '220865': 'Queimada Nova',
    '220870': 'Redenção do Gurguéia', '220880': 'Regeneração', '220885': 'Riacho Frio', '220887': 'Ribeira do Piauí',
    '220890': 'Ribeiro Gonçalves', '220900': 'Rio Grande do Piauí', '220910': 'Santa Cruz do Piauí', '220915': 'Santa Cruz dos Milagres',
    '220920': 'Santa Filomena', '220930': 'Santa Luz', '220935': 'Santana do Piauí', '220937': 'Santa Rosa do Piauí',
    '220940': 'Santo Antônio de Lisboa', '220945': 'Santo Antônio dos Milagres', '220950': 'Santo Inácio do Piauí',
    '220955': 'São Braz do Piauí', '220960': 'São Félix do Piauí', '220965': 'São Francisco de Assis do Piauí', '220970': 'São Francisco do Piauí',
    '220975': 'São Gonçalo do Gurguéia', '220980': 'São Gonçalo do Piauí', '220985': 'São João da Canabrava', '220987': 'São João da Fronteira',
    '220990': 'São João da Serra', '220995': 'São João da Varjota', '220997': 'São João do Arraial', '221000': 'São João do Piauí',
    '221005': 'São José do Divino', '221010': 'São José do Peixe', '221020': 'São José do Piauí', '221030': 'São Julião',
    '221035': 'São Lourenço do Piauí', '221037': 'São Luis do Piauí', '221038': 'São Miguel da Baixa Grande', '221039': 'São Miguel do Fidalgo',
    '221040': 'São Miguel do Tapuio', '221050': 'São Pedro do Piauí', '221060': 'São Raimundo Nonato', '221062': 'Sebastião Barros',
    '221063': 'Sebastião Leal', '221065': 'Sigefredo Pacheco', '221070': 'Simões', '221080': 'Simplício Mendes',
    '221090': 'Socorro do Piauí', '221093': 'Sussuapara', '221095': 'Tamboril do Piauí', '221097': 'Tanque do Piauí',
    '221100': 'Teresina', '221110': 'União', '221120': 'Uruçuí', '221130': 'Valença do Piauí',
    '221135': 'Várzea Branca', '221140': 'Várzea Grande', '221150': 'Vera Mendes', '221160': 'Vila Nova do Piauí',
    '221170': 'Wall Ferraz'
}

# Dicionário de Territórios de Saúde do Piauí
territorios_municipios = {
    "Planície Litorânea": ["Bom Princípio do Piauí", "Buriti dos Lopes", "Cajueiro da Praia", "Caraúbas do Piauí", "Caxingó", "Cocal", "Cocal dos Alves", "Ilha Grande", "Luís Correia", "Murici dos Portelas", "Parnaíba"],
    "Cocais": ["Barras", "Batalha", "Brasileira", "Campo Largo do Piauí", "Capitão de Campos", "Domingos Mourão", "Esperantina", "Joaquim Pires", "Joca Marques", "Lagoa de São Francisco", "Luzilândia", "Madeiro", "Matias Olímpio", "Milton Brandão", "Morro do Chapéu do Piauí", "Nossa Senhora dos Remédios", "Pedro II", "Piracuruca", "Piripiri", "Porto", "São João da Fronteira", "São João do Arraial", "São José do Divino"],
    "Carnaubais": ["Assunção do Piauí", "Boa Hora", "Boqueirão do Piauí", "Buriti dos Montes", "Cabeceiras do Piauí", "Campo Maior", "Castelo do Piauí", "Cocal de Telha", "Jatobá do Piauí", "Juazeiro do Piauí", "Nossa Senhora de Nazaré", "Novo Santo Antônio", "São João da Serra", "São Miguel do Tapuio", "Sigefredo Pacheco"],
    "Entre Rios": ["Agricolândia", "Água Branca", "Alto Longá", "Altos", "Amarante", "Angical do Piauí", "Barro Duro", "Beneditinos", "Coivaras", "Curralinhos", "Demerval Lobão", "Hugo Napoleão", "Jardim do Mulato", "José de Freitas", "Lagoa Alegre", "Lagoa do Piauí", "Lagoinha do Piauí", "Miguel Alves", "Miguel Leão", "Monsenhor Gil", "Nazária", "Olho D'Água do Piauí", "Palmeirais", "Passagem Franca do Piauí", "Pau D'Arco do Piauí", "Regeneração", "Santo Antônio dos Milagres", "São Gonçalo do Piauí", "São Pedro do Piauí", "Teresina", "União"],
    "Vale do Sambito": ["Aroazes", "Barra D'Alcântara", "Elesbão Veloso", "Francinópolis", "Inhuma", "Lagoa do Sítio", "Novo Oriente do Piauí", "Pimenteiras", "Prata do Piauí", "Santa Cruz dos Milagres", "São Félix do Piauí", "São Miguel da Baixa Grande", "Valença do Piauí", "Várzea Grande"],
    "Vale do Rio Guaribas": ["Alagoinha do Piauí", "Alegrete do Piauí", "Aroeiras do Itaim", "Bocaina", "Campo Grande do Piauí", "Dom Expedito Lopes", "Francisco Santos", "Fronteiras", "Geminiano", "Ipiranga do Piauí", "Itainópolis", "Monsenhor Hipólito", "Paquetá", "Picos", "Pio IX", "Santa Cruz do Piauí", "Santana do Piauí", "Santo Antônio de Lisboa", "São João da Canabrava", "São José do Piauí", "São Julião", "São Luis do Piauí", "Sussuapara", "Vera Mendes", "Vila Nova do Piauí", "Wall Ferraz"],
    "Vale do Canindé": ["Bela Vista do Piauí", "Cajazeiras do Piauí", "Campinas do Piauí", "Colônia do Piauí", "Conceição do Canindé", "Floresta do Piauí", "Isaías Coelho", "Oeiras", "Santa Rosa do Piauí", "Santo Inácio do Piauí", "São Francisco de Assis do Piauí", "São João da Varjota", "Simplício Mendes", "Tanque do Piauí"],
    "Serra da Capivara": ["Anísio de Abreu", "Bonfim do Piauí", "Campo Alegre do Fidalgo", "Capitão Gervásio Oliveira", "Caracol", "Coronel José Dias", "Dirceu Arcoverde", "Dom Inocêncio", "Fartura do Piauí", "Guaribas", "João Costa", "Jurema", "Lagoa do Barro do Piauí", "Nova Santa Rita", "Pedro Laurentino", "São Braz do Piauí", "São João do Piauí", "São Lourenço do Piauí", "São Raimundo Nonato", "Várzea Branca"],
    "Vale dos Rios Piauí e Itaueira": ["Arraial", "Bertolínia", "Brejo do Piauí", "Canavieira", "Canto do Buriti", "Colônia do Gurguéia", "Eliseu Martins", "Flores do Piauí", "Floriano", "Francisco Ayres", "Guadalupe", "Itaueira", "Jerumenha", "Landri Sales", "Manoel Emídio", "Marcos Parente", "Nazaré do Piauí", "Paes Landim", "Pajeú do Piauí", "Pavussu", "Porto Alegre do Piauí", "Ribeira do Piauí", "Rio Grande do Piauí", "São Francisco do Piauí", "São José do Peixe", "São Miguel do Fidalgo", "Socorro do Piauí", "Tamboril do Piauí"],
    "Tabuleiros do Alto Parnaíba": ["Antônio Almeida", "Baixa Grande do Ribeiro", "Ribeiro Gonçalves", "Sebastião Leal", "Uruçuí"],
    "Chapada das Mangabeiras": ["Alvorada do Gurguéia", "Avelino Lopes", "Barreiras do Piauí", "Bom Jesus", "Corrente", "Cristalândia do Piauí", "Cristino Castro", "Curimatá", "Currais", "Gilbués", "Júlio Borges", "Monte Alegre do Piauí", "Morro Cabeça no Tempo", "Palmeira do Piauí", "Parnaguá", "Redenção do Gurguéia", "Riacho Frio", "Santa Filomena", "Santa Luz", "São Gonçalo do Gurguéia", "Sebastião Barros"],
    "Chapada Vale do Itaim": ["Acauã", "Belém do Piauí", "Betânia do Piauí", "Caldeirão Grande do Piauí", "Caridade do Piauí", "Curral Novo do Piauí", "Francisco Macedo", "Jacobina do Piauí", "Jaicós", "Marcolândia", "Massapê do Piauí", "Padre Marcos", "Patos do Piauí", "Paulistana", "Queimada Nova", "Simões"]
}

# Função para mapear município para território
def map_municipio_to_territorio(municipio, territorios_municipios):
    for territorio, municipios in territorios_municipios.items():
        if municipio in municipios:
            return territorio
    return "Território Desconhecido"

def normalize_name(name):
    """Normalize municipality names by removing accents and converting to lowercase."""
    if not isinstance(name, str):
        return name
    name = ''.join(c for c in unicodedata.normalize('NFD', name) if unicodedata.category(c) != 'Mn')
    return name.lower()

class SistemaNotificacao:
    def __init__(self, root):
        self.root = root
        self.root.title("SISAD")
        self.root.state('zoomed')
        self.df = None
        self.db_path = None
        self.current_report_df = None
        self.current_report_type = None
        self.current_territorio = None
        self.current_map_path = None
        self.graph_buffer = None
        self.signature = "Desenvolvido por Teodoro Júnior - Enfermeiro | Analista de Dados em Saúde"
        self.ano = None
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('Helvetica', 10), padding=10)
        style.configure('TLabel', font=('Helvetica', 10))
        style.configure('TCombobox', font=('Helvetica', 10))
        style.map('TButton', background=[('active', '#2563eb'), ('disabled', '#d1d5db')],
                  foreground=[('active', '#ffffff'), ('disabled', '#6b7280')])

        self.main_frame = ttk.Frame(self.root, padding=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.title_label = ttk.Label(self.main_frame, text="Sistema de Análise de Dados da Dengue no Piauí (SISAD-PI)",
                                    font=('Helvetica', 16, 'bold'), foreground='#1e3a8a')
        self.title_label.pack(pady=10)

        self.signature_label = ttk.Label(self.main_frame, text=self.signature, font=('Helvetica', 8), foreground='#4b5563')
        self.signature_label.pack(pady=5)

        self.file_frame = ttk.Frame(self.main_frame)
        self.file_frame.pack(fill=tk.X, pady=10)

        self.btn_selecionar = ttk.Button(self.file_frame, text="Selecione o Arquivo DBF", command=self.selecionar_dbf)
        self.btn_selecionar.pack(side=tk.LEFT, padx=5)

        self.lbl_arquivo = ttk.Label(self.file_frame, text="Nenhum arquivo selecionado", foreground='#4b5563')
        self.lbl_arquivo.pack(side=tk.LEFT, padx=10)

        self.btn_frame = ttk.Frame(self.main_frame)
        self.btn_frame.pack(fill=tk.X, pady=10)

        self.btn_municipios = ttk.Button(self.btn_frame, text="Casos por Município", command=self.show_municipio_menu, state='disabled')
        self.btn_municipios.pack(side=tk.LEFT, padx=5, pady=5)

        self.municipio_menu = tk.Menu(self.root, tearoff=0)
        self.municipio_menu.add_command(label="Sorotipo", command=self.relatorio_municipio_sorotipo)
        self.municipio_menu.add_command(label="Classificação", command=self.relatorio_municipio_classificacao)
        self.municipio_menu.add_command(label="Evolução do Caso", command=self.relatorio_municipio_evolucao)
        self.municipio_menu.add_command(label="Classificação x Critério de Confirmação", command=self.show_classificacao_criterio_selection)
        self.municipio_menu.add_command(label="Critério de Confirmação", command=self.show_criterio_selection)
        self.municipio_menu.add_command(label="Sinais Clínicos", command=self.show_sinais_clinicos_selection)
        self.municipio_menu.add_command(label="Doenças Pré-existentes", command=self.show_doencas_preexistentes_selection)

        self.btn_territorios = ttk.Button(self.btn_frame, text="Território de Saúde", command=self.show_territorio_selection, state='disabled')
        self.btn_territorios.pack(side=tk.LEFT, padx=5, pady=5)

        self.territorio_menu = tk.Menu(self.root, tearoff=0)
        self.territorio_menu.add_command(label="Sorotipo", command=lambda: self.relatorio_territorio_sorotipo(self.selected_territorio))
        self.territorio_menu.add_command(label="Classificação", command=lambda: self.relatorio_territorio_classificacao(self.selected_territorio))
        self.territorio_menu.add_command(label="Evolução do Caso", command=lambda: self.relatorio_territorio_evolucao(self.selected_territorio))

        self.btn_georreferenciamento = ttk.Button(self.btn_frame, text="Georreferenciamento", command=self.show_georreferenciamento_selection, state='disabled')
        self.btn_georreferenciamento.pack(side=tk.LEFT, padx=5, pady=5)

        self.btn_idade_sexo = ttk.Button(self.btn_frame, text="Idade e Sexo", command=self.show_idade_sexo_selection, state='disabled')
        self.btn_idade_sexo.pack(side=tk.LEFT, padx=5, pady=5)

        self.btn_dashboard = ttk.Button(self.btn_frame, text="Dashboard", command=self.show_dashboard_selection, state='disabled')
        self.btn_dashboard.pack(side=tk.LEFT, padx=5, pady=5)

        self.btn_salvar = ttk.Button(self.btn_frame, text="Salvar Relatório", command=self.show_salvar_menu, state='disabled')
        self.btn_salvar.pack(side=tk.LEFT, padx=5, pady=5)

        self.salvar_menu = tk.Menu(self.root, tearoff=0)
        self.salvar_menu.add_command(label="Salvar em Excel", command=self.salvar_excel)
        self.salvar_menu.add_command(label="Salvar em PDF", command=self.salvar_pdf)

        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.text_frame = ttk.Frame(self.content_frame)
        self.text_frame.pack(fill=tk.BOTH, expand=True)

        self.txt_dados = tk.Text(self.text_frame, height=20, width=80, font=('Courier', 10), wrap=tk.NONE)
        self.scroll_y = ttk.Scrollbar(self.text_frame, orient=tk.VERTICAL, command=self.txt_dados.yview)
        self.scroll_x = ttk.Scrollbar(self.text_frame, orient=tk.HORIZONTAL, command=self.txt_dados.xview)
        self.txt_dados.configure(yscrollcommand=self.scroll_y.set, xscrollcommand=self.scroll_x.set)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.txt_dados.pack(fill=tk.BOTH, expand=True)

        self.status_var = tk.StringVar(value="Pronto")
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W,
                                   padding=5, foreground='#4b5563')
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def selecionar_dbf(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo DBF",
            filetypes=[("DBF files", "*.dbf")]
        )
        if file_path:
            try:
                self.db_path = file_path
                self.lbl_arquivo.configure(text=f"Arquivo selecionado: {os.path.basename(file_path)}")
                self.status_var.set("Carregando arquivo DBF...")
                
                dbf = DBF(file_path, encoding='latin-1')
                self.df = pd.DataFrame(iter(dbf))
                
                if self.df.empty:
                    messagebox.showerror("Erro", "O arquivo DBF está vazio ou não contém dados válidos.")
                    self.status_var.set("Erro ao carregar arquivo")
                    return
                
                if 'NU_ANO' in self.df.columns and not self.df['NU_ANO'].isna().all():
                    self.ano = str(int(self.df['NU_ANO'].dropna().iloc[0]))
                    logger.debug(f"Ano extraído do banco de dados: {self.ano}")
                else:
                    self.ano = "2025"
                    logger.warning("Coluna NU_ANO não encontrada ou vazia. Usando ano atual (2025).")
                
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, str(self.df.head(10)))
                
                self.btn_municipios.configure(state='normal')
                self.btn_territorios.configure(state='normal')
                self.btn_georreferenciamento.configure(state='normal')
                self.btn_idade_sexo.configure(state='normal')
                self.btn_dashboard.configure(state='normal')
                
                messagebox.showinfo("Sucesso", "Arquivo DBF carregado com sucesso!")
                self.status_var.set(f"Arquivo carregado: {os.path.basename(file_path)}")
            except Exception as e:
                logger.error(f"Falha ao carregar arquivo DBF: {str(e)}", exc_info=True)
                messagebox.showerror("Erro", f"Falha ao carregar o arquivo DBF: {str(e)}")
                self.df = None
                self.db_path = None
                self.ano = None
                self.lbl_arquivo.configure(text="Nenhum arquivo selecionado")
                self.btn_municipios.configure(state='disabled')
                self.btn_territorios.configure(state='disabled')
                self.btn_georreferenciamento.configure(state='disabled')
                self.btn_idade_sexo.configure(state='disabled')
                self.btn_dashboard.configure(state='disabled')
                self.btn_salvar.configure(state='disabled')
                self.status_var.set("Erro ao carregar arquivo")

    def show_municipio_menu(self):
        try:
            x = self.btn_municipios.winfo_rootx()
            y = self.btn_municipios.winfo_rooty() + self.btn_municipios.winfo_height()
            self.municipio_menu.post(x, y)
        except Exception as e:
            logger.error(f"Falha ao abrir menu: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao abrir o menu: {str(e)}")

    def show_territorio_selection(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Selecionar Território de Saúde")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Selecione o Território de Saúde:", font=('Helvetica', 10)).pack(pady=10)

        territorio_var = tk.StringVar()
        options = ["Todos"] + sorted(territorios_municipios.keys())
        combobox = ttk.Combobox(dialog, textvariable=territorio_var, values=options, 
                                state='readonly', font=('Helvetica', 10))
        combobox.pack(pady=5)
        combobox.set("Selecione um Território")

        def on_select():
            selected_territorio = territorio_var.get()
            if not selected_territorio or selected_territorio == "Selecione um Território":
                messagebox.showwarning("Aviso", "Por favor, selecione um Território de Saúde.")
                return
            self.selected_territorio = selected_territorio
            dialog.destroy()
            self.territorio_menu.delete(0, tk.END)
            self.territorio_menu.add_command(label="Sorotipo", command=lambda: self.relatorio_territorio_sorotipo(selected_territorio))
            self.territorio_menu.add_command(label="Classificação", command=lambda: self.relatorio_territorio_classificacao(selected_territorio))
            self.territorio_menu.add_command(label="Evolução do Caso", command=lambda: self.relatorio_territorio_evolucao(selected_territorio))
            try:
                x = self.btn_territorios.winfo_rootx()
                y = self.btn_territorios.winfo_rooty() + self.btn_territorios.winfo_height()
                self.territorio_menu.post(x, y)
            except Exception as e:
                logger.error(f"Falha ao abrir submenu: {str(e)}", exc_info=True)
                messagebox.showerror("Erro", f"Falha ao abrir o submenu: {str(e)}")

        ttk.Button(dialog, text="Confirmar", command=on_select).pack(pady=10)

    def show_georreferenciamento_selection(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return

        territorio_dialog = tk.Toplevel(self.root)
        territorio_dialog.title("Selecionar Território para Georreferenciamento")
        territorio_dialog.geometry("300x150")
        territorio_dialog.transient(self.root)
        territorio_dialog.grab_set()

        ttk.Label(territorio_dialog, text="Selecione o Território:", font=('Helvetica', 10)).pack(pady=10)

        territorio_var = tk.StringVar()
        options = ["Todos"] + sorted(territorios_municipios.keys())
        combobox = ttk.Combobox(territorio_dialog, textvariable=territorio_var, values=options, 
                                state='readonly', font=('Helvetica', 10))
        combobox.pack(pady=5)
        combobox.set("Selecione um Território")

        def on_territorio_select():
            selected_territorio = territorio_var.get()
            if not selected_territorio or selected_territorio == "Selecione um Território":
                messagebox.showwarning("Aviso", "Por favor, selecione um Território.")
                return
            territorio_dialog.destroy()

            categoria_dialog = tk.Toplevel(self.root)
            categoria_dialog.title("Selecionar Categoria para Georreferenciamento")
            categoria_dialog.geometry("300x150")
            categoria_dialog.transient(self.root)
            categoria_dialog.grab_set()

            ttk.Label(categoria_dialog, text="Selecione a Categoria:", font=('Helvetica', 10)).pack(pady=10)

            categoria_var = tk.StringVar()
            categoria_options = ["Sorotipo", "Classificação", "Evolução do Caso"]
            categoria_combobox = ttk.Combobox(categoria_dialog, textvariable=categoria_var, values=categoria_options, 
                                            state='readonly', font=('Helvetica', 10))
            categoria_combobox.pack(pady=5)
            categoria_combobox.set("Selecione uma Categoria")

            def on_categoria_select():
                selected_categoria = categoria_var.get()
                if not selected_categoria or selected_categoria == "Selecione uma Categoria":
                    messagebox.showwarning("Aviso", "Por favor, selecione uma Categoria.")
                    return
                categoria_dialog.destroy()

                subitem_dialog = tk.Toplevel(self.root)
                subitem_dialog.title(f"Selecionar Sub-item para {selected_categoria}")
                subitem_dialog.geometry("300x150")
                subitem_dialog.transient(self.root)
                subitem_dialog.grab_set()

                ttk.Label(subitem_dialog, text=f"Selecione o Sub-item para {selected_categoria}:", font=('Helvetica', 10)).pack(pady=10)

                subitem_var = tk.StringVar()
                if selected_categoria == "Sorotipo":
                    subitem_options = ["DEN 1", "DEN 2", "DEN 3", "DEN 4", "Desconhecido"]
                elif selected_categoria == "Classificação":
                    subitem_options = ["Dengue", "Dengue com Sinais de Alarme", "Dengue Grave", "Descartado", "Inconclusivo", "Em Branco"]
                else:
                    subitem_options = ["Cura", "Óbito por Dengue", "Óbito por Outras Causas", "Óbito em Investigação", "Ignorado", "Em Branco"]

                subitem_combobox = ttk.Combobox(subitem_dialog, textvariable=subitem_var, values=subitem_options, 
                                                state='readonly', font=('Helvetica', 10))
                subitem_combobox.pack(pady=5)
                subitem_combobox.set("Selecione um Sub-item")

                def on_subitem_select():
                    selected_subitem = subitem_var.get()
                    if not selected_subitem or selected_subitem == "Selecione um Sub-item":
                        messagebox.showwarning("Aviso", "Por favor, selecione um Sub-item.")
                        return
                    subitem_dialog.destroy()
                    self.relatorio_territorio_georreferenciamento(selected_territorio, selected_categoria, selected_subitem)

                ttk.Button(subitem_dialog, text="Confirmar", command=on_subitem_select).pack(pady=10)

            ttk.Button(categoria_dialog, text="Confirmar", command=on_categoria_select).pack(pady=10)

        ttk.Button(territorio_dialog, text="Confirmar", command=on_territorio_select).pack(pady=10)

    def show_idade_sexo_selection(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return

        territorio_dialog = tk.Toplevel(self.root)
        territorio_dialog.title("Selecionar Território para Idade e Sexo")
        territorio_dialog.geometry("300x150")
        territorio_dialog.transient(self.root)
        territorio_dialog.grab_set()

        ttk.Label(territorio_dialog, text="Selecione o Território:", font=('Helvetica', 10)).pack(pady=10)

        territorio_var = tk.StringVar()
        options = ["Todos"] + sorted(territorios_municipios.keys())
        combobox = ttk.Combobox(territorio_dialog, textvariable=territorio_var, values=options, 
                                state='readonly', font=('Helvetica', 10))
        combobox.pack(pady=5)
        combobox.set("Selecione um Território")

        def on_territorio_select():
            selected_territorio = territorio_var.get()
            if not selected_territorio or selected_territorio == "Selecione um Território":
                messagebox.showwarning("Aviso", "Por favor, selecione um Território.")
                return
            territorio_dialog.destroy()
            self.relatorio_idade_sexo(selected_territorio)

        ttk.Button(territorio_dialog, text="Confirmar", command=on_territorio_select).pack(pady=10)

    def show_dashboard_selection(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return

        self.create_dashboard_window()

    def show_classificacao_criterio_selection(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return

        selection_window = tk.Toplevel(self.root)
        selection_window.title("Selecionar Classificações - Classificação x Critério de Confirmação")
        selection_window.geometry("400x300")
        selection_window.transient(self.root)
        selection_window.grab_set()

        ttk.Label(selection_window, text="Selecione as Classificações:", font=('Helvetica', 10)).pack(pady=5)

        classificacoes = ['Dengue', 'Dengue com Sinais de Alarme', 'Dengue Grave', 'Descartado', 'Inconclusivo', 'Em Branco']
        classificacao_vars = {classi: tk.BooleanVar(value=True) for classi in classificacoes}

        check_frame = ttk.Frame(selection_window)
        check_frame.pack(pady=5, fill=tk.BOTH, expand=True)

        for classi, var in classificacao_vars.items():
            ttk.Checkbutton(check_frame, text=classi, variable=var).pack(anchor='w', padx=10)

        def gerar_relatorio():
            selected_classificacoes = [classi for classi, var in classificacao_vars.items() if var.get()]
            if not selected_classificacoes:
                messagebox.showwarning("Aviso", "Selecione pelo menos uma classificação.")
                return
            selection_window.destroy()
            self.relatorio_municipio_classificacao_criterio(selected_classificacoes)

        ttk.Button(selection_window, text="Gerar Relatório", command=gerar_relatorio).pack(pady=10)

    def show_criterio_selection(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return

        selection_window = tk.Toplevel(self.root)
        selection_window.title("Selecionar Critérios de Confirmação")
        selection_window.geometry("400x300")
        selection_window.transient(self.root)
        selection_window.grab_set()

        ttk.Label(selection_window, text="Selecione os Critérios de Confirmação:", font=('Helvetica', 10)).pack(pady=5)

        criterios = ['Laboratorial', 'Clínico Epidemiológico', 'Em Investigação']
        criterio_vars = {crit: tk.BooleanVar(value=True) for crit in criterios}

        check_frame = ttk.Frame(selection_window)
        check_frame.pack(pady=5, fill=tk.BOTH, expand=True)

        for crit, var in criterio_vars.items():
            ttk.Checkbutton(check_frame, text=crit, variable=var).pack(anchor='w', padx=10)

        def gerar_relatorio():
            selected_criterios = [crit for crit, var in criterio_vars.items() if var.get()]
            if not selected_criterios:
                messagebox.showwarning("Aviso", "Selecione pelo menos um critério de confirmação.")
                return
            selection_window.destroy()
            self.relatorio_municipio_criterio(selected_criterios)

        ttk.Button(selection_window, text="Gerar Relatório", command=gerar_relatorio).pack(pady=10)

    def show_sinais_clinicos_selection(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return

        selection_window = tk.Toplevel(self.root)
        selection_window.title("Selecionar Sinais Clínicos")
        selection_window.geometry("400x500")
        selection_window.transient(self.root)
        selection_window.grab_set()

        ttk.Label(selection_window, text="Selecione os Sinais Clínicos:", font=('Helvetica', 10)).pack(pady=5)

        sinais_clinicos = [
            'Febre', 'Mialgia', 'Cefaleia', 'Exantema', 'Vômito', 'Náuseas', 
            'Dor nas costas', 'Conjuntivite', 'Artrite', 'Artralgia intensa', 
            'Petéquias', 'Leucopenia', 'Prova do laço positiva', 'Dor retroorbital'
        ]
        sinais_vars = {sinal: tk.BooleanVar(value=True) for sinal in sinais_clinicos}

        check_frame = ttk.Frame(selection_window)
        check_frame.pack(pady=5, fill=tk.BOTH, expand=True)

        for sinal, var in sinais_vars.items():
            ttk.Checkbutton(check_frame, text=sinal, variable=var).pack(anchor='w', padx=10)

        def gerar_relatorio():
            selected_sinais = [sinal for sinal, var in sinais_vars.items() if var.get()]
            if not selected_sinais:
                messagebox.showwarning("Aviso", "Selecione pelo menos um sinal clínico.")
                return
            selection_window.destroy()
            self.relatorio_municipio_sinais_clinicos(selected_sinais)

        ttk.Button(selection_window, text="Gerar Relatório", command=gerar_relatorio).pack(pady=10)

    def show_doencas_preexistentes_selection(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return

        selection_window = tk.Toplevel(self.root)
        selection_window.title("Selecionar Doenças Pré-existentes")
        selection_window.geometry("400x350")
        selection_window.transient(self.root)
        selection_window.grab_set()

        ttk.Label(selection_window, text="Selecione as Doenças Pré-existentes:", font=('Helvetica', 10)).pack(pady=5)

        doencas = [
            'Diabetes', 'Doenças hematológicas', 'Hepatopatias', 'Doença renal crônica',
            'Hipertensão arterial', 'Doença ácido-péptica', 'Doenças auto-imunes'
        ]
        doencas_vars = {doenca: tk.BooleanVar(value=True) for doenca in doencas}

        check_frame = ttk.Frame(selection_window)
        check_frame.pack(pady=5, fill=tk.BOTH, expand=True)

        for doenca, var in doencas_vars.items():
            ttk.Checkbutton(check_frame, text=doenca, variable=var).pack(anchor='w', padx=10)

        def gerar_relatorio():
            selected_doencas = [doenca for doenca, var in doencas_vars.items() if var.get()]
            if not selected_doencas:
                messagebox.showwarning("Aviso", "Selecione pelo menos uma doença pré-existente.")
                return
            selection_window.destroy()
            self.relatorio_municipio_doencas_preexistentes(selected_doencas)

        ttk.Button(selection_window, text="Gerar Relatório", command=gerar_relatorio).pack(pady=10)

    def show_salvar_menu(self):
        try:
            x = self.btn_salvar.winfo_rootx()
            y = self.btn_salvar.winfo_rooty() + self.btn_salvar.winfo_height()
            self.salvar_menu.post(x, y)
        except Exception as e:
            logger.error(f"Falha ao abrir menu de salvar: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao abrir o menu de salvar: {str(e)}")

    def salvar_excel(self):
        if self.current_report_df is None:
            messagebox.showwarning("Aviso", "Nenhum relatório para salvar.")
            return

        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Salvar Relatório como Excel"
            )
            if file_path:
                self.current_report_df.to_excel(file_path, index=False)
                messagebox.showinfo("Sucesso", f"Relatório salvo como Excel em: {file_path}")
                self.status_var.set(f"Relatório salvo como Excel em: {file_path}")
        except Exception as e:
            logger.error(f"Falha ao salvar Excel: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao salvar o relatório como Excel: {str(e)}")
            self.status_var.set("Erro ao salvar relatório")

    def salvar_pdf(self, figures=None):
        if self.current_report_df is None:
            messagebox.showwarning("Aviso", "Nenhum relatório para salvar.")
            return

        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                title="Salvar Relatório como PDF"
            )
            if not file_path:
                logger.info("Nenhum caminho de arquivo selecionado para salvar o PDF.")
                return

            logger.info(f"Iniciando geração do PDF: {file_path}")

            page_size = landscape(A4)
            page_width = A4[1] - 60
            page_height = A4[0] - 60

            doc = SimpleDocTemplate(file_path, pagesize=page_size, leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=50)
            elements = []
            buffers_to_close = []

            styles = getSampleStyleSheet()
            title_style = styles['Heading1']
            title_style.fontSize = 16
            title_style.alignment = 1
            subtitle_style = styles['Heading2']
            subtitle_style.fontSize = 14
            footer_style = ParagraphStyle(
                name='FooterStyle',
                fontSize=8,
                leading=10,
                alignment=1,
                textColor=colors.grey
            )
            logger.debug("Estilos definidos com sucesso.")

            report_title = f"Relatório: {self.current_territorio} ({self.current_report_type}) - {self.ano}"
            elements.append(Paragraph(report_title, title_style))
            elements.append(Spacer(1, 12))
            logger.debug("Título do relatório adicionado ao PDF.")

            data = [self.current_report_df.columns.tolist()] + self.current_report_df.values.tolist()
            logger.debug(f"Dados da tabela preparados: {len(data)} linhas.")

            num_cols = len(self.current_report_df.columns)
            col_widths = []
            total_width = page_width
            min_col_width = 40

            col_lengths = [len(str(col)) for col in self.current_report_df.columns]
            logger.debug(f"Comprimentos dos nomes das colunas: {col_lengths}")
            
            median_length = statistics.median(col_lengths)
            logger.debug(f"Mediana do comprimento dos nomes das colunas: {median_length}")
            
            base_width = max(min_col_width, median_length * 5)
            col_widths = [base_width for _ in range(num_cols)]

            total_col_width = sum(col_widths)
            if total_col_width > total_width:
                scale_factor = total_width / total_col_width
                col_widths = [w * scale_factor for w in col_widths]
            elif total_col_width < total_width:
                scale_factor = total_width / total_col_width
                col_widths = [w * scale_factor for w in col_widths]
            logger.debug(f"Larguras das colunas ajustadas: {col_widths}")

            font_size = max(6, min(8, 10 - (num_cols - 2) * 0.5))
            logger.debug(f"Tamanho da fonte ajustado: {font_size}")

            table = Table(data, colWidths=col_widths, repeatRows=1)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), font_size),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('WORDWRAP', (0, 0), (-1, -1), True),
                ('BOX', (0, 0), (-1, -1), 1, colors.black),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            ]))
            table.hAlign = 'CENTER'
            elements.append(table)
            logger.debug("Tabela adicionada ao PDF.")

            if self.current_territorio != f"Casos por Município - {self.ano}":
                elements.append(PageBreak())
                logger.debug("Quebra de página adicionada para o gráfico/mapa.")

                if self.current_report_type == "Georreferenciamento" and self.current_map_path:
                    logger.info(f"Tentando adicionar o mapa ao PDF: {self.current_map_path}")
                    if not os.path.exists(self.current_map_path):
                        logger.warning("Arquivo do mapa não encontrado.")
                        messagebox.showwarning("Aviso", "O arquivo do mapa não foi encontrado. O PDF será gerado sem o mapa.")
                        self.status_var.set("Erro: Arquivo do mapa não encontrado")
                    else:
                        try:
                            elements.append(Paragraph("Mapa Georreferenciado", subtitle_style))
                            elements.append(Spacer(1, 12))
                            logger.debug("Abrindo imagem do mapa para inclusão no PDF.")
                            with open(self.current_map_path, 'rb') as map_file:
                                img = PILImage.open(map_file)
                                img_width, img_height = img.size
                                aspect = img_height / float(img_width)

                                max_img_width = page_width - 20
                                max_img_height = page_height - 60
                                if img_width > max_img_width:
                                    img_width = max_img_width
                                    img_height = img_width * aspect
                                if img_height > max_img_height:
                                    img_height = max_img_height
                                    img_width = img_height / aspect

                                logger.debug(f"Dimensões ajustadas do mapa: largura={img_width}, altura={img_height}")
                                elements.append(Image(self.current_map_path, width=img_width, height=img_height, hAlign='CENTER'))
                                logger.debug("Mapa adicionado ao PDF com sucesso.")
                        except Exception as e:
                            logger.error(f"Falha ao carregar o mapa: {str(e)}", exc_info=True)
                            messagebox.showwarning("Aviso", f"Falha ao carregar o mapa: {str(e)}. O PDF será gerado sem o mapa.")
                            self.status_var.set("Erro ao carregar o mapa")
                elif self.current_report_type == "Idade e Sexo" and hasattr(self, 'graph_buffer') and self.graph_buffer:
                    logger.info("Adicionando gráfico de Idade e Sexo ao PDF.")
                    try:
                        elements.append(Paragraph("Gráfico de Distribuição por Idade e Sexo", subtitle_style))
                        elements.append(Spacer(1, 12))
                        self.graph_buffer.seek(0)
                        img = PILImage.open(self.graph_buffer)
                        img_width, img_height = img.size
                        aspect = img_height / float(img_width)

                        max_img_width = page_width - 20
                        max_img_height = page_height - 60
                        if img_width > max_img_width:
                            img_width = max_img_width
                            img_height = img_width * aspect
                        if img_height > max_img_height:
                            img_height = max_img_height
                            img_width = img_height / aspect

                        logger.debug(f"Dimensões ajustadas do gráfico: largura={img_width}, altura={img_height}")
                        elements.append(Image(self.graph_buffer, width=img_width, height=img_height, hAlign='CENTER'))
                        logger.debug("Gráfico de Idade e Sexo adicionado ao PDF com sucesso.")
                        buffers_to_close.append(self.graph_buffer)
                    except Exception as e:
                        logger.error(f"Falha ao carregar gráfico de Idade e Sexo: {str(e)}", exc_info=True)
                        messagebox.showwarning("Aviso", f"Falha ao carregar o gráfico: {str(e)}. O PDF será gerado sem o gráfico.")
                        self.status_var.set("Erro ao carregar o gráfico")
                        if hasattr(self, 'graph_buffer') and not self.graph_buffer.closed:
                            self.graph_buffer.close()
                elif self.current_report_type == "Dashboard" and figures:
                    logger.info("Adicionando dashboard ao PDF.")
                    try:
                        for idx, fig in enumerate(figures):
                            elements.append(PageBreak())
                            elements.append(Paragraph(f"Dashboard - Página {idx + 1}", subtitle_style))
                            elements.append(Spacer(1, 12))
                            
                            buf = io.BytesIO()
                            fig.savefig(buf, format='png', bbox_inches='tight', dpi=150)
                            buf.seek(0)
                            
                            img = PILImage.open(buf)
                            img_width, img_height = img.size
                            aspect = img_height / float(img_width)

                            max_img_width = page_width - 20
                            max_img_height = page_height - 60
                            if img_width > max_img_width:
                                img_width = max_img_width
                                img_height = img_width * aspect
                            if img_height > max_img_height:
                                img_height = max_img_height
                                img_width = img_height / aspect

                            logger.debug(f"Dimensões ajustadas do dashboard (página {idx + 1}): largura={img_width}, altura={img_height}")
                            elements.append(Image(buf, width=img_width, height=img_height, hAlign='CENTER'))
                            logger.debug(f"Dashboard página {idx + 1} adicionado ao PDF com sucesso.")
                            buffers_to_close.append(buf)
                    except Exception as e:
                        logger.error(f"Falha ao carregar o dashboard: {str(e)}", exc_info=True)
                        messagebox.showwarning("Aviso", f"Falha ao carregar o dashboard: {str(e)}. O PDF será gerado sem o dashboard.")
                        self.status_var.set("Erro ao carregar o dashboard")
                else:
                    logger.info("Gerando gráfico para o relatório.")
                    if self.current_report_type == "Sorotipo":
                        columns_to_plot = ['DEN 1', 'DEN 2', 'DEN 3', 'DEN 4']
                        title = f"Sorotipos em {self.current_territorio} - {self.ano}"
                    elif self.current_report_type == "Classificação":
                        columns_to_plot = ['Dengue', 'Dengue com Sinais de Alarme', 'Dengue Grave', 'Descartado', 'Inconclusivo', 'Em Branco']
                        title = f"Classificação em {self.current_territorio} - {self.ano}"
                    else:
                        columns_to_plot = ['Cura', 'Óbito por Dengue', 'Óbito por Outras Causas', 'Óbito em Investigação', 'Ignorado', 'Em Branco']
                        title = f"Evolução do Caso em {self.current_territorio} - {self.ano}"

                    plt.figure(figsize=(10, 6))
                    bar_width = 0.15
                    index = range(len(self.current_report_df) - 1)
                    colors_palette = sns.color_palette("Set2", len(columns_to_plot))

                    for i, col in enumerate(columns_to_plot):
                        values = self.current_report_df[col].iloc[:-1]
                        bars = plt.bar([x + bar_width * i for x in index], values, 
                                       bar_width, label=col, color=colors_palette[i], edgecolor='black')
                        for bar in bars:
                            height = bar.get_height()
                            if height > 0:
                                plt.text(bar.get_x() + bar.get_width() / 2, height, f'{int(height)}', 
                                        ha='center', va='bottom', fontsize=6, color='black')

                    plt.xlabel('Territórios' if 'Todos' in self.current_territorio else 'Municípios', fontsize=10)
                    plt.ylabel('Número de Casos (Escala Logarítmica)', fontsize=10)
                    plt.title(title, fontsize=12, pad=20)
                    plt.xticks([i + bar_width * (len(columns_to_plot) - 1) / 2 for i in index], 
                               self.current_report_df['Nome_Municipio'].iloc[:-1] if 'Todos' not in self.current_territorio else self.current_report_df['Território'].iloc[:-1], 
                               rotation=45, ha='right', fontsize=8)
                    plt.yscale('log')
                    plt.ylim(bottom=0.1)
                    plt.legend(fontsize=8, loc='upper right', frameon=True, edgecolor='black')
                    plt.grid(True, which="both", ls="--", alpha=0.7, axis='y')
                    plt.tight_layout()

                    buf = io.BytesIO()
                    plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
                    plt.close()
                    buf.seek(0)
                    logger.debug("Gráfico gerado e salvo em buffer temporário.")

                    try:
                        logger.debug("Abrindo imagem do gráfico para inclusão no PDF.")
                        img = PILImage.open(buf)
                        img_width, img_height = img.size
                        aspect = img_height / float(img_width)

                        max_img_width = page_width - 20
                        max_img_height = page_height - 60
                        if img_width > max_img_width:
                            img_width = max_img_width
                            img_height = img_width * aspect
                        if img_height > max_img_height:
                            img_height = max_img_height
                            img_width = img_height / aspect

                        logger.debug(f"Dimensões ajustadas do gráfico: largura={img_width}, altura={img_height}")
                        elements.append(Paragraph("Gráfico de Distribuição", subtitle_style))
                        elements.append(Spacer(1, 12))
                        if buf.closed:
                            logger.error("Buffer do gráfico está fechado antes de ser usado pelo reportlab.")
                            raise ValueError("Buffer do gráfico foi fechado prematuramente.")
                        elements.append(Image(buf, width=img_width, height=img_height, hAlign='CENTER'))
                        logger.debug("Gráfico adicionado ao PDF com sucesso.")
                        buffers_to_close.append(buf)
                    except Exception as e:
                        logger.error(f"Falha ao carregar o gráfico: {str(e)}", exc_info=True)
                        messagebox.showwarning("Aviso", f"Falha ao carregar o gráfico: {str(e)}. O PDF será gerado sem o gráfico.")
                        self.status_var.set("Erro ao carregar o gráfico")
                        buf.close()

            def add_footer(canvas, doc):
                canvas.saveState()
                canvas.setFont('Helvetica', 8)
                canvas.setFillColor(colors.grey)
                canvas.drawCentredString(page_size[0] / 2, 0.5 * inch, self.signature)
                canvas.restoreState()

            frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height - inch)
            template = PageTemplate(id='AllPages', frames=frame, onPage=add_footer)
            doc.addPageTemplates([template])
            logger.debug("Template de página com footer configurado.")

            logger.info("Iniciando construção do PDF.")
            doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)
            logger.info("PDF gerado com sucesso.")

            for buf in buffers_to_close:
                if not buf.closed:
                    buf.close()
                    logger.debug("Buffer do gráfico fechado após construção do PDF.")

            messagebox.showinfo("Sucesso", f"Relatório salvo como PDF em: {file_path}")
            self.status_var.set(f"Relatório salvo como PDF em: {file_path}")
        except Exception as e:
            logger.error(f"Falha ao salvar PDF: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao salvar o relatório como PDF: {str(e)}")
            self.status_var.set("Erro ao salvar relatório")
            for buf in buffers_to_close:
                if not buf.closed:
                    buf.close()
                    logger.debug("Buffer do gráfico fechado após erro.")

    def salvar_png(self, figs):
        try:
            if not figs:
                messagebox.showwarning("Aviso", "Nenhum gráfico para salvar.")
                return

            # Para cada figura, pedir ao usuário um local para salvar
            for idx, fig in enumerate(figs):
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".png",
                    filetypes=[("PNG files", "*.png"), ("All files", "*.*")],
                    title=f"Salvar Dashboard Página {idx + 1} como PNG",
                    initialfile=f"dashboard_page_{idx + 1}.png"
                )
                if file_path:
                    fig.savefig(file_path, format='png', dpi=150, bbox_inches='tight')
                    messagebox.showinfo("Sucesso", f"Página {idx + 1} salva como PNG em: {file_path}")
                    self.status_var.set(f"Página {idx + 1} salva como PNG em: {file_path}")
        except Exception as e:
            logger.error(f"Falha ao salvar PNG: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao salvar o dashboard como PNG: {str(e)}")
            self.status_var.set("Erro ao salvar dashboard")

    def relatorio_municipio_sorotipo(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return
        
        try:
            self.status_var.set("Gerando relatório de Casos por Município (Sorotipo)...")
            logger.info("Iniciando geração do relatório de Casos por Município (Sorotipo).")
            
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
            
            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set("Nenhum dado para o Piauí (Sorotipo)")
                logger.info("Nenhum dado encontrado para o Piauí (Sorotipo).")
                return
            
            base_df = pd.DataFrame(list(MUNICIPIOS_PIAUI.items()), columns=['ID_MN_RESI', 'Nome_Municipio'])
            
            df_dengue = df_filtered[df_filtered['ID_AGRAVO'] == 'A90'].copy()
            logger.info(f"Tamanho do DataFrame após filtro por Dengue (A90): {len(df_dengue)}")
            
            df_dengue['SOROTIPO'] = df_dengue['SOROTIPO'].astype(str).replace('', '0').replace('nan', '0')
            
            pivot_table = pd.pivot_table(
                df_dengue,
                index='ID_MN_RESI',
                columns='SOROTIPO',
                aggfunc='size',
                fill_value=0
            )
            
            sorotipo_map = {'1': 'DEN 1', '2': 'DEN 2', '3': 'DEN 3', '4': 'DEN 4'}
            for sorotipo in ['1', '2', '3', '4']:
                if sorotipo not in pivot_table.columns:
                    pivot_table[sorotipo] = 0
            
            pivot_table.columns = [sorotipo_map.get(col, f'Sorotipo {col}') for col in pivot_table.columns]
            
            sorotipo_cols = ['DEN 1', 'DEN 2', 'DEN 3', 'DEN 4']
            pivot_table = pivot_table[[col for col in sorotipo_cols if col in pivot_table.columns]]
            
            for col in sorotipo_cols:
                if col not in pivot_table.columns:
                    pivot_table[col] = 0
            
            pivot_table = pivot_table.astype(int)
            
            total_casos = df_dengue.groupby('ID_MN_RESI').size().astype(int)
            pivot_table['Total de Casos'] = total_casos
            
            result_table = base_df.merge(pivot_table, on='ID_MN_RESI', how='left').fillna(0)
            
            numeric_cols = ['DEN 1', 'DEN 2', 'DEN 3', 'DEN 4', 'Total de Casos']
            result_table[numeric_cols] = result_table[numeric_cols].astype(int)
            
            columns = ['Nome_Municipio', 'DEN 1', 'DEN 2', 'DEN 3', 'DEN 4', 'Total de Casos']
            result_table = result_table[columns]
            
            result_table = result_table.sort_values(by='Nome_Municipio')
            
            totals = {
                'Nome_Municipio': 'Total',
                'DEN 1': result_table['DEN 1'].sum(),
                'DEN 2': result_table['DEN 2'].sum(),
                'DEN 3': result_table['DEN 3'].sum(),
                'DEN 4': result_table['DEN 4'].sum(),
                'Total de Casos': result_table['Total de Casos'].sum()
            }
            totals_df = pd.DataFrame([totals], columns=columns)
            
            result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            self.current_report_df = result_table
            self.current_report_type = "Sorotipo"
            self.current_territorio = f"Casos por Município - {self.ano}"
            
            self.btn_salvar.configure(state='normal')
            
            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Sorotipo)\n\n")
            self.txt_dados.insert(tk.END, result_table.to_string(index=False))
            self.status_var.set(f"Relatório de {self.current_territorio} (Sorotipo) gerado")
            logger.info("Relatório de Casos por Município (Sorotipo) gerado com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
            self.status_var.set("Erro ao gerar relatório")

    def relatorio_municipio_classificacao(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return
        
        try:
            self.status_var.set("Gerando relatório de Casos por Município (Classificação)...")
            logger.info("Iniciando geração do relatório de Casos por Município (Classificação).")
            
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
            
            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set("Nenhum dado para o Piauí (Classificação)")
                logger.info("Nenhum dado encontrado para o Piauí (Classificação).")
                return
            
            base_df = pd.DataFrame(list(MUNICIPIOS_PIAUI.items()), columns=['ID_MN_RESI', 'Nome_Municipio'])
            
            df_mapped = df_filtered.copy()
            df_mapped['CLASSI_FIN'] = pd.to_numeric(df_mapped['CLASSI_FIN'], errors='coerce').fillna(0).astype(int)
            logger.info(f"Valores únicos de CLASSI_FIN após mapeamento: {df_mapped['CLASSI_FIN'].unique()}")
            
            classi_fin_map = {
                10: 'Dengue', 11: 'Dengue com Sinais de Alarme', 12: 'Dengue Grave',
                5: 'Descartado', 8: 'Inconclusivo', 0: 'Em Branco'
            }
            df_mapped['Classificação'] = df_mapped['CLASSI_FIN'].map(classi_fin_map).fillna('Em Branco')
            
            pivot_table = pd.pivot_table(
                df_mapped,
                index='ID_MN_RESI',
                columns='Classificação',
                aggfunc='size',
                fill_value=0
            )
            
            expected_classifications = ['Dengue', 'Dengue com Sinais de Alarme', 'Dengue Grave', 'Descartado', 'Inconclusivo', 'Em Branco']
            for classi in expected_classifications:
                if classi not in pivot_table.columns:
                    pivot_table[classi] = 0
            
            pivot_table = pivot_table[expected_classifications]
            
            pivot_table = pivot_table.astype(int)
            
            total_casos = df_mapped.groupby('ID_MN_RESI').size().astype(int)
            pivot_table['Total de Casos'] = total_casos
            
            result_table = base_df.merge(pivot_table, on='ID_MN_RESI', how='left').fillna(0)
            
            numeric_cols = expected_classifications + ['Total de Casos']
            result_table[numeric_cols] = result_table[numeric_cols].astype(int)
            
            columns = ['Nome_Municipio'] + expected_classifications + ['Total de Casos']
            result_table = result_table[columns]
            
            result_table = result_table.sort_values(by='Nome_Municipio')
            
            totals = {'Nome_Municipio': 'Total'}
            for col in expected_classifications:
                totals[col] = result_table[col].sum()
            totals['Total de Casos'] = result_table['Total de Casos'].sum()
            totals_df = pd.DataFrame([totals], columns=columns)
            
            result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            self.current_report_df = result_table
            self.current_report_type = "Classificação"
            self.current_territorio = f"Casos por Município - {self.ano}"
            
            self.btn_salvar.configure(state='normal')
            
            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Classificação)\n\n")
            self.txt_dados.insert(tk.END, result_table.to_string(index=False))
            self.status_var.set(f"Relatório de {self.current_territorio} (Classificação) gerado")
            logger.info("Relatório de Casos por Município (Classificação) gerado com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
            self.status_var.set("Erro ao gerar relatório")

    def relatorio_municipio_evolucao(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return
        
        try:
            self.status_var.set("Gerando relatório de Casos por Município (Evolução do Caso)...")
            logger.info("Iniciando geração do relatório de Casos por Município (Evolução do Caso).")
            
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
            
            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set("Nenhum dado para o Piauí (Evolução do Caso)")
                logger.info("Nenhum dado encontrado para o Piauí (Evolução do Caso).")
                return
            
            base_df = pd.DataFrame(list(MUNICIPIOS_PIAUI.items()), columns=['ID_MN_RESI', 'Nome_Municipio'])
            
            df_mapped = df_filtered.copy()
            df_mapped['EVOLUCAO'] = pd.to_numeric(df_mapped['EVOLUCAO'], errors='coerce').fillna(0).astype(int)
            logger.info(f"Valores únicos de EVOLUCAO após mapeamento: {df_mapped['EVOLUCAO'].unique()}")
            
            evolucao_map = {
                1: 'Cura', 2: 'Óbito por Dengue', 3: 'Óbito por Outras Causas',
                4: 'Óbito em Investigação', 9: 'Ignorado', 0: 'Em Branco'
            }
            df_mapped['Evolução'] = df_mapped['EVOLUCAO'].map(evolucao_map).fillna('Em Branco')
            
            pivot_table = pd.pivot_table(
                df_mapped,
                index='ID_MN_RESI',
                columns='Evolução',
                aggfunc='size',
                fill_value=0
            )
            
            expected_evolutions = ['Cura', 'Óbito por Dengue', 'Óbito por Outras Causas', 'Óbito em Investigação', 'Ignorado', 'Em Branco']
            for evol in expected_evolutions:
                if evol not in pivot_table.columns:
                    pivot_table[evol] = 0
            
            pivot_table = pivot_table[expected_evolutions]
            
            pivot_table = pivot_table.astype(int)
            
            total_casos = df_mapped.groupby('ID_MN_RESI').size().astype(int)
            pivot_table['Total de Casos'] = total_casos
            
            result_table = base_df.merge(pivot_table, on='ID_MN_RESI', how='left').fillna(0)
            
            numeric_cols = expected_evolutions + ['Total de Casos']
            result_table[numeric_cols] = result_table[numeric_cols].astype(int)
            
            columns = ['Nome_Municipio'] + expected_evolutions + ['Total de Casos']
            result_table = result_table[columns]
            
            result_table = result_table.sort_values(by='Nome_Municipio')
            
            totals = {'Nome_Municipio': 'Total'}
            for col in expected_evolutions:
                totals[col] = result_table[col].sum()
            totals['Total de Casos'] = result_table['Total de Casos'].sum()
            totals_df = pd.DataFrame([totals], columns=columns)
            
            result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            self.current_report_df = result_table
            self.current_report_type = "Evolução do Caso"
            self.current_territorio = f"Casos por Município - {self.ano}"
            
            self.btn_salvar.configure(state='normal')
            
            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Evolução do Caso)\n\n")
            self.txt_dados.insert(tk.END, result_table.to_string(index=False))
            self.status_var.set(f"Relatório de {self.current_territorio} (Evolução do Caso) gerado")
            logger.info("Relatório de Casos por Município (Evolução do Caso) gerado com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
            self.status_var.set("Erro ao gerar relatório")

    def relatorio_municipio_classificacao_criterio(self, selected_classificacoes):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return
        
        try:
            self.status_var.set("Gerando relatório de Casos por Município (Classificação x Critério de Confirmação)...")
            logger.info("Iniciando geração do relatório de Casos por Município (Classificação x Critério de Confirmação).")
            
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
            
            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set("Nenhum dado para o Piauí (Classificação x Critério de Confirmação)")
                logger.info("Nenhum dado encontrado para o Piauí (Classificação x Critério de Confirmação).")
                return
            
            base_df = pd.DataFrame(list(MUNICIPIOS_PIAUI.items()), columns=['ID_MN_RESI', 'Nome_Municipio'])
            
            df_mapped = df_filtered.copy()
            df_mapped['CLASSI_FIN'] = pd.to_numeric(df_mapped['CLASSI_FIN'], errors='coerce').fillna(0).astype(int)
            logger.info(f"Valores únicos de CLASSI_FIN após mapeamento: {df_mapped['CLASSI_FIN'].unique()}")
            
            classi_fin_map = {
                10: 'Dengue', 11: 'Dengue com Sinais de Alarme', 12: 'Dengue Grave',
                5: 'Descartado', 8: 'Inconclusivo', 0: 'Em Branco'
            }
            df_mapped['Classificação'] = df_mapped['CLASSI_FIN'].map(classi_fin_map).fillna('Em Branco')
            
            df_mapped = df_mapped[df_mapped['Classificação'].isin(selected_classificacoes)].copy()
            logger.info(f"Tamanho do DataFrame após filtro por classificações selecionadas: {len(df_mapped)}")
            if df_mapped.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para as classificações selecionadas.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set("Nenhum dado para as classificações selecionadas")
                logger.info(f"Nenhum dado encontrado para as Classificações: {selected_classificacoes}.")
                return
            
            criterio_map = {
                1: 'Laboratorial',
                2: 'Clínico Epidemiológico',
                3: 'Em Investigação',
                4: 'Desconhecido'
            }
            df_mapped['CRITERIO'] = pd.to_numeric(df_mapped['CRITERIO'], errors='coerce').fillna(4).astype(int)
            df_mapped['Critério de Confirmação'] = df_mapped['CRITERIO'].map(criterio_map).fillna('Desconhecido')
            
            df_mapped['Classificação_Critério'] = df_mapped['Classificação'] + ' - ' + df_mapped['Critério de Confirmação']
            
            pivot_table = pd.pivot_table(
                df_mapped,
                index='ID_MN_RESI',
                columns='Classificação_Critério',
                aggfunc='size',
                fill_value=0
            )
            
            expected_combinations = []
            criterios = ['Laboratorial', 'Clínico Epidemiológico', 'Em Investigação']
            for classi in selected_classificacoes:
                for crit in criterios:
                    expected_combinations.append(f"{classi} - {crit}")
            
            for combo in expected_combinations:
                if combo not in pivot_table.columns:
                    pivot_table[combo] = 0
            
            pivot_table = pivot_table[expected_combinations]
            
            pivot_table = pivot_table.astype(int)
            
            pivot_table['Total de Casos'] = pivot_table[expected_combinations].sum(axis=1)
            
            result_table = base_df.merge(pivot_table, on='ID_MN_RESI', how='left').fillna(0)
            
            numeric_cols = expected_combinations + ['Total de Casos']
            result_table[numeric_cols] = result_table[numeric_cols].astype(int)
            
            columns = ['Nome_Municipio'] + expected_combinations + ['Total de Casos']
            result_table = result_table[columns]
            
            result_table = result_table.sort_values(by='Nome_Municipio')
            
            totals = {'Nome_Municipio': 'Total'}
            for col in expected_combinations + ['Total de Casos']:
                totals[col] = result_table[col].sum()
            totals_df = pd.DataFrame([totals], columns=columns)
            
            result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            self.current_report_df = result_table
            self.current_report_type = "Classificação x Critério de Confirmação"
            self.current_territorio = f"Casos por Município - {self.ano}"
            
            self.btn_salvar.configure(state='normal')
            
            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Classificação x Critério de Confirmação)\n")
            self.txt_dados.insert(tk.END, f"Classificações Selecionadas: {', '.join(selected_classificacoes)}\n\n")
            self.txt_dados.insert(tk.END, result_table.to_string(index=False))
            self.status_var.set(f"Relatório de {self.current_territorio} (Classificação x Critério de Confirmação) gerado")
            logger.info("Relatório de Casos por Município (Classificação x Critério de Confirmação) gerado com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
            self.status_var.set("Erro ao gerar relatório")

    def relatorio_municipio_criterio(self, selected_criterios):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return
        
        try:
            self.status_var.set("Gerando relatório de Casos por Município (Critério de Confirmação)...")
            logger.info("Iniciando geração do relatório de Casos por Município (Critério de Confirmação).")
            
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
            
            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set("Nenhum dado para o Piauí (Critério de Confirmação)")
                logger.info("Nenhum dado encontrado para o Piauí (Critério de Confirmação).")
                return
            
            base_df = pd.DataFrame(list(MUNICIPIOS_PIAUI.items()), columns=['ID_MN_RESI', 'Nome_Municipio'])
            
            df_mapped = df_filtered.copy()
            criterio_map = {
                1: 'Laboratorial',
                2: 'Clínico Epidemiológico',
                3: 'Em Investigação',
                4: 'Desconhecido'
            }
            df_mapped['CRITERIO'] = pd.to_numeric(df_mapped['CRITERIO'], errors='coerce').fillna(4).astype(int)
            df_mapped['Critério de Confirmação'] = df_mapped['CRITERIO'].map(criterio_map).fillna('Desconhecido')
            
            df_mapped = df_mapped[df_mapped['Critério de Confirmação'].isin(selected_criterios)].copy()
            logger.info(f"Tamanho do DataFrame após filtro por critérios selecionados: {len(df_mapped)}")
            if df_mapped.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para os critérios selecionados.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set("Nenhum dado para os critérios selecionados")
                logger.info(f"Nenhum dado encontrado para os Critérios: {selected_criterios}.")
                return
            
            pivot_table = pd.pivot_table(
                df_mapped,
                index='ID_MN_RESI',
                columns='Critério de Confirmação',
                aggfunc='size',
                fill_value=0
            )
            
            for crit in selected_criterios:
                if crit not in pivot_table.columns:
                    pivot_table[crit] = 0
            
            pivot_table = pivot_table[selected_criterios]
            
            pivot_table = pivot_table.astype(int)
            
            total_casos = df_mapped.groupby('ID_MN_RESI').size().astype(int)
            pivot_table['Total de Casos'] = total_casos
            
            result_table = base_df.merge(pivot_table, on='ID_MN_RESI', how='left').fillna(0)
            
            numeric_cols = selected_criterios + ['Total de Casos']
            result_table[numeric_cols] = result_table[numeric_cols].astype(int)
            
            columns = ['Nome_Municipio'] + selected_criterios + ['Total de Casos']
            result_table = result_table[columns]
            
            result_table = result_table.sort_values(by='Nome_Municipio')
            
            totals = {'Nome_Municipio': 'Total'}
            for col in selected_criterios + ['Total de Casos']:
                totals[col] = result_table[col].sum()
            totals_df = pd.DataFrame([totals], columns=columns)
            
            result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            self.current_report_df = result_table
            self.current_report_type = "Critério de Confirmação"
            self.current_territorio = f"Casos por Município - {self.ano}"
            
            self.btn_salvar.configure(state='normal')
            
            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Critério de Confirmação)\n")
            self.txt_dados.insert(tk.END, f"Critérios Selecionados: {', '.join(selected_criterios)}\n\n")
            self.txt_dados.insert(tk.END, result_table.to_string(index=False))
            self.status_var.set(f"Relatório de {self.current_territorio} (Critério de Confirmação) gerado")
            logger.info("Relatório de Casos por Município (Critério de Confirmação) gerado com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
            self.status_var.set("Erro ao gerar relatório")

    def relatorio_municipio_sinais_clinicos(self, selected_sinais):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return
        
        try:
            self.status_var.set("Gerando relatório de Casos por Município (Sinais Clínicos)...")
            logger.info("Iniciando geração do relatório de Casos por Município (Sinais Clínicos).")
            
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
            
            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set("Nenhum dado para o Piauí (Sinais Clínicos)")
                logger.info("Nenhum dado encontrado para o Piauí (Sinais Clínicos).")
                return
            
            base_df = pd.DataFrame(list(MUNICIPIOS_PIAUI.items()), columns=['ID_MN_RESI', 'Nome_Municipio'])
            
            sinais_map = {
                'Febre': 'FEBRE',
                'Mialgia': 'MIALGIA',
                'Cefaleia': 'CEFALEIA',
                'Exantema': 'EXANTEMA',
                'Vômito': 'VOMITO',
                'Náuseas': 'NAUSEA',
                'Dor nas costas': 'DOR_COSTAS',
                'Conjuntivite': 'CONJUNTVIT',
                'Artrite': 'ARTRITE',
                'Artralgia intensa': 'ARTRALGIA',
                'Petéquias': 'PETEQUIA_N',
                'Leucopenia': 'LEUCOPENIA',
                'Prova do laço positiva': 'LACO',
                'Dor retroorbital': 'DOR_RETRO'
            }
            
            missing_fields = [sinal for sinal in selected_sinais if sinais_map[sinal] not in df_filtered.columns]
            if missing_fields:
                messagebox.showerror("Erro", f"Os seguintes campos não foram encontrados no arquivo DBF: {', '.join([sinais_map[sinal] for sinal in missing_fields])}")
                self.status_var.set("Erro: Campos de sinais clínicos ausentes no DBF")
                logger.error(f"Campos ausentes no DBF: {missing_fields}")
                return
            
            df_mapped = df_filtered.copy()
            result_data = {'ID_MN_RESI': df_mapped['ID_MN_RESI']}
            
            for sinal in selected_sinais:
                campo = sinais_map[sinal]
                df_mapped[campo] = pd.to_numeric(df_mapped[campo], errors='coerce').fillna(0).astype(int)
                result_data[sinal] = (df_mapped[campo] == 1).astype(int)
            
            df_results = pd.DataFrame(result_data)
            
            pivot_table = df_results.groupby('ID_MN_RESI').sum()
            
            result_table = base_df.merge(pivot_table, on='ID_MN_RESI', how='left').fillna(0)
            
            numeric_cols = selected_sinais
            result_table[numeric_cols] = result_table[numeric_cols].astype(int)
            
            columns = ['Nome_Municipio'] + selected_sinais
            result_table = result_table[columns]
            
            result_table = result_table.sort_values(by='Nome_Municipio')
            
            totals = {'Nome_Municipio': 'Total'}
            for col in selected_sinais:
                totals[col] = result_table[col].sum()
            totals_df = pd.DataFrame([totals], columns=columns)
            
            result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            self.current_report_df = result_table
            self.current_report_type = "Sinais Clínicos"
            self.current_territorio = f"Casos por Município - {self.ano}"
            
            self.btn_salvar.configure(state='normal')
            
            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Sinais Clínicos)\n")
            self.txt_dados.insert(tk.END, f"Sinais Clínicos Selecionados: {', '.join(selected_sinais)}\n\n")
            self.txt_dados.insert(tk.END, result_table.to_string(index=False))
            self.status_var.set(f"Relatório de {self.current_territorio} (Sinais Clínicos) gerado")
            logger.info("Relatório de Casos por Município (Sinais Clínicos) gerado com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
            self.status_var.set("Erro ao gerar relatório")

    def relatorio_municipio_doencas_preexistentes(self, selected_doencas):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return
        
        try:
            self.status_var.set("Gerando relatório de Casos por Município (Doenças Pré-existentes)...")
            logger.info("Iniciando geração do relatório de Casos por Município (Doenças Pré-existentes).")
            
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
            
            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set("Nenhum dado para o Piauí (Doenças Pré-existentes)")
                logger.info("Nenhum dado encontrado para o Piauí (Doenças Pré-existentes).")
                return
            
            base_df = pd.DataFrame(list(MUNICIPIOS_PIAUI.items()), columns=['ID_MN_RESI', 'Nome_Municipio'])
            
            doencas_map = {
                'Diabetes': 'DIABETES',
                'Doenças hematológicas': 'HEMATOLOG',
                'Hepatopatias': 'HEPATOPAT',
                'Doença renal crônica': 'RENAL',
                'Hipertensão arterial': 'HIPERTENSA',
                'Doença ácido-péptica': 'ACIDO_PEPT',
                'Doenças auto-imunes': 'AUTO_IMUNE'
            }
            
            missing_fields = [doenca for doenca in selected_doencas if doencas_map[doenca] not in df_filtered.columns]
            if missing_fields:
                messagebox.showerror("Erro", f"Os seguintes campos não foram encontrados no arquivo DBF: {', '.join([doencas_map[doenca] for doenca in missing_fields])}")
                self.status_var.set("Erro: Campos de doenças pré-existentes ausentes no DBF")
                logger.error(f"Campos ausentes no DBF: {missing_fields}")
                return
            
            df_mapped = df_filtered.copy()
            result_data = {'ID_MN_RESI': df_mapped['ID_MN_RESI']}
            
            for doenca in selected_doencas:
                campo = doencas_map[doenca]
                df_mapped[campo] = pd.to_numeric(df_mapped[campo], errors='coerce').fillna(0).astype(int)
                result_data[doenca] = (df_mapped[campo] == 1).astype(int)
            
            df_results = pd.DataFrame(result_data)
            
            pivot_table = df_results.groupby('ID_MN_RESI').sum()
            
            result_table = base_df.merge(pivot_table, on='ID_MN_RESI', how='left').fillna(0)
            
            numeric_cols = selected_doencas
            result_table[numeric_cols] = result_table[numeric_cols].astype(int)
            
            columns = ['Nome_Municipio'] + selected_doencas
            result_table = result_table[columns]
            
            result_table = result_table.sort_values(by='Nome_Municipio')
            
            totals = {'Nome_Municipio': 'Total'}
            for col in selected_doencas:
                totals[col] = result_table[col].sum()
            totals_df = pd.DataFrame([totals], columns=columns)
            
            result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            self.current_report_df = result_table
            self.current_report_type = "Doenças Pré-existentes"
            self.current_territorio = f"Casos por Município - {self.ano}"
            
            self.btn_salvar.configure(state='normal')
            
            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Doenças Pré-existentes)\n")
            self.txt_dados.insert(tk.END, f"Doenças Selecionadas: {', '.join(selected_doencas)}\n\n")
            self.txt_dados.insert(tk.END, result_table.to_string(index=False))
            self.status_var.set(f"Relatório de {self.current_territorio} (Doenças Pré-existentes) gerado")
            logger.info("Relatório de Casos por Município (Doenças Pré-existentes) gerado com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
            self.status_var.set("Erro ao gerar relatório")

    def relatorio_territorio_sorotipo(self, selected_territorio):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return
        
        try:
            self.status_var.set(f"Gerando relatório de {selected_territorio} (Sorotipo)...")
            logger.info(f"Iniciando geração do relatório de {selected_territorio} (Sorotipo).")
            
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
            
            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set(f"Nenhum dado para {selected_territorio} (Sorotipo)")
                logger.info(f"Nenhum dado encontrado para o Piauí (Sorotipo).")
                return
            
            df_dengue = df_filtered[df_filtered['ID_AGRAVO'] == 'A90'].copy()
            logger.info(f"Tamanho do DataFrame após filtro por Dengue (A90): {len(df_dengue)}")
            
            df_dengue['Nome_Municipio'] = df_dengue['ID_MN_RESI'].map(MUNICIPIOS_PIAUI).fillna('Desconhecido')
            
            df_dengue['SOROTIPO'] = df_dengue['SOROTIPO'].astype(str).replace('', '0').replace('nan', '0')
            
            if selected_territorio == "Todos":
                df_dengue['Território'] = df_dengue['Nome_Municipio'].apply(lambda x: map_municipio_to_territorio(x, territorios_municipios))
                
                pivot_table = pd.pivot_table(
                    df_dengue,
                    index='Território',
                    columns='SOROTIPO',
                    aggfunc='size',
                    fill_value=0
                )
                
                sorotipo_map = {'1': 'DEN 1', '2': 'DEN 2', '3': 'DEN 3', '4': 'DEN 4'}
                for sorotipo in ['1', '2', '3', '4']:
                    if sorotipo not in pivot_table.columns:
                        pivot_table[sorotipo] = 0
                
                pivot_table.columns = [sorotipo_map.get(col, f'Sorotipo {col}') for col in pivot_table.columns]
                
                sorotipo_cols = ['DEN 1', 'DEN 2', 'DEN 3', 'DEN 4']
                pivot_table = pivot_table[[col for col in sorotipo_cols if col in pivot_table.columns]]
                
                for col in sorotipo_cols:
                    if col not in pivot_table.columns:
                        pivot_table[col] = 0
                
                pivot_table = pivot_table.astype(int)
                
                result_table = pivot_table.reset_index()
                
                columns = ['Território', 'DEN 1', 'DEN 2', 'DEN 3', 'DEN 4']
                result_table = result_table[columns]
                
                result_table = result_table.sort_values(by='Território')
                
                totals = {
                    'Território': 'Total',
                    'DEN 1': result_table['DEN 1'].sum(),
                    'DEN 2': result_table['DEN 2'].sum(),
                    'DEN 3': result_table['DEN 3'].sum(),
                    'DEN 4': result_table['DEN 4'].sum(),
                }
                totals_df = pd.DataFrame([totals], columns=columns)
                
                result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            else:
                municipios_territorio = territorios_municipios.get(selected_territorio, [])
                logger.info(f"Municípios do território {selected_territorio}: {municipios_territorio}")
                base_df = pd.DataFrame({'Nome_Municipio': municipios_territorio})
                
                df_territorio = df_dengue[df_dengue['Nome_Municipio'].isin(municipios_territorio)].copy()
                logger.info(f"Tamanho do DataFrame após filtro por território: {len(df_territorio)}")
                
                if df_territorio.empty:
                    self.txt_dados.delete(1.0, tk.END)
                    self.txt_dados.insert(tk.END, f"Nenhum dado encontrado para o território {selected_territorio}.\n")
                    self.btn_salvar.configure(state='disabled')
                    self.status_var.set(f"Nenhum dado para {selected_territorio} (Sorotipo)")
                    logger.info(f"Nenhum dado encontrado para o território {selected_territorio} (Sorotipo).")
                    return
                
                pivot_table = pd.pivot_table(
                    df_territorio,
                    index='Nome_Municipio',
                    columns='SOROTIPO',
                    aggfunc='size',
                    fill_value=0
                )
                
                sorotipo_map = {'1': 'DEN 1', '2': 'DEN 2', '3': 'DEN 3', '4': 'DEN 4'}
                for sorotipo in ['1', '2', '3', '4']:
                    if sorotipo not in pivot_table.columns:
                        pivot_table[sorotipo] = 0
                
                pivot_table.columns = [sorotipo_map.get(col, f'Sorotipo {col}') for col in pivot_table.columns]
                
                sorotipo_cols = ['DEN 1', 'DEN 2', 'DEN 3', 'DEN 4']
                pivot_table = pivot_table[[col for col in sorotipo_cols if col in pivot_table.columns]]
                
                for col in sorotipo_cols:
                    if col not in pivot_table.columns:
                        pivot_table[col] = 0
                
                pivot_table = pivot_table.astype(int)
                
                result_table = pivot_table.reset_index()
                result_table = base_df.merge(result_table, on='Nome_Municipio', how='left').fillna(0)
                
                numeric_cols = ['DEN 1', 'DEN 2', 'DEN 3', 'DEN 4']
                result_table[numeric_cols] = result_table[numeric_cols].astype(int)
                
                columns = ['Nome_Municipio', 'DEN 1', 'DEN 2', 'DEN 3', 'DEN 4']
                result_table = result_table[columns]
                
                result_table = result_table.sort_values(by='Nome_Municipio')
                
                totals = {
                    'Nome_Municipio': 'Total',
                    'DEN 1': result_table['DEN 1'].sum(),
                    'DEN 2': result_table['DEN 2'].sum(),
                    'DEN 3': result_table['DEN 3'].sum(),
                    'DEN 4': result_table['DEN 4'].sum(),
                }
                totals_df = pd.DataFrame([totals], columns=columns)
                
                result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            self.current_report_df = result_table
            self.current_report_type = "Sorotipo"
            self.current_territorio = selected_territorio
            
            self.btn_salvar.configure(state='normal')
            
            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Sorotipo) - {self.ano}\n\n")
            self.txt_dados.insert(tk.END, result_table.to_string(index=False))
            self.status_var.set(f"Relatório de {self.current_territorio} (Sorotipo) gerado")
            logger.info(f"Relatório de {self.current_territorio} (Sorotipo) gerado com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
            self.status_var.set("Erro ao gerar relatório")

    def relatorio_territorio_classificacao(self, selected_territorio):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return
        
        try:
            self.status_var.set(f"Gerando relatório de {selected_territorio} (Classificação)...")
            logger.info(f"Iniciando geração do relatório de {selected_territorio} (Classificação).")
            
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
            
            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set(f"Nenhum dado para {selected_territorio} (Classificação)")
                logger.info(f"Nenhum dado encontrado para o Piauí (Classificação).")
                return
            
            df_mapped = df_filtered.copy()
            df_mapped['CLASSI_FIN'] = pd.to_numeric(df_mapped['CLASSI_FIN'], errors='coerce').fillna(0).astype(int)
            logger.info(f"Valores únicos de CLASSI_FIN após mapeamento: {df_mapped['CLASSI_FIN'].unique()}")
            
            classi_fin_map = {
                10: 'Dengue', 11: 'Dengue com Sinais de Alarme', 12: 'Dengue Grave',
                5: 'Descartado', 8: 'Inconclusivo', 0: 'Em Branco'
            }
            df_mapped['Classificação'] = df_mapped['CLASSI_FIN'].map(classi_fin_map).fillna('Em Branco')
            
            df_mapped['Nome_Municipio'] = df_mapped['ID_MN_RESI'].map(MUNICIPIOS_PIAUI).fillna('Desconhecido')
            
            if selected_territorio == "Todos":
                df_mapped['Território'] = df_mapped['Nome_Municipio'].apply(lambda x: map_municipio_to_territorio(x, territorios_municipios))
                
                pivot_table = pd.pivot_table(
                    df_mapped,
                    index='Território',
                    columns='Classificação',
                    aggfunc='size',
                    fill_value=0
                )
                
                expected_classifications = ['Dengue', 'Dengue com Sinais de Alarme', 'Dengue Grave', 'Descartado', 'Inconclusivo', 'Em Branco']
                for classi in expected_classifications:
                    if classi not in pivot_table.columns:
                        pivot_table[classi] = 0
                
                pivot_table = pivot_table[expected_classifications]
                
                pivot_table = pivot_table.astype(int)
                
                result_table = pivot_table.reset_index()
                
                columns = ['Território'] + expected_classifications
                result_table = result_table[columns]
                
                result_table = result_table.sort_values(by='Território')
                
                totals = {'Território': 'Total'}
                for col in expected_classifications:
                    totals[col] = result_table[col].sum()
                totals_df = pd.DataFrame([totals], columns=columns)
                
                result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            else:
                municipios_territorio = territorios_municipios.get(selected_territorio, [])
                logger.info(f"Municípios do território {selected_territorio}: {municipios_territorio}")
                base_df = pd.DataFrame({'Nome_Municipio': municipios_territorio})
                
                df_territorio = df_mapped[df_mapped['Nome_Municipio'].isin(municipios_territorio)].copy()
                logger.info(f"Tamanho do DataFrame após filtro por território: {len(df_territorio)}")
                
                if df_territorio.empty:
                    self.txt_dados.delete(1.0, tk.END)
                    self.txt_dados.insert(tk.END, f"Nenhum dado encontrado para o território {selected_territorio}.\n")
                    self.btn_salvar.configure(state='disabled')
                    self.status_var.set(f"Nenhum dado para {selected_territorio} (Classificação)")
                    logger.info(f"Nenhum dado encontrado para o território {selected_territorio} (Classificação).")
                    return
                
                pivot_table = pd.pivot_table(
                    df_territorio,
                    index='Nome_Municipio',
                    columns='Classificação',
                    aggfunc='size',
                    fill_value=0
                )
                
                expected_classifications = ['Dengue', 'Dengue com Sinais de Alarme', 'Dengue Grave', 'Descartado', 'Inconclusivo', 'Em Branco']
                for classi in expected_classifications:
                    if classi not in pivot_table.columns:
                        pivot_table[classi] = 0
                
                pivot_table = pivot_table[expected_classifications]
                
                pivot_table = pivot_table.astype(int)
                
                result_table = pivot_table.reset_index()
                result_table = base_df.merge(result_table, on='Nome_Municipio', how='left').fillna(0)
                
                numeric_cols = expected_classifications
                result_table[numeric_cols] = result_table[numeric_cols].astype(int)
                
                columns = ['Nome_Municipio'] + expected_classifications
                result_table = result_table[columns]
                
                result_table = result_table.sort_values(by='Nome_Municipio')
                
                totals = {'Nome_Municipio': 'Total'}
                for col in expected_classifications:
                    totals[col] = result_table[col].sum()
                totals_df = pd.DataFrame([totals], columns=columns)
                
                result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            self.current_report_df = result_table
            self.current_report_type = "Classificação"
            self.current_territorio = selected_territorio
            
            self.btn_salvar.configure(state='normal')
            
            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Classificação) - {self.ano}\n\n")
            self.txt_dados.insert(tk.END, result_table.to_string(index=False))
            self.status_var.set(f"Relatório de {self.current_territorio} (Classificação) gerado")
            logger.info(f"Relatório de {self.current_territorio} (Classificação) gerado com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
            self.status_var.set("Erro ao gerar relatório")

    def relatorio_territorio_evolucao(self, selected_territorio):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return
        
        try:
            self.status_var.set(f"Gerando relatório de {selected_territorio} (Evolução do Caso)...")
            logger.info(f"Iniciando geração do relatório de {selected_territorio} (Evolução do Caso).")
            
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
            
            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set(f"Nenhum dado para {selected_territorio} (Evolução do Caso)")
                logger.info(f"Nenhum dado encontrado para o Piauí (Evolução do Caso).")
                return
            
            df_mapped = df_filtered.copy()
            df_mapped['EVOLUCAO'] = pd.to_numeric(df_mapped['EVOLUCAO'], errors='coerce').fillna(0).astype(int)
            logger.info(f"Valores únicos de EVOLUCAO após mapeamento: {df_mapped['EVOLUCAO'].unique()}")
            
            evolucao_map = {
                1: 'Cura', 2: 'Óbito por Dengue', 3: 'Óbito por Outras Causas',
                4: 'Óbito em Investigação', 9: 'Ignorado', 0: 'Em Branco'
            }
            df_mapped['Evolução'] = df_mapped['EVOLUCAO'].map(evolucao_map).fillna('Em Branco')
            
            df_mapped['Nome_Municipio'] = df_mapped['ID_MN_RESI'].map(MUNICIPIOS_PIAUI).fillna('Desconhecido')
            
            if selected_territorio == "Todos":
                df_mapped['Território'] = df_mapped['Nome_Municipio'].apply(lambda x: map_municipio_to_territorio(x, territorios_municipios))
                
                pivot_table = pd.pivot_table(
                    df_mapped,
                    index='Território',
                    columns='Evolução',
                    aggfunc='size',
                    fill_value=0
                )
                
                expected_evolutions = ['Cura', 'Óbito por Dengue', 'Óbito por Outras Causas', 'Óbito em Investigação', 'Ignorado', 'Em Branco']
                for evol in expected_evolutions:
                    if evol not in pivot_table.columns:
                        pivot_table[evol] = 0
                
                pivot_table = pivot_table[expected_evolutions]
                
                pivot_table = pivot_table.astype(int)
                
                result_table = pivot_table.reset_index()
                
                columns = ['Território'] + expected_evolutions
                result_table = result_table[columns]
                
                result_table = result_table.sort_values(by='Território')
                
                totals = {'Território': 'Total'}
                for col in expected_evolutions:
                    totals[col] = result_table[col].sum()
                totals_df = pd.DataFrame([totals], columns=columns)
                
                result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            else:
                municipios_territorio = territorios_municipios.get(selected_territorio, [])
                logger.info(f"Municípios do território {selected_territorio}: {municipios_territorio}")
                base_df = pd.DataFrame({'Nome_Municipio': municipios_territorio})
                
                df_territorio = df_mapped[df_mapped['Nome_Municipio'].isin(municipios_territorio)].copy()
                logger.info(f"Tamanho do DataFrame após filtro por território: {len(df_territorio)}")
                
                if df_territorio.empty:
                    self.txt_dados.delete(1.0, tk.END)
                    self.txt_dados.insert(tk.END, f"Nenhum dado encontrado para o território {selected_territorio}.\n")
                    self.btn_salvar.configure(state='disabled')
                    self.status_var.set(f"Nenhum dado para {selected_territorio} (Evolução do Caso)")
                    logger.info(f"Nenhum dado encontrado para o território {selected_territorio} (Evolução do Caso).")
                    return
                
                pivot_table = pd.pivot_table(
                    df_territorio,
                    index='Nome_Municipio',
                    columns='Evolução',
                    aggfunc='size',
                    fill_value=0
                )
                
                expected_evolutions = ['Cura', 'Óbito por Dengue', 'Óbito por Outras Causas', 'Óbito em Investigação', 'Ignorado', 'Em Branco']
                for evol in expected_evolutions:
                    if evol not in pivot_table.columns:
                        pivot_table[evol] = 0
                
                pivot_table = pivot_table[expected_evolutions]
                
                pivot_table = pivot_table.astype(int)
                
                result_table = pivot_table.reset_index()
                result_table = base_df.merge(result_table, on='Nome_Municipio', how='left').fillna(0)
                
                numeric_cols = expected_evolutions
                result_table[numeric_cols] = result_table[numeric_cols].astype(int)
                
                columns = ['Nome_Municipio'] + expected_evolutions
                result_table = result_table[columns]
                
                result_table = result_table.sort_values(by='Nome_Municipio')
                
                totals = {'Nome_Municipio': 'Total'}
                for col in expected_evolutions:
                    totals[col] = result_table[col].sum()
                totals_df = pd.DataFrame([totals], columns=columns)
                
                result_table = pd.concat([result_table, totals_df], ignore_index=True)
            
            self.current_report_df = result_table
            self.current_report_type = "Evolução do Caso"
            self.current_territorio = selected_territorio
            
            self.btn_salvar.configure(state='normal')
            
            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Evolução do Caso) - {self.ano}\n\n")
            self.txt_dados.insert(tk.END, result_table.to_string(index=False))
            self.status_var.set(f"Relatório de {self.current_territorio} (Evolução do Caso) gerado")
            logger.info(f"Relatório de {self.current_territorio} (Evolução do Caso) gerado com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
            self.status_var.set("Erro ao gerar relatório")

    def relatorio_idade_sexo(self, selected_territorio):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return

        selection_window = tk.Toplevel(self.root)
        selection_window.title("Selecionar Classificações - Idade e Sexo")
        selection_window.geometry("400x300")
        selection_window.transient(self.root)
        selection_window.grab_set()

        ttk.Label(selection_window, text="Selecione as Classificações:", font=('Helvetica', 10)).pack(pady=5)

        classificacoes = ['Dengue', 'Dengue com Sinais de Alarme', 'Dengue Grave', 'Descartado', 'Inconclusivo', 'Em Branco']
        classificacao_vars = {classi: tk.BooleanVar(value=True) for classi in classificacoes}

        check_frame = ttk.Frame(selection_window)
        check_frame.pack(pady=5, fill=tk.BOTH, expand=True)

        for classi, var in classificacao_vars.items():
            ttk.Checkbutton(check_frame, text=classi, variable=var).pack(anchor='w', padx=10)

        def gerar_relatorio():
            selected_classificacoes = [classi for classi, var in classificacao_vars.items() if var.get()]
            if not selected_classificacoes:
                messagebox.showwarning("Aviso", "Selecione pelo menos uma classificação.")
                return
            selection_window.destroy()
            try:
                self.status_var.set(f"Gerando relatório de {selected_territorio} (Idade e Sexo)...")
                logger.info(f"Iniciando geração do relatório de {selected_territorio} (Idade e Sexo) com Classificações: {selected_classificacoes}.")
                
                self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
                self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
                
                df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
                logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
                if df_filtered.empty:
                    self.txt_dados.delete(1.0, tk.END)
                    self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                    self.btn_salvar.configure(state='disabled')
                    self.status_var.set(f"Nenhum dado para {selected_territorio} (Idade e Sexo)")
                    logger.info(f"Nenhum dado encontrado para o Piauí (Idade e Sexo).")
                    return
                
                df_mapped = df_filtered.copy()
                df_mapped['Nome_Municipio'] = df_mapped['ID_MN_RESI'].map(MUNICIPIOS_PIAUI).fillna('Desconhecido')
                
                if selected_territorio != "Todos":
                    municipios_territorio = territorios_municipios.get(selected_territorio, [])
                    logger.info(f"Municípios do território {selected_territorio}: {municipios_territorio}")
                    df_mapped = df_mapped[df_mapped['Nome_Municipio'].isin(municipios_territorio)].copy()
                    logger.info(f"Tamanho do DataFrame após filtro por território: {len(df_mapped)}")
                    if df_mapped.empty:
                        self.txt_dados.delete(1.0, tk.END)
                        self.txt_dados.insert(tk.END, f"Nenhum dado encontrado para o território {selected_territorio}.\n")
                        self.btn_salvar.configure(state='disabled')
                        self.status_var.set(f"Nenhum dado para {selected_territorio} (Idade e Sexo)")
                        logger.info(f"Nenhum dado encontrado para o território {selected_territorio} (Idade e Sexo).")
                        return
                
                df_mapped['CLASSI_FIN'] = pd.to_numeric(df_mapped['CLASSI_FIN'], errors='coerce').fillna(0).astype(int)
                logger.info(f"Valores únicos de CLASSI_FIN após mapeamento: {df_mapped['CLASSI_FIN'].unique()}")
                classi_fin_map = {
                    10: 'Dengue', 11: 'Dengue com Sinais de Alarme', 12: 'Dengue Grave',
                    5: 'Descartado', 8: 'Inconclusivo', 0: 'Em Branco'
                }
                df_mapped['Classificação'] = df_mapped['CLASSI_FIN'].map(classi_fin_map).fillna('Em Branco')
                
                df_mapped = df_mapped[df_mapped['Classificação'].isin(selected_classificacoes)].copy()
                logger.info(f"Tamanho do DataFrame após filtro por classificações selecionadas: {len(df_mapped)}")
                if df_mapped.empty:
                    self.txt_dados.delete(1.0, tk.END)
                    self.txt_dados.insert(tk.END, f"Nenhum dado encontrado para as classificações selecionadas em {selected_territorio}.\n")
                    self.btn_salvar.configure(state='disabled')
                    self.status_var.set(f"Nenhum dado para {selected_territorio} com as classificações selecionadas")
                    logger.info(f"Nenhum dado encontrado para {selected_territorio} com Classificações: {selected_classificacoes}.")
                    return
                
                df_mapped['Idade'] = df_mapped['NU_IDADE_N'].astype(str).str[1:].astype(int, errors='ignore')
                df_mapped['Idade'] = pd.to_numeric(df_mapped['Idade'], errors='coerce').fillna(-1).astype(int)
                
                bins = [0, 4, 9, 14, 19, 24, 29, 34, 39, 44, 49, 54, 59, 64, 69, 74, 79, 120]
                labels = ['0-4', '5-9', '10-14', '15-19', '20-24', '25-29', '30-34', '35-39', 
                          '40-44', '45-49', '50-54', '55-59', '60-64', '65-69', '70-74', '75-79', '80+']
                df_mapped['Faixa Etária'] = pd.cut(df_mapped['Idade'], bins=bins, labels=labels, include_lowest=True, right=False)
                df_mapped['Faixa Etária'] = df_mapped['Faixa Etária'].cat.add_categories(['Desconhecido']).fillna('Desconhecido')
                
                sexo_map = {'M': 'Masculino', 'F': 'Feminino', 'I': 'Ignorado'}
                df_mapped['Sexo'] = df_mapped['CS_SEXO'].map(sexo_map).fillna('Ignorado')
                
                pivot_table = pd.pivot_table(
                    df_mapped,
                    index='Faixa Etária',
                    columns='Sexo',
                    aggfunc='size',
                    fill_value=0
                )
                
                expected_sexos = ['Masculino', 'Feminino', 'Ignorado']
                for sexo in expected_sexos:
                    if sexo not in pivot_table.columns:
                        pivot_table[sexo] = 0
                
                pivot_table = pivot_table[expected_sexos]
                
                pivot_table = pivot_table.astype(int)
                
                pivot_table['Total'] = pivot_table.sum(axis=1)
                
                result_table = pivot_table.reset_index()
                
                totals = {'Faixa Etária': 'Total'}
                for col in expected_sexos:
                    totals[col] = result_table[col].sum()
                totals['Total'] = result_table['Total'].sum()
                totals_df = pd.DataFrame([totals], columns=['Faixa Etária'] + expected_sexos + ['Total'])
                
                result_table = pd.concat([result_table, totals_df], ignore_index=True)
                
                self.current_report_df = result_table
                self.current_report_type = "Idade e Sexo"
                self.current_territorio = selected_territorio
                
                self.btn_salvar.configure(state='normal')
                
                fig, ax = plt.subplots(figsize=(12, 6))
                bar_width = 0.35
                index = range(len(result_table) - 1)
                masculino = result_table['Masculino'].iloc[:-1]
                feminino = result_table['Feminino'].iloc[:-1]
                
                ax.bar([i - bar_width/2 for i in index], masculino, bar_width, label='Masculino', color='skyblue', edgecolor='black')
                ax.bar([i + bar_width/2 for i in index], feminino, bar_width, label='Feminino', color='lightcoral', edgecolor='black')
                
                ax.set_xlabel('Faixa Etária', fontsize=10)
                ax.set_ylabel('Número de Casos', fontsize=10)
                ax.set_title(f'Distribuição por Idade e Sexo em {selected_territorio} - {self.ano}', fontsize=12)
                ax.set_xticks(index)
                ax.set_xticklabels(result_table['Faixa Etária'].iloc[:-1], rotation=45, ha='right', fontsize=8)
                ax.legend(fontsize=8)
                ax.grid(True, which="both", ls="--", alpha=0.7, axis='y')
                plt.tight_layout()
                
                self.graph_buffer = io.BytesIO()
                plt.savefig(self.graph_buffer, format='png', bbox_inches='tight', dpi=150)
                plt.close(fig)
                self.graph_buffer.seek(0)
                
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Idade e Sexo) - {self.ano}\n")
                self.txt_dados.insert(tk.END, f"Classificações Selecionadas: {', '.join(selected_classificacoes)}\n\n")
                self.txt_dados.insert(tk.END, result_table.to_string(index=False))
                self.status_var.set(f"Relatório de {self.current_territorio} (Idade e Sexo) gerado")
                logger.info(f"Relatório de {self.current_territorio} (Idade e Sexo) gerado com sucesso.")
            except Exception as e:
                logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
                messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
                self.status_var.set("Erro ao gerar relatório")
                if hasattr(self, 'graph_buffer') and not self.graph_buffer.closed:
                    self.graph_buffer.close()

        ttk.Button(selection_window, text="Gerar Relatório", command=gerar_relatorio).pack(pady=10)

    def relatorio_territorio_georreferenciamento(self, selected_territorio, selected_categoria, selected_subitem):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return

        try:
            import geobr
        except ImportError:
            messagebox.showerror("Erro", "O pacote 'geobr' não está instalado. Instale com 'pip install geobr'.")
            self.status_var.set("Erro: geobr não instalado")
            logger.error("Pacote 'geobr' não instalado.")
            return

        try:
            self.status_var.set(f"Gerando relatório de {selected_territorio} (Georreferenciamento - {selected_categoria}: {selected_subitem})...")
            logger.info(f"Iniciando geração do relatório de {selected_territorio} (Georreferenciamento - {selected_categoria}: {selected_subitem}).")

            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')
            logger.info(f"Valores únicos de ID_MN_RESI no DataFrame original: {self.df['ID_MN_RESI'].unique()}")

            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            logger.info(f"Valores únicos de ID_MN_RESI após filtro do Piauí: {df_filtered['ID_MN_RESI'].unique()}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set(f"Nenhum dado para {selected_territorio} (Georreferenciamento)")
                logger.info(f"Nenhum dado encontrado para o Piauí (Georreferenciamento).")
                return

            df_mapped = df_filtered.copy()
            df_mapped['Nome_Municipio'] = df_mapped['ID_MN_RESI'].map(MUNICIPIOS_PIAUI).fillna('Desconhecido')
            logger.info(f"Municípios mapeados: {df_mapped['Nome_Municipio'].unique()}")

            if selected_territorio != "Todos":
                municipios_territorio = territorios_municipios.get(selected_territorio, [])
                logger.info(f"Municípios do território {selected_territorio}: {municipios_territorio}")
                df_mapped = df_mapped[df_mapped['Nome_Municipio'].isin(municipios_territorio)].copy()
                logger.info(f"Tamanho do DataFrame após filtro por território: {len(df_mapped)}")
                if df_mapped.empty:
                    self.txt_dados.delete(1.0, tk.END)
                    self.txt_dados.insert(tk.END, f"Nenhum dado encontrado para o território {selected_territorio}.\n")
                    self.btn_salvar.configure(state='disabled')
                    self.status_var.set(f"Nenhum dado para {selected_territorio} (Georreferenciamento)")
                    logger.info(f"Nenhum dado encontrado para o território {selected_territorio} (Georreferenciamento).")
                    return

            if selected_categoria == "Sorotipo":
                df_mapped = df_mapped[df_mapped['ID_AGRAVO'] == 'A90'].copy()
                logger.info(f"Tamanho do DataFrame após filtro por Dengue (A90): {len(df_mapped)}")
                df_mapped['SOROTIPO'] = df_mapped['SOROTIPO'].astype(str).replace('', '0').replace('nan', '0')
                sorotipo_map = {'1': 'DEN 1', '2': 'DEN 2', '3': 'DEN 3', '4': 'DEN 4', '0': 'Desconhecido'}
                df_mapped['Sorotipo'] = df_mapped['SOROTIPO'].map(sorotipo_map).fillna('Desconhecido')
                df_mapped = df_mapped[df_mapped['Sorotipo'] == selected_subitem].copy()
                logger.info(f"Tamanho do DataFrame após filtro por Sorotipo ({selected_subitem}): {len(df_mapped)}")
                pivot_table = pd.pivot_table(
                    df_mapped,
                    index='ID_MN_RESI',
                    aggfunc='size',
                    fill_value=0
                ).reset_index(name='Casos')

            elif selected_categoria == "Classificação":
                df_mapped['CLASSI_FIN'] = pd.to_numeric(df_mapped['CLASSI_FIN'], errors='coerce').fillna(0).astype(int)
                classi_fin_map = {
                    10: 'Dengue', 11: 'Dengue com Sinais de Alarme', 12: 'Dengue Grave',
                    5: 'Descartado', 8: 'Inconclusivo', 0: 'Em Branco'
                }
                df_mapped['Classificação'] = df_mapped['CLASSI_FIN'].map(classi_fin_map).fillna('Em Branco')
                logger.info(f"Valores únicos de Classificação após mapeamento: {df_mapped['Classificação'].unique()}")
                df_mapped = df_mapped[df_mapped['Classificação'] == selected_subitem].copy()
                logger.info(f"Tamanho do DataFrame após filtro por Classificação ({selected_subitem}): {len(df_mapped)}")
                pivot_table = pd.pivot_table(
                    df_mapped,
                    index='ID_MN_RESI',
                    aggfunc='size',
                    fill_value=0
                ).reset_index(name='Casos')

            else:
                df_mapped['EVOLUCAO'] = pd.to_numeric(df_mapped['EVOLUCAO'], errors='coerce').fillna(0).astype(int)
                evolucao_map = {
                    1: 'Cura', 2: 'Óbito por Dengue', 3: 'Óbito por Outras Causas',
                    4: 'Óbito em Investigação', 9: 'Ignorado', 0: 'Em Branco'
                }
                df_mapped['Evolução'] = df_mapped['EVOLUCAO'].map(evolucao_map).fillna('Em Branco')
                df_mapped = df_mapped[df_mapped['Evolução'] == selected_subitem].copy()
                logger.info(f"Tamanho do DataFrame após filtro por Evolução ({selected_subitem}): {len(df_mapped)}")
                pivot_table = pd.pivot_table(
                    df_mapped,
                    index='ID_MN_RESI',
                    aggfunc='size',
                    fill_value=0
                ).reset_index(name='Casos')

            base_df = pd.DataFrame(list(MUNICIPIOS_PIAUI.items()), columns=['ID_MN_RESI', 'Nome_Municipio'])
            result_table = base_df.merge(pivot_table, on='ID_MN_RESI', how='left').fillna(0)
            result_table['Casos'] = result_table['Casos'].astype(int)
            logger.info(f"Tabela de resultados antes do filtro por território: {result_table[['Nome_Municipio', 'ID_MN_RESI', 'Casos']].to_string()}")

            if selected_territorio != "Todos":
                municipios_territorio = territorios_municipios.get(selected_territorio, [])
                result_table = result_table[result_table['Nome_Municipio'].isin(municipios_territorio)].copy()
                logger.info(f"Tabela de resultados após filtro por território: {result_table[['Nome_Municipio', 'ID_MN_RESI', 'Casos']].to_string()}")

            total_casos = result_table['Casos'].sum()
            totals = {
                'Nome_Municipio': 'Total',
                'ID_MN_RESI': '',
                'Casos': total_casos
            }
            totals_df = pd.DataFrame([totals], columns=['Nome_Municipio', 'ID_MN_RESI', 'Casos'])
            result_table = pd.concat([result_table, totals_df], ignore_index=True)
            logger.info(f"Total de casos calculado: {total_casos}")

            self.current_report_df = result_table[['Nome_Municipio', 'Casos']]
            self.current_report_type = "Georreferenciamento"
            self.current_territorio = selected_territorio

            # Load geographical data
            municipios_pi = geobr.read_municipality(code_muni='PI', year=2022)
            municipios_pi['code_muni'] = municipios_pi['code_muni'].astype(str).str[:-1].str.zfill(6)
            logger.info(f"Códigos de município do geobr antes do merge: {municipios_pi['code_muni'].unique()}")
            logger.info(f"Códigos de município no result_table antes do merge: {result_table['ID_MN_RESI'].unique()}")

            # Normalize names for comparison and territory filtering
            municipios_pi['name_muni_normalized'] = municipios_pi['name_muni'].apply(normalize_name)
            result_table['Nome_Municipio_normalized'] = result_table['Nome_Municipio'].apply(normalize_name)

            if selected_territorio != "Todos":
                municipios_territorio = territorios_municipios.get(selected_territorio, [])
                municipios_territorio_normalized = [normalize_name(m) for m in municipios_territorio]
                logger.info(f"Municípios normalizados do território {selected_territorio}: {municipios_territorio_normalized}")
                municipios_pi = municipios_pi[municipios_pi['name_muni_normalized'].isin(municipios_territorio_normalized)].copy()
                logger.info(f"Municípios filtrados para {selected_territorio} antes do merge: {municipios_pi['name_muni'].tolist()}")

            # Log the municipalities and codes in both DataFrames before the merge
            logger.info(f"municipios_pi antes do merge:\n{municipios_pi[['name_muni', 'code_muni']].to_string()}")
            logger.info(f"result_table antes do merge:\n{result_table[['Nome_Municipio', 'ID_MN_RESI', 'Casos']].to_string()}")

            # First attempt: Merge on code_muni and ID_MN_RESI
            merged_gdf = municipios_pi.merge(result_table, left_on='code_muni', right_on='ID_MN_RESI', how='left')
            merged_gdf['Casos'] = merged_gdf['Casos'].fillna(0).astype(int)
            logger.info(f"Soma de casos após merge com geobr (primeira tentativa): {merged_gdf['Casos'].sum()}")

            # Check for unmatched municipalities
            unmatched = result_table[~result_table['ID_MN_RESI'].isin(merged_gdf['code_muni'])]
            if not unmatched.empty:
                logger.warning(f"Municípios não encontrados em municipios_pi durante o merge por código:\n{unmatched[['Nome_Municipio', 'ID_MN_RESI', 'Casos']].to_string()}")
                # Fallback: Merge on normalized names if code-based merge fails
                logger.info("Tentando merge por nomes normalizados como fallback...")
                merged_gdf = municipios_pi.merge(result_table, left_on='name_muni_normalized', right_on='Nome_Municipio_normalized', how='left')
                merged_gdf['Casos'] = merged_gdf['Casos'].fillna(0).astype(int)
                logger.info(f"Soma de casos após merge com geobr (segunda tentativa por nome): {merged_gdf['Casos'].sum()}")

            logger.info(f"Após o merge, número de linhas no merged_gdf: {len(merged_gdf)}")
            logger.info(f"Municípios após o merge: {merged_gdf['name_muni'].tolist()}")
            logger.info(f"Casos após o merge: {merged_gdf['Casos'].tolist()}")

            # Verificar se o merged_gdf está vazio ou não contém geometrias válidas
            if merged_gdf.empty or merged_gdf['geometry'].is_empty.all():
                logger.error("O GeoDataFrame está vazio ou não contém geometrias válidas para plotagem.")
                messagebox.showerror("Erro", "Não foi possível plotar o mapa: GeoDataFrame vazio ou sem geometrias válidas.")
                self.status_var.set("Erro: GeoDataFrame vazio")
                # Still display the table of results even if the map cannot be plotted
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Georreferenciamento - {selected_categoria}: {selected_subitem}) - {self.ano}\n\n")
                self.txt_dados.insert(tk.END, result_table[['Nome_Municipio', 'Casos']].to_string(index=False))
                self.btn_salvar.configure(state='normal')
                return

            # Generate the map
            fig, ax = plt.subplots(figsize=(10, 8))
            norm = Normalize(vmin=0, vmax=merged_gdf['Casos'].max() if merged_gdf['Casos'].max() > 0 else 1)
            cmap = plt.cm.Reds
            merged_gdf.plot(column='Casos', ax=ax, cmap=cmap, norm=norm, legend=False, edgecolor='black', linewidth=0.5)
            
            # Add colorbar
            sm = ScalarMappable(cmap=cmap, norm=norm)
            cbar = plt.colorbar(sm, ax=ax, shrink=0.5)
            cbar.set_label('Número de Casos', fontsize=10)
            
            # Add title and labels
            ax.set_title(f'{selected_categoria}: {selected_subitem} em {selected_territorio} - {self.ano}', fontsize=12)
            ax.set_axis_off()
            
            # Add municipality names if the number of municipalities is small
            if len(merged_gdf) <= 50:
                for idx, row in merged_gdf.iterrows():
                    if row['Casos'] > 0:
                        centroid = row['geometry'].centroid
                        ax.annotate(
                            text=f"{row['name_muni']}\n({int(row['Casos'])})",
                            xy=(centroid.x, centroid.y),
                            xytext=(3, 3),
                            textcoords="offset points",
                            fontsize=6,
                            color='black',
                            ha='center',
                            va='center',
                            bbox=dict(facecolor='white', alpha=0.5, edgecolor='none', boxstyle='round,pad=0.2')
                        )

            plt.tight_layout()
            
            # Save map to a temporary file
            map_path = f"map_{selected_territorio}_{selected_categoria}_{selected_subitem}_{self.ano}.png"
            fig.savefig(map_path, dpi=150, bbox_inches='tight')
            logger.info(f"Mapa salvo em: {map_path}")
            self.current_map_path = map_path
            
            # Display the map in a new window
            self.show_map_window(fig, f"Mapa Georreferenciado: {selected_categoria} - {selected_subitem} ({selected_territorio})")
            
            self.btn_salvar.configure(state='normal')
            
            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Relatório: {self.current_territorio} (Georreferenciamento - {selected_categoria}: {selected_subitem}) - {self.ano}\n\n")
            self.txt_dados.insert(tk.END, result_table[['Nome_Municipio', 'Casos']].to_string(index=False))
            self.status_var.set(f"Relatório de {self.current_territorio} (Georreferenciamento) gerado")
            logger.info(f"Relatório de {self.current_territorio} (Georreferenciamento) gerado com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar relatório: {str(e)}")
            self.status_var.set("Erro ao gerar relatório")
            plt.close()

    def show_map_window(self, fig, title):
        try:
            # Create a new top-level window
            map_window = tk.Toplevel(self.root)
            map_window.title(title)
            map_window.transient(self.root)

            # Maximize the window by default
            map_window.state('zoomed')
            map_window.protocol("WM_DELETE_WINDOW", lambda: [plt.close(fig), map_window.destroy()])

            # Create a main frame for the window
            main_frame = ttk.Frame(map_window)
            main_frame.pack(fill=tk.BOTH, expand=True)

            # Create a paned window to split the left panel (buttons) and right panel (map)
            paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
            paned_window.pack(fill=tk.BOTH, expand=True)

            # Left panel for buttons
            button_frame = ttk.Frame(paned_window, width=150)
            paned_window.add(button_frame, weight=1)

            # Right panel for the map
            map_frame = ttk.Frame(paned_window)
            paned_window.add(map_frame, weight=4)

            # Add the title label, centered in the map frame
            title_label = ttk.Label(
                map_frame,
                text=title,
                font=('Helvetica', 14, 'bold'),
                foreground='#1e3a8a'
            )
            title_label.pack(pady=10)

            # Create a frame for the canvas
            canvas_frame = ttk.Frame(map_frame)
            canvas_frame.pack(fill=tk.BOTH, expand=True)

            # Embed the matplotlib figure using FigureCanvasTkAgg
            canvas = FigureCanvasTkAgg(fig, master=canvas_frame)
            canvas.draw()
            canvas_widget = canvas.get_tk_widget()
            canvas_widget.pack(fill=tk.BOTH, expand=True)

            # Add a toolbar for zooming and panning at the top of the map
            toolbar = NavigationToolbar2Tk(canvas, map_frame)
            toolbar.update()
            toolbar.pack(side=tk.TOP, fill=tk.X)

            # Style for buttons to make them more visible
            style = ttk.Style()
            style.configure('Custom.TButton', font=('Helvetica', 12), padding=10)

            # Save menu for PDF and PNG options
            save_menu = tk.Menu(button_frame, tearoff=0)
            save_menu.add_command(label="Salvar como PDF", command=lambda: self.salvar_pdf(figures=[fig]))
            save_menu.add_command(label="Salvar como PNG", command=lambda: self.salvar_png([fig]))

            # Add save button with dropdown menu
            save_button = ttk.Button(
                button_frame,
                text="Salvar",
                command=lambda: save_menu.post(save_button.winfo_rootx(), save_button.winfo_rooty() + save_button.winfo_height()),
                style='Custom.TButton'
            )
            save_button.pack(pady=10, padx=10, fill=tk.X)

            # Add maximize/restore button
            maximize_button = ttk.Button(
                button_frame,
                text="Maximizar",
                command=lambda: [map_window.state('zoomed'), maximize_button.pack_forget(), minimize_button.pack(pady=10, padx=10, fill=tk.X)],
                style='Custom.TButton'
            )
            # Initially hidden since the window starts maximized
            maximize_button.pack_forget()

            minimize_button = ttk.Button(
                button_frame,
                text="Minimizar",
                command=lambda: [map_window.state('normal'), minimize_button.pack_forget(), maximize_button.pack(pady=10, padx=10, fill=tk.X)],
                style='Custom.TButton'
            )
            minimize_button.pack(pady=10, padx=10, fill=tk.X)

            # Add close button
            close_button = ttk.Button(
                button_frame,
                text="Fechar",
                command=lambda: [plt.close(fig), map_window.destroy()],
                style='Custom.TButton'
            )
            close_button.pack(pady=10, padx=10, fill=tk.X)

            # Add signature at the bottom of the main frame
            signature_label = ttk.Label(
                main_frame,
                text=self.signature,
                font=('Helvetica', 8),
                foreground='#4b5563'
            )
            signature_label.pack(pady=5)

            # Make the canvas responsive to window resizing
            def on_resize(event):
                canvas.draw()
            canvas_widget.bind('<Configure>', on_resize)

            logger.info(f"Janela de mapa '{title}' exibida com sucesso.")
        except Exception as e:
            logger.error(f"Falha ao exibir janela de mapa: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao exibir o mapa: {str(e)}")
            self.status_var.set("Erro ao exibir mapa")
            plt.close(fig)

    def create_dashboard_window(self):
        dashboard_window = tk.Toplevel(self.root)
        dashboard_window.title("Dashboard de Casos")
        dashboard_window.transient(self.root)
        dashboard_window.grab_set()

        # Maximize the window by default
        dashboard_window.state('zoomed')

        # Main frame
        main_frame = ttk.Frame(dashboard_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Button frame at the top
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5, side=tk.TOP)

        # Style for buttons to make them more visible
        style = ttk.Style()
        style.configure('Custom.TButton', font=('Helvetica', 12), padding=10)

        # Filter frame below the button frame
        filter_frame = ttk.LabelFrame(main_frame, text="Filtros", padding=10)
        filter_frame.pack(fill=tk.X, pady=5)

        # Territory selection
        territory_label = ttk.Label(filter_frame, text="Território:")
        territory_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        territory_var = tk.StringVar()
        territory_options = ["Todos"] + sorted(territorios_municipios.keys())
        territory_combobox = ttk.Combobox(filter_frame, textvariable=territory_var, values=territory_options, state='readonly')
        territory_combobox.grid(row=0, column=1, padx=5, pady=5)
        territory_combobox.set("Todos")

        # Criterion selection
        criterion_label = ttk.Label(filter_frame, text="Critério:")
        criterion_label.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        criterion_var = tk.StringVar()
        criterion_combobox = ttk.Combobox(filter_frame, textvariable=criterion_var, values=["Classificação", "Evolução do Caso"], state='readonly')
        criterion_combobox.grid(row=0, column=3, padx=5, pady=5)
        criterion_combobox.set("Classificação")

        # Top N municipalities
        top_n_label = ttk.Label(filter_frame, text="Top N Municípios:")
        top_n_label.grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        top_n_var = tk.StringVar(value="5")
        top_n_entry = ttk.Entry(filter_frame, textvariable=top_n_var, width=5)
        top_n_entry.grid(row=0, column=5, padx=5, pady=5)

        # Classifications selection
        classifications_label = ttk.Label(filter_frame, text="Classificações:")
        classifications_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        classifications_frame = ttk.Frame(filter_frame)
        classifications_frame.grid(row=1, column=1, columnspan=5, padx=5, pady=5, sticky=tk.W)
        classifications = ['Dengue', 'Dengue com Sinais de Alarme', 'Dengue Grave', 'Descartado', 'Inconclusivo', 'Em Branco']
        classification_vars = {classi: tk.BooleanVar(value=True) for classi in classifications}
        for idx, (classi, var) in enumerate(classification_vars.items()):
            ttk.Checkbutton(classifications_frame, text=classi, variable=var).grid(row=idx // 3, column=idx % 3, padx=5, pady=2, sticky=tk.W)

        # Frame for charts and page navigation
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Page navigation frame
        nav_frame = ttk.Frame(content_frame)
        nav_frame.pack(fill=tk.X, pady=5)

        page_label = ttk.Label(nav_frame, text="Página:")
        page_label.pack(side=tk.LEFT, padx=5)

        page_var = tk.StringVar()
        page_combobox = ttk.Combobox(nav_frame, textvariable=page_var, state='readonly', width=5)
        page_combobox.pack(side=tk.LEFT, padx=5)

        # Frame for charts
        charts_frame = ttk.Frame(content_frame)
        charts_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Dictionary to store figures for each page
        self.dashboard_figs = {}
        self.current_page = None

        def update_charts(page):
            if page not in self.dashboard_figs:
                return
            # Clear previous charts
            for widget in charts_frame.winfo_children():
                widget.destroy()
            
            # Display charts for the selected page
            for idx, fig in enumerate(self.dashboard_figs[page]):
                fig_frame = ttk.Frame(charts_frame)
                fig_frame.grid(row=idx // 2, column=idx % 2, padx=10, pady=10, sticky='nsew')
                
                canvas = FigureCanvasTkAgg(fig, master=fig_frame)
                canvas.draw()
                canvas_widget = canvas.get_tk_widget()
                canvas_widget.pack(fill=tk.BOTH, expand=True)

                toolbar = NavigationToolbar2Tk(canvas, fig_frame)
                toolbar.update()
                toolbar.pack(side=tk.TOP, fill=tk.X)

                # Make the canvas responsive to window resizing
                def on_resize(event):
                    canvas.draw()
                canvas_widget.bind('<Configure>', on_resize)

            # Configure grid weights for responsive layout
            for i in range(2):
                charts_frame.grid_rowconfigure(i, weight=1)
                charts_frame.grid_columnconfigure(i, weight=1)

            self.current_page = page

        def generate_dashboard():
            selected_territory = territory_var.get()
            selected_criterion = criterion_var.get()
            try:
                top_n = int(top_n_var.get())
                if top_n <= 0:
                    raise ValueError("O número de municípios deve ser positivo.")
            except ValueError as e:
                messagebox.showwarning("Aviso", "Por favor, insira um número válido de municípios (positivo e inteiro).")
                return

            selected_classifications = [classi for classi, var in classification_vars.items() if var.get()]
            if not selected_classifications and selected_criterion == "Classificação":
                messagebox.showwarning("Aviso", "Selecione pelo menos uma classificação.")
                return

            # Generate the dashboard and organize figures into pages
            figs = self.generate_dashboard(selected_territory, selected_criterion, top_n, selected_classifications)

            if figs:
                # Organize figures into pages (up to 4 per page)
                self.dashboard_figs = {}
                charts_per_page = 4
                for page in range(len(figs) // charts_per_page + (1 if len(figs) % charts_per_page else 0)):
                    start_idx = page * charts_per_page
                    end_idx = min((page + 1) * charts_per_page, len(figs))
                    self.dashboard_figs[str(page + 1)] = figs[start_idx:end_idx]

                # Update the page combobox
                page_combobox['values'] = list(self.dashboard_figs.keys())
                if self.dashboard_figs:
                    page_var.set('1')
                    update_charts('1')

                self.current_report_type = "Dashboard"
                self.current_territorio = selected_territory
                self.status_var.set(f"Dashboard de {self.current_territorio} ({selected_criterion}) gerado")
            else:
                self.status_var.set("Erro ao gerar dashboard")

        # Save menu for PDF and PNG options
        save_menu = tk.Menu(button_frame, tearoff=0)
        save_menu.add_command(label="Salvar como PDF", command=lambda: self.salvar_pdf(figures=[fig for page_figs in self.dashboard_figs.values() for fig in page_figs]))
        save_menu.add_command(label="Salvar como PNG", command=lambda: self.salvar_png([fig for page_figs in self.dashboard_figs.values() for fig in page_figs]))

        # Add buttons with enhanced visibility at the top
        ttk.Button(
            button_frame,
            text="Gerar Dashboard",
            command=generate_dashboard,
            style='Custom.TButton'
        ).pack(side=tk.LEFT, padx=20)

        ttk.Button(
            button_frame,
            text="Atualizar",
            command=generate_dashboard,
            style='Custom.TButton'
        ).pack(side=tk.LEFT, padx=20)

        # Add maximize/restore button
        maximize_button = ttk.Button(
            button_frame,
            text="Maximizar",
            command=lambda: [dashboard_window.state('zoomed'), maximize_button.pack_forget(), minimize_button.pack(side=tk.LEFT, padx=20)],
            style='Custom.TButton'
        )
        # Initially hidden since the window starts maximized
        maximize_button.pack_forget()

        minimize_button = ttk.Button(
            button_frame,
            text="Minimizar",
            command=lambda: [dashboard_window.state('normal'), minimize_button.pack_forget(), maximize_button.pack(side=tk.LEFT, padx=20)],
            style='Custom.TButton'
        )
        minimize_button.pack(side=tk.LEFT, padx=20)

        ttk.Button(
            button_frame,
            text="Salvar",
            command=lambda: save_menu.post(button_frame.winfo_children()[-1].winfo_rootx(), button_frame.winfo_children()[-1].winfo_rooty() + button_frame.winfo_children()[-1].winfo_height()),
            style='Custom.TButton'
        ).pack(side=tk.LEFT, padx=20)

        ttk.Button(
            button_frame,
            text="Fechar",
            command=dashboard_window.destroy,
            style='Custom.TButton'
        ).pack(side=tk.LEFT, padx=20)

        # Add signature at the bottom
        signature_label = ttk.Label(
            main_frame,
            text=self.signature,
            font=('Helvetica', 8),
            foreground='#4b5563'
        )
        signature_label.pack(pady=5, side=tk.BOTTOM)

        # Bind page selection to update charts
        def on_page_select(event):
            page = page_var.get()
            update_charts(page)
        page_combobox.bind('<<ComboboxSelected>>', on_page_select)

    def generate_dashboard(self, selected_territorio, selected_criterio, top_n, selected_classifications):
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo DBF carregado.")
            self.status_var.set("Nenhum arquivo carregado")
            return []

        try:
            self.status_var.set(f"Gerando Dashboard para {selected_territorio} ({selected_criterio})...")
            logger.info(f"Iniciando geração do Dashboard para {selected_territorio} ({selected_criterio}) com top {top_n} municípios.")

            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].fillna('0').astype(str).str.strip()
            self.df['ID_MN_RESI'] = self.df['ID_MN_RESI'].apply(lambda x: x.zfill(6) if x.isdigit() else '000000')

            df_filtered = self.df[self.df['ID_MN_RESI'].str.startswith('22')].copy()
            logger.info(f"Tamanho do DataFrame após filtro do Piauí: {len(df_filtered)}")
            if df_filtered.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, "Nenhum dado encontrado para o estado do Piauí.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set(f"Nenhum dado para {selected_territorio} (Dashboard)")
                logger.info(f"Nenhum dado encontrado para o Piauí (Dashboard).")
                return []

            df_mapped = df_filtered.copy()
            df_mapped['Nome_Municipio'] = df_mapped['ID_MN_RESI'].map(MUNICIPIOS_PIAUI).fillna('Desconhecido')

            if selected_territorio != "Todos":
                municipios_territorio = territorios_municipios.get(selected_territorio, [])
                logger.info(f"Municípios do território {selected_territorio}: {municipios_territorio}")
                df_mapped = df_mapped[df_mapped['Nome_Municipio'].isin(municipios_territorio)].copy()
                logger.info(f"Tamanho do DataFrame após filtro por território: {len(df_mapped)}")
                if df_mapped.empty:
                    self.txt_dados.delete(1.0, tk.END)
                    self.txt_dados.insert(tk.END, f"Nenhum dado encontrado para o território {selected_territorio}.\n")
                    self.btn_salvar.configure(state='disabled')
                    self.status_var.set(f"Nenhum dado para {selected_territorio} (Dashboard)")
                    logger.info(f"Nenhum dado encontrado para o território {selected_territorio} (Dashboard).")
                    return []

            if selected_criterio == "Classificação":
                df_mapped['CLASSI_FIN'] = pd.to_numeric(df_mapped['CLASSI_FIN'], errors='coerce').fillna(0).astype(int)
                classi_fin_map = {
                    10: 'Dengue', 11: 'Dengue com Sinais de Alarme', 12: 'Dengue Grave',
                    5: 'Descartado', 8: 'Inconclusivo', 0: 'Em Branco'
                }
                df_mapped['Critério'] = df_mapped['CLASSI_FIN'].map(classi_fin_map).fillna('Em Branco')
                categories = [cat for cat in ['Dengue', 'Dengue com Sinais de Alarme', 'Dengue Grave', 'Descartado', 'Inconclusivo', 'Em Branco'] if cat in selected_classifications]
                df_mapped = df_mapped[df_mapped['Critério'].isin(selected_classifications)].copy()
            else:
                df_mapped['EVOLUCAO'] = pd.to_numeric(df_mapped['EVOLUCAO'], errors='coerce').fillna(0).astype(int)
                evolucao_map = {
                    1: 'Cura', 2: 'Óbito por Dengue', 3: 'Óbito por Outras Causas',
                    4: 'Óbito em Investigação', 9: 'Ignorado', 0: 'Em Branco'
                }
                df_mapped['Critério'] = df_mapped['EVOLUCAO'].map(evolucao_map).fillna('Em Branco')
                categories = ['Cura', 'Óbito por Dengue', 'Óbito por Outras Causas', 'Óbito em Investigação', 'Ignorado', 'Em Branco']

            if df_mapped.empty:
                self.txt_dados.delete(1.0, tk.END)
                self.txt_dados.insert(tk.END, f"Nenhum dado encontrado para os critérios selecionados em {selected_territorio}.\n")
                self.btn_salvar.configure(state='disabled')
                self.status_var.set(f"Nenhum dado para {selected_territorio} com os critérios selecionados")
                logger.info(f"Nenhum dado encontrado para {selected_territorio} com os critérios selecionados.")
                return []

            # Create a pivot table for all categories
            pivot_table = pd.pivot_table(
                df_mapped,
                index='Nome_Municipio',
                columns='Critério',
                aggfunc='size',
                fill_value=0
            )

            for category in categories:
                if category not in pivot_table.columns:
                    pivot_table[category] = 0

            pivot_table = pivot_table[categories]
            pivot_table = pivot_table.astype(int)

            # Generate different types of charts for each category
            figs = []
            chart_types = ['bar', 'horizontal_bar', 'line', 'area']  # Different chart types

            for idx, category in enumerate(categories):
                top_municipios = pivot_table[category].nlargest(top_n)
                if top_municipios.empty or top_municipios.sum() == 0:
                    continue  # Skip if no data

                # Select chart type based on index (cycling through chart_types)
                chart_type = chart_types[idx % len(chart_types)]
                fig, ax = plt.subplots(figsize=(5, 3))  # Reduced size to prevent cutting

                if chart_type == 'bar':
                    # Bar chart
                    ax.bar(top_municipios.index, top_municipios.values, label=category, color='skyblue')
                    for i, v in enumerate(top_municipios.values):
                        if v > 0:
                            ax.text(i, v, f'{int(v)}', ha='center', va='bottom', fontsize=6)
                    ax.set_title(f'{category} (Barras)', fontsize=8)
                    ax.set_xlabel('Municípios', fontsize=6)
                    ax.set_ylabel('Número de Casos', fontsize=6)
                    ax.tick_params(axis='x', rotation=45, labelsize=5)
                    ax.legend(fontsize=5)
                    ax.grid(True, which="both", ls="--", alpha=0.7, axis='y')

                elif chart_type == 'horizontal_bar':
                    # Horizontal bar chart
                    ax.barh(top_municipios.index, top_municipios.values, color='lightcoral')
                    for i, v in enumerate(top_municipios.values):
                        if v > 0:
                            ax.text(v, i, f'{int(v)}', va='center', ha='left', fontsize=6)
                    ax.set_title(f'{category} (Barras Horizontais)', fontsize=8)
                    ax.set_xlabel('Número de Casos', fontsize=6)
                    ax.set_ylabel('Municípios', fontsize=6)
                    ax.tick_params(axis='y', labelsize=5)
                    ax.grid(True, which="both", ls="--", alpha=0.7, axis='x')

                elif chart_type == 'line':
                    # Line chart
                    ax.plot(top_municipios.index, top_municipios.values, marker='o', color='green', label=category)
                    for i, v in enumerate(top_municipios.values):
                        if v > 0:
                            ax.text(i, v, f'{int(v)}', ha='center', va='bottom', fontsize=6)
                    ax.set_title(f'{category} (Linha)', fontsize=8)
                    ax.set_xlabel('Municípios', fontsize=6)
                    ax.set_ylabel('Número de Casos', fontsize=6)
                    ax.tick_params(axis='x', rotation=45, labelsize=5)
                    ax.legend(fontsize=5)
                    ax.grid(True, which="both", ls="--", alpha=0.7)

                elif chart_type == 'area':
                    # Area chart
                    ax.fill_between(top_municipios.index, top_municipios.values, color='purple', alpha=0.3, label=category)
                    ax.plot(top_municipios.index, top_municipios.values, color='purple')
                    for i, v in enumerate(top_municipios.values):
                        if v > 0:
                            ax.text(i, v, f'{int(v)}', ha='center', va='bottom', fontsize=6)
                    ax.set_title(f'{category} (Área)', fontsize=8)
                    ax.set_xlabel('Municípios', fontsize=6)
                    ax.set_ylabel('Número de Casos', fontsize=6)
                    ax.tick_params(axis='x', rotation=45, labelsize=5)
                    ax.legend(fontsize=5)
                    ax.grid(True, which="both", ls="--", alpha=0.7)

                plt.tight_layout()
                figs.append(fig)

            # Update the report DataFrame
            self.current_report_df = pivot_table.reset_index()

            # Save the dashboard to a temporary file for PDF export (first page for reference)
            if figs:
                dashboard_path = f"dashboard_{selected_territorio}_{selected_criterio}_{self.ano}.png"
                figs[0].savefig(dashboard_path, dpi=150, bbox_inches='tight')
                logger.info(f"Dashboard salvo em: {dashboard_path}")
                self.current_map_path = dashboard_path

            self.txt_dados.delete(1.0, tk.END)
            self.txt_dados.insert(tk.END, f"Dashboard: {self.current_territorio} ({selected_criterio}) - {self.ano}\n\n")
            self.txt_dados.insert(tk.END, "Dashboard gerado com sucesso. Veja a janela de visualização.\n")
            logger.info(f"Dashboard de {self.current_territorio} ({selected_criterio}) gerado com sucesso.")
            return figs

        except Exception as e:
            logger.error(f"Falha ao gerar dashboard: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao gerar dashboard: {str(e)}")
            self.status_var.set("Erro ao gerar dashboard")
            return []

def main():
    root = tk.Tk()
    app = SistemaNotificacao(root)
    root.mainloop()

if __name__ == "__main__":
    main()