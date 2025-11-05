import pandas as pd
from elasticsearch import Elasticsearch, helpers
import urllib3
import os

# 🔇 Desativa avisos de certificado não verificado
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Credenciais
user = 'gabrieloliveira30'
pwd = 'zdyaWZDwZ0GDEkaLMsE'
dengue_index = 'sinan-webservice-dengue-pi'
chikungunya_index = 'sinan-webservice-febre-chikungunya-pi'

# URL do cluster
url = 'https://sinan-es.saude.gov.br'

# Cliente Elasticsearch versão 8.x
es = Elasticsearch(
    url,
    basic_auth=(user, pwd),
    verify_certs=False
)

# Consulta básica
query = {"query": {"match_all": {}}}

# 📂 Garante que a pasta data/ exista
os.makedirs("data", exist_ok=True)

def salvar_em_arquivo(results, prefixo):
    """
    Lê os resultados do Elasticsearch e salva todos os registros em um único CSV fixo.
    Ex: data/dengue.csv
    """
    data = []
    total_count = 0

    for result in results:
        data.append(result['_source'])
        total_count += 1
        if total_count % 2000 == 0:
            print(f"{total_count} registros de {prefixo} processados...")

    if data:
        df = pd.DataFrame(data)
        filename = os.path.join("data", f"{prefixo}.csv")  # <-- Salva na pasta data/
        df.to_csv(filename, index=False, encoding='utf-8-sig')
        print(f"💾 Arquivo atualizado: {filename} ({total_count} registros)")
    else:
        print(f"⚠️ Nenhum registro encontrado para {prefixo}")

# ---- DENGUE ----
print("📥 Baixando registros de Dengue...")
dengue_results = helpers.scan(client=es, query=query, index=dengue_index)
salvar_em_arquivo(dengue_results, prefixo="dengue")

# ---- CHIKUNGUNYA ----
print("\n📥 Baixando registros de Chikungunya...")
chik_results = helpers.scan(client=es, query=query, index=chikungunya_index)
salvar_em_arquivo(chik_results, prefixo="chikungunya")

print("\n✅ Atualização concluída com sucesso!")
