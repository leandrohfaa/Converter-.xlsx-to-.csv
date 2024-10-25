import os
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
import re

def converter_xlsx_para_csv(caminho_xlsx, caminho_csv):
    try:
        # Lê a planilha Excel
        df = pd.read_excel(caminho_xlsx, engine='openpyxl')
        # Salva como CSV
        df.to_csv(caminho_csv, index=False)
        print(f"Convertido com sucesso: '{caminho_xlsx}' para '{caminho_csv}'")
    except Exception as e:
        print(f"Erro ao converter '{caminho_xlsx}': {e}")

def extrair_numero(nome_arquivo):
    """Função para extrair o número do nome do arquivo."""
    match = re.search(r'(\d+)', nome_arquivo)
    return int(match.group(1)) if match else float('inf')  # Retorna infinito se não encontrar

def converter_varios_arquivos_simultaneamente(diretorio, mapeamento_nomes):
    # Lista para armazenar os caminhos dos arquivos .xlsx
    lista_arquivos = []

    # Percorre o diretório especificado e adiciona arquivos .xlsx à lista
    for arquivo in os.listdir(diretorio):
        if arquivo.endswith('.xlsx'):
            lista_arquivos.append(os.path.join(diretorio, arquivo))

    # Ordena a lista de arquivos com base no número extraído do nome do arquivo
    lista_arquivos.sort(key=lambda x: extrair_numero(os.path.basename(x)))

    with ThreadPoolExecutor() as executor:
        # Para cada arquivo XLSX, gera o caminho CSV correspondente e executa a conversão
        for caminho_xlsx in lista_arquivos:
            nome_xlsx = os.path.basename(caminho_xlsx)
            print(f"Processando: {nome_xlsx}")  # Debug: Mostra o arquivo que está sendo processado

            # Obtém o nome desejado do mapeamento
            nome_csv = mapeamento_nomes.get(nome_xlsx, nome_xlsx.replace('.xlsx', ''))  # Usa o próprio nome se não estiver no mapeamento
            
            if nome_csv not in mapeamento_nomes.values():
                print(f"Nenhum nome correspondente encontrado para: {nome_xlsx}")  # Debug: Nome não encontrado no mapeamento
            
            caminho_csv = os.path.join(diretorio, nome_csv + '.csv')
            executor.submit(converter_xlsx_para_csv, caminho_xlsx, caminho_csv)

# Diretório contendo os arquivos XLSX a serem convertidos
diretorio = '/home/pedro-pc/Área de trabalho/Agricula'

# Dicionário com mapeamento de arquivos XLSX para nomes CSV
mapeamento_nomes = {
    'SIT - Sistema Integrado de Tecnologia.xlsx': 'Setor',
    'SIT - Sistema Integrado de Tecnologia (1).xlsx': 'Unidades de Medidas',
    'SIT - Sistema Integrado de Tecnologia (2).xlsx': 'Grupos de Recursos',
    'SIT - Sistema Integrado de Tecnologia (3).xlsx': 'Recurso',
    'SIT - Sistema Integrado de Tecnologia (4).xlsx': 'Produtos',
    'SIT - Sistema Integrado de Tecnologia (5).xlsx': 'Componentes',
    'SIT - Sistema Integrado de Tecnologia (6).xlsx': 'Operações',
    'SIT - Sistema Integrado de Tecnologia (7).xlsx': 'Lista de Atividade',
    'SIT - Sistema Integrado de Tecnologia (8).xlsx': 'Motivos de Paradas',
    'SIT - Sistema Integrado de Tecnologia (9).xlsx': 'Detalhes de Paradas',
    'SIT - Sistema Integrado de Tecnologia (10).xlsx': 'Linhas de Produção',
}

# Chama a função para converter múltiplos arquivos simultaneamente
converter_varios_arquivos_simultaneamente(diretorio, mapeamento_nomes)



