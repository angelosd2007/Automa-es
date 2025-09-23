# importando bibliotecas
# %
import re
from pykeepass import PyKeePass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import getpass
import pandas as pd
import numpy as np
import requests
import os
import time
# %

segunda_rodada = input('É a segunda rodada? (s/n) s para sim ou n para não: ')

# Configurações do caminho
download_dir = r'C:\Users\u004047\Downloads'
# %
# Acessando usuário e senha KeePass
frase_segura = getpass.getpass("Digite sua senha KeePass: ")
kp = PyKeePass(r'C:\Users\u004047\Documents\Consulta SCR\senhas.kdbx', password=frase_segura)

# Pergunta qual empresa acessar
empresa = input("Qual empresa deseja acessar? (1 para financeira, 2 para pagamentos): ")

# Configurações do navegador
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True
})

# Acessando o site com webdriver
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://www3.bcb.gov.br/rdr/consultaDemandaIFView.do?method=consultaDemandaNumeroInicial")

# Aguarda o carregamento completo da página inicial
WebDriverWait(driver, 10).until(lambda d: d.execute_script("return document.readyState") == "complete")

# Encontra as credenciais
if empresa == '1':
    entry = kp.find_entries(title='Neon Financeira Bacen', first=True)
else:
    entry = kp.find_entries(title='Nome_da_Entrada_Pagamentos', first=True)

if entry is None:
    print("Entrada não encontrada no KeePass.")
    driver.quit()
    exit()

username = entry.username
password = entry.password

# Insere o usuário e senha
driver.find_element(By.XPATH, '//*[@id="userNameInput"]').send_keys(username)
driver.find_element(By.XPATH, '//*[@id="passwordInput"]').send_keys(password)

# Clica no botão de login
driver.find_element(By.XPATH, '//*[@id="submitButton"]').click()
# %
# % condição
if segunda_rodada == 'n':
    # Carrega a planilha com os Identificadores
    df = pd.read_excel(r'C:\Users\u004047\Downloads\Lista de Demandas - Demandas - Classificar.xlsx', header=3)
    # Ajusta a Base para o formato correto:
    df = df[['Status', 'Canal Entrada', 'Tipo', 'Código Interno ', 'Identificador', 'Assunto', 'Descrição', 'Demandante Direto', 'Recebida', 'Data Sistema']]
    # Renomea as Colunas necessárias:
    df = df.rename(columns={'Assunto': 'Tema', 'Código Interno ': 'Código Interno'})
    # Cria Coluna Analista
    df.insert(3, 'Analista', '')
    # %
    # Filtra apenas BACEN
    df_bc = df[df['Canal Entrada'] == 'BACEN NEON']
    df_ouv = df[df['Canal Entrada'] == 'OUV Site Neon']
    # %
    # Define colunas de acordo com a análise:
    df_bc['Descrição'] = None
    df_bc['Mensagem DEATI'] = None

    # Função para limpar caracteres inválidos
    def limpar_texto(texto):
        if texto:
            texto = re.sub(r'[<>:"/\\|?*]', '', texto)  # Remove caracteres básicos
            texto = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', texto)  # Remove outros caracteres de controle
            texto = texto.replace('→', '-')  # Substitui caracteres especiais
            return texto
        return texto
if segunda_rodada == "s":
    # %Carrega a planilha com os Identificadores
    df = pd.read_excel(r'C:\Users\u004047\Downloads\Lista de Demandas - Demandas - Classificar.xlsx', header=3)
    df_ant = pd.read_excel(r'C:\Users\u004047\Downloads\classificas_por_IA.xlsx')
    # %
    # Ajusta a Base para o formato correto:
    df = df[['Status', 'Canal Entrada', 'Tipo', 'Código Interno ', 'Identificador', 'Assunto', 'Descrição', 'Demandante Direto', 'Recebida', 'Data Sistema']]
    # Renomea as Colunas necessárias:
    df = df.rename(columns={'Assunto': 'Tema', 'Código Interno ': 'Código Interno'})
    # %
    # Cria Coluna Analista
    df.insert(3, 'Analista', '')
    # % excluíndo colunas do df_new
    df_ant.drop(['Assunto', 'Resumo'], axis=1, inplace=True)
    # % Padronizando colunas antes do concat
    df['Código Interno'] = (df['Código Interno'].astype(str).str.strip().str.upper())
    df_ant['Código Interno'] = (df_ant['Código Interno'].astype(str).str.strip().str.upper())
    # %
    # # encontrando os novos casos:
    df = df[~df['Código Interno'].isin(df_ant['Código Interno'])]
    # %
    # Filtra apenas BACEN
    df_bc = df[df['Canal Entrada'] == 'BACEN NEON']
    df_ouv = df[df['Canal Entrada'] == 'OUV Site Neon']
    # %
    # Define colunas de acordo com a análise:
    df_bc['Descrição'] = None
    df_bc['Mensagem DEATI'] = None

    # Função para limpar caracteres inválidos
    def limpar_texto(texto):
        if texto:
            texto = re.sub(r'[<>:"/\\|?*]', '', texto)  # Remove caracteres básicos
            texto = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', texto)  # Remove outros caracteres de controle
            texto = texto.replace('→', '-')  # Substitui caracteres especiais
            return texto
        return texto
else:
    print('Digite uma opção válida (s ou n)')   
# Loop para processar cada demanda
for index, row in df_bc.iterrows():
    identificador = row['Identificador']

    # Aguarda a página carregar e o elemento estar presente
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="idConsulta"]'))
    )

    # Tenta inserir o identificador
    driver.find_element(By.XPATH, '//*[@id="idConsulta"]').clear()
    driver.find_element(By.XPATH, '//*[@id="idConsulta"]').send_keys(identificador)
    driver.find_element(By.XPATH, '/html/body/div[10]/form/div/input').click()

    # Aguarda a descrição estar presente e visível
    descricao_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//*[@id="respondeDemandaForm"]/table/tbody/tr[24]/td[2]'))
    )

    # Captura todo o texto contido no <td>
    descricao = descricao_element.text
    descricao = limpar_texto(descricao)

    # Salva a descrição na coluna 'Descrição' do DataFrame
    df_bc.at[index, 'Descrição'] = descricao

    if "Mensagem" in descricao and "---------- Final dos dados preenchidos pelo cidadão ----------" in descricao:
        partes = descricao.split("Mensagem", 1)
        relato = partes[1].split("---------- Final dos dados preenchidos pelo cidadão ----------", 1)[0]
        relato_limpo = relato.strip()
        df_bc.at[index, 'Descrição'] = relato_limpo
    elif "DEATI" in descricao:
        partes = descricao.split("DEATI", 1)
        df_bc.at[index, 'Mensagem DEATI'] = limpar_texto(partes[1].strip())
        df_bc.at[index, 'Descrição'] = limpar_texto(partes[0].strip())
    else:
        df_bc.at[index, 'Descrição'] = ""

    # Retorna ao link inicial para processar a próxima demanda
    driver.get("https://www3.bcb.gov.br/rdr/consultaDemandaIFView.do?method=consultaDemandaNumeroInicial")
    WebDriverWait(driver, 10).until(lambda d: d.execute_script("return document.readyState") == "complete")
# Tratando a coluna Descrição
# Condição: a coluna 'Descrição' é nula (isna) OU é uma string vazia ('')
condicao = (df_bc['Descrição'].isna()) | (df_bc['Descrição'] == '')
# Aplica a lógica: se a condição for verdadeira, usa 'Mensagem DEATI', senão, mantém 'Descrição'
df_bc['Descrição'] = np.where(condicao, df_bc['Mensagem DEATI'], df_bc['Descrição'])
# Ecluíndo coluna Mensagem DEATI
df_bc = df_bc.drop(columns=['Mensagem DEATI'], axis=1)
# Juntando os DFs antes de mandar para a IA
df = pd.concat([df_bc, df_ouv], axis=0)
# Salvar a base para a IA pegar e analisar
# Salvar a planilha como "Classificar atualizada"
df.to_excel(r'C:\Users\u004047\Downloads\IA_classificar_demandas.xlsx', index=False)

# Caminho do arquivo
file_path = r'C:\Users\u004047\Downloads\IA_classificar_demandas.xlsx'
output_file_path = r'C:\Users\u004047\Downloads\classificas_por_IA.xlsx'
API_TOKEN = "sk-9ff5338c1aad424da8b429b9901bcea4"
BASE_URL = "https://gepeto.svc.in.devneon.com.br/"

# Função para definir o cabeçalho de autenticação
def get_headers(token):
    return {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

# Ler a planilha do Excel
df = pd.read_excel(file_path)

# Filtra somente registros com 'Descrição'
df = df[df['Descrição'].notnull()]

# Criar novas colunas para armazenar as respostas
df['Assunto'] = ''
df['Resumo'] = ''

# Função para enviar uma mensagem e receber uma resposta da API
def enviar_mensagem(mensagem):
    payload = {
        "model": "gpt-4o-mini",  
        "messages": [
            {"role": "system", "content": """Você é um especialista de atendimento ao cliente altamente eficiente. Sua tarefa é analisar a reclamação do cliente e fornecer:
1.  **Assunto Principal:** Identifique e classifique o assunto principal da reclamação em até 3 palavras. Baseie-se nos seguintes exemplos, mas não se limite a eles: Contestação de Compra, Golpe Pix, Conta Bloqueada, SCR (BACEN), Negociação de Dívidas, Superendividamento, Pix Equívoco, Falha Acesso App, Alteração Cadastral. Seja o mais específico possível dentro do limite de palavras.
2.  **Resumo Conciso:** Elabore um resumo da demanda em no máximo 2 linhas. Este resumo deve destacar os pontos e preocupações centrais apresentados pelo cliente. É fundamental que você mencione explicitamente se o cliente relatou contatos anteriores com nossos canais de atendimento (e quais, se detalhado na reclamação).

Responda estritamente no seguinte formato, sem nenhuma introdução ou comentário adicional:
Assunto: [Seu assunto aqui]
Resumo: [Seu resumo aqui]

Se a mensagem do usuário for ininteligível, não contiver uma reclamação clara, ou for impossível de analisar, responda:
Assunto: Análise Prejudicada
Resumo: Não foi possível identificar uma reclamação clara ou informações suficientes para análise na mensagem fornecida.
"""},
            {"role": "user", "content": mensagem}
        ]
    }

    response = requests.post(
        f"{BASE_URL}/api/chat/completions",
        headers=get_headers(API_TOKEN),
        json=payload
    )

    if response.status_code == 200:
        resposta = response.json()
        print("Processado", cod_interno)
        return resposta["choices"][0]["message"]["content"]
    else:
        print(f"Erro na requisição: {response.status_code} - {response.text}")
        return None

# Processar cada descrição
for index, row in df.iterrows():
    descricao = row['Descrição']
    cod_interno = row['Código Interno']
    
    # Monta o prompt
    prompt = f"""
Leia as reclamações em {descricao} e classifique dentro dos assuntos possíveis que o cliente está reclamando, por exemplo: Contestação de Compra, Golpe Pix,
Conta Bloqueada, SCR (registro no BACEN), Negociação de Dividas, Superendividamento, Pix Equivoco, Falha no Acesso ao Aplicativo, Alteração Cadastral etc.

Redija um resumo conciso, com até 2 linhas, destacando os principais pontos e preocupações apresentados pelo cliente, incluíndo contatos anteriores em nossos canais de atendimento.
Responda da seguinte forma: Assunto:, Resumo:
"""

    # Enviando a mensagem e recebendo a resposta
    resposta = enviar_mensagem(prompt)
    if resposta:
        # Aqui, você pode separar o assunto e o resumo, caso eles estejam no formato adequado
        partes_resposta = resposta.split("\n")
        if len(partes_resposta) >= 2:
            df.at[index, 'Assunto'] = partes_resposta[0].replace('Assunto:', '').strip()
            df.at[index, 'Resumo'] = partes_resposta[1].replace('Resumo:', '').strip()
    print('Aguardando 5 segundos para nova requisição')
    time.sleep(5)

# Salvar o resultado em uma nova planilha
df.to_excel(output_file_path, index=False)
print(f"Arquivo atualizado salvo em: {output_file_path}")
