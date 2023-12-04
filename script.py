# Script para pegar as ocorrências abertas do sistema Implanta e organizar por área solicitante

import pandas as pd
from dotenv import load_dotenv
import os, sys
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

load_dotenv()

# Definindo as áreas do Conselho (coordenadorias)
AREAS = {
    0: 'Nenhum',
    1: 'CTMV',
    2: 'ÉTICO',
    3: 'FISCALIZAÇÃO',
    4: 'REGISTRO',
    5: 'FINANCEIRO',
    6: 'INFORMÁTICA',
    7: 'ADMINISTRATIVO',
    8: 'COMUNICAÇÃO',
    9: 'JURÍDICO',
    10: 'GABINETE',    
}

# Lista as áreas para mostrar ao usuário
def print_areas() -> str:
    text = ''
    for i in AREAS:
        text += str(i) + ': ' + AREAS[i] + '\n'
    return text

# Início da automação

# Option para não fechar o Chrome ao encerrar
# chrome_options = Options()
# chrome_options.add_experimental_option('detach', True)
# browser = webdriver.Chrome(options=chrome_options)

# Iniciando webdriver
browser = webdriver.Chrome()
browser.get('https://site.implantainfo.com.br/Solicitacoes')

# Login no sistema
username = os.getenv('USUARIO')
password = os.getenv('SENHA')

browser.find_element(By.ID, 'cpf').send_keys(username)
browser.find_element(By.ID, 'senha').send_keys(password)
browser.find_element(By.XPATH, '//*[@id="Session_Login"]/div/a').click()
browser.implicitly_wait(20)

# Filtros
situacao = browser.find_element(By.ID, 'Situacao')
situacao = Select(situacao).select_by_value('1')
browser.find_element(By.XPATH, '//*[@id="Solicitacoes_Index_Pesquisar"]/div/div[4]/a').click()
browser.implicitly_wait(20)



def preencher_elementos(i:int) -> dict:
    item = {}
    browser.find_element(By.XPATH, f'//*[@id="Solicitacoes_Index_pesquisa"]/li[{i}]/div/div[1]/a').click()

    # Salva os dados que queremos
    item['ocorrencia'] =    browser.find_element(By.XPATH, f'//*[@id="Solicitacoes_Index_pesquisa"]/li[{i}]/div/div[1]/div[1]/p[1]').text
    item['titulo'] =        browser.find_element(By.XPATH, f'//*[@id="Solicitacoes_Index_pesquisa"]/li[{i}]/div/div[1]/div[2]/p[1]').text
    item['solicitante'] =   browser.find_element(By.XPATH, f'//*[@id="Solicitacoes_Index_pesquisa"]/li[{i}]/div/div[2]/div[4]/p[1]').text                               
    item['data'] =          browser.find_element(By.XPATH, f'//*[@id="Solicitacoes_Index_pesquisa"]/li[{i}]/div/div[2]/div[1]/p[1]').text
    return item

tabela = []
def preencherTabela():            
    for i in range(1, 11):
        try:
             # Abre todas as solicitações
            item = preencher_elementos(i)
            if item['solicitante'] == '':
                item = preencher_elementos(i)

            tabela.append(item)
            #print(item)
            item = {}
            
        except NoSuchElementException:
            # para o for, quando os itens acabarem (na última tela)
            break
# Preenche a primeira página
print("Acessando dados...")
preencherTabela()

# Preenche as páginas seguintes
while True:
    ultimaPagina = browser.find_elements(By.XPATH, '/html/body/div[3]/div[2]/ul/li')[-1]
    if 'próxima' in ultimaPagina.text.lower():
        ultimaPagina.click()
        browser.implicitly_wait(5)
        preencherTabela()        
    else:
        break

tabela_df = pd.DataFrame(tabela)

## Verifica setores do Solicitante
try:
    solicitantes = pd.read_excel('_internal\\solicitantes.xlsx')
except FileNotFoundError:
    solicitantes = pd.DataFrame(columns=['solicitante', 'area'])

solicitante_raw = tabela_df['solicitante'].drop_duplicates()

for i in solicitante_raw:
    if i not in solicitantes['solicitante'].to_list():
        print('Digite o código da área do solicitante ***'+ i+ '***: ')
        while True:
            print('Opções: \n', print_areas())                    
            area = int(input())
            if area in AREAS.keys():
                break
            else:
                print('Digite um código válido. ')

        solicitantes.loc[len(solicitantes.index)] = [i, AREAS[area]]

def verificar_areas(solicitante):
    return solicitantes[solicitantes['solicitante'] == solicitante]['area'].item()

tabela_df['area'] = tabela_df['solicitante'].apply(verificar_areas)

# Preenche chamados sem área identificada automaticamente
# Pode ser, por exemplo, que o chamado foi aberto tendo como solicitante um funcionário da implanta
for i, data in tabela_df[tabela_df['area'] == 'Nenhum'].iterrows():
    print('\n\n************************************************************************************')
    print('* Não identificamos de qual área é este chamado (o usuário não pertence a nenhuma área):')
    print('* Título: ' + data['titulo'])
    print('* Solicitante: ' + data['solicitante'])
    print('* Deseja incluir em alguma área? Se sim, digite o código da área. Se não, deixe em branco.')
    print('* Opções: \n', print_areas())   
    incluir = int(input())
    if incluir in AREAS.keys():
        tabela_df.loc[i, 'area'] = AREAS[incluir]

while True:
    try:
        # Salva todos os dados juntos
        tabela_df.to_excel('tabela_todos.xlsx')
        print("Registros: ", tabela_df.shape[0])
        # Salva a tabela com os solicitantes
        solicitantes.to_excel('_internal\\solicitantes.xlsx', index=False)

        # Salva cada área em uma tabela separada
        for i in tabela_df['area'].to_list():
            tabela_df[tabela_df['area'] == i].to_excel('TABELA_'+i+'.xlsx', index=False)
        break
    except PermissionError as e:        
        print(f"Permissão de arquivo negada. Feche o arquivo {e.filename} e tente novamente.")
        print('Digite qualquer tecla para tentar novamente. ')
        input()

print('Concluído! Fechando...')
browser.quit()
sys.exit()