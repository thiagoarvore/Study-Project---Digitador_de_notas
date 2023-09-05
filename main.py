from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.actions.wheel_input import ScrollOrigin
import pandas as pd
import datetime
import os
import sys
from sys import exit

def obtertotalalunos():
    # Obter num total de alunos
    total_alunos = secretaria.find_element(By.XPATH, '//*[@id="frmNotas"]/b/p/font')
    b = ''
    for a in total_alunos.text:
        try:
            c = int(a)        
            b+=a
        except:
            continue
    total_de_alunos = int(b)
    return total_de_alunos

num_matricula = 212895
senha_user = 212895 #input('Escreva sua senha: ') 
check_bimestre = False

while check_bimestre == False:
    bimestre = 3 #int(input('Insira o bimestre desejado: '))
    if bimestre in range(1,5):
        check_bimestre = True
ano_atual = datetime.datetime.now().year
# Pasta onde o programa estiver no pc do user
pasta_programa = os.path.dirname(os.path.abspath(__file__))
# Pasta onde as tabelas de nota estão localizadas
pasta_arquivos = os.path.join(pasta_programa, "TABELAS")
# Entrar na secretaria online
secretaria = webdriver.Chrome()
secretaria.get('https://professoronline.objetivo.br/lyceump/donline/dolmain.asp')
janela_principal = secretaria.current_window_handle
try: 
    alert = Alert(secretaria)
    alert.accept()
except:
    pass
secretaria.set_window_size(1920,1080)
login = secretaria.find_elements(By.XPATH, '//input[@id="txtnumero_matricula"]')
login[0].send_keys(num_matricula)
senha = secretaria.find_elements(By.XPATH, '//input[@type="password"]')
try: 
    senha[0].send_keys(senha_user)
    botao_login = secretaria.find_elements(By.XPATH, '//input[@id="image1"]')
    botao_login[0].click()
    try: 
        alert = Alert(secretaria)
        alert.accept()
    except:
        pass
    secretaria.set_window_size(1920,1080)
except:
    input('Sua senha parece estar errada, aperte enter para reiniciar.')
# Acessar a lista de turmas
sleep(2)
passar_mouse_notas = secretaria.find_elements(By.XPATH, '//li[@id="M3"]')
actions = ActionChains(secretaria)
actions.move_to_element(passar_mouse_notas[0])
actions.perform
wait = WebDriverWait(secretaria, 10)
bot_notas = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="M3"]')))
bot_notas.click()
sleep(1)
bot_notas2 = secretaria.find_element(By.XPATH, '//*[@id="M3L1"]')
sleep(1)
bot_notas2.click()
sleep(3)
# Só pegar turmas que o professor tem no ano atual
turmas_site = []
for a in range(3, 20):
    try:    
        xpath_turma = f'//*/tbody/tr[{a}]/td[2][@class="font01"]'
        xpath_ano = f'//*/tbody/tr[{a}]/td[4][@class="font01"]'
        ano = secretaria.find_element(By.XPATH, xpath_ano)
        if str(ano_atual) in ano.text:
            turma_alvo = secretaria.find_element(By.XPATH, xpath_turma)        
            turmas_site.append(turma_alvo)
    except:
        break
total_de_turmas = len(turmas_site)
# Abrir turma no site a partir de nova lista 
# (porque ele não lembra os elementos depois de voltar para a página)
for i in range(0,total_de_turmas-1):
    coluna_alunos = 0
    coluna_nota = 0 
    turmas_site_interno = []
    for a in range(3, 20):
        try:    
            xpath_turma = f'//*/tbody/tr[{a}]/td[2][@class="font01"]'
            xpath_ano = f'//*/tbody/tr[{a}]/td[4][@class="font01"]'
            ano = secretaria.find_element(By.XPATH, xpath_ano)
            if str(ano_atual) in ano.text:
                turma_alvo = secretaria.find_element(By.XPATH, xpath_turma)        
                turmas_site_interno.append(turma_alvo)
        except:
            break
    sleep(0.1)
    turma = turmas_site_interno[i]    
    nome_turma = turma.text[3:]
    print(nome_turma)
    # Selecionar o bimestre
    turma.click()
    dropdown_bim = secretaria.find_element(By.XPATH, 
                                            '//*[@id="frmNotas"]/b/table/tbody/tr/td[3]/select')
    opcao_bim = Select(dropdown_bim)
    opcao_bim.select_by_value(str(bimestre))
    num_de_alunos = obtertotalalunos()

    # Escolher a tabela correta da turma correspondente na pasta de tabelas
    for nome_arquivo in os.listdir(pasta_arquivos):
        if nome_turma in nome_arquivo:
            caminho_arquivo = os.path.join(pasta_arquivos, nome_arquivo)
            if caminho_arquivo.endswith('.xlsx'):
                planilha = pd.read_excel(caminho_arquivo)
            else:
                print(f'Não consegui encontrar a tabela de notas do {nome_turma}')                
    # Pegar as planilhas de cada bimestre 
    planilhas = pd.ExcelFile(caminho_arquivo).sheet_names
    tabela1 = f'{nome_turma}.xlsx'
    planilha_alvo = 0
    # Escolher a planilha do bimestre correto
    for planilha_nome in planilhas:        
        if str(bimestre) in list(planilha_nome):
            caminho_planilha = os.path.join('TABELAS', tabela1)
            planilha_alvo = pd.read_excel(caminho_planilha, sheet_name=planilha_nome)
            break            
    if planilha_alvo is None:
        input(f'Não encontrei a planilha do bimestre desejado na tabela de notas do {nome_turma}, aperte enter para reiniciar.')
        sys.exit()
    # Estabelecer as colunas corretas da tabela
    coluna_alunos = planilha_alvo['NOME']
    coluna_nota = planilha_alvo['NOTA']       
    # Percorrer cada aluno
    lista_nota_ok = []
    lista_teste = []
    loop = 0
    tabela = secretaria.find_element(By.XPATH, '//*[@id="tableDiv_General"]')
    for i in range(1, num_de_alunos+1):
        nota = 0        
        # Encontrar o aluno e deixar o nome minúsculo
        try:
            xpath_aluno = f'//*[@id="tableDiv_General"]/div/div[1]/div[2]/table/tbody/tr[{i}]/td[2]'
            aluno = secretaria.find_element(By.XPATH, xpath_aluno)
            nome_aluno = aluno.text[:-14]
            nome_aluno_ok = nome_aluno.lower()
            sleep(0.3)
            # Procurar o nome e achar a nota
            for u, nome in enumerate(coluna_alunos):
                if nome_aluno_ok in nome.lower(): #colocando o nome em minuscula
                    nota = coluna_nota[u]
                    break            
            # Achar onde colocar a nota no site
            if bimestre == 1 or bimestre == 3:     
                xpath_nota_aluno = f'//*[@id="Open_Text_General"]/tbody/tr[{i}]/td[3]/input'
            if bimestre == 2 or bimestre == 4:
                pass ##############################
            ######################## Arrumar isso #######
            ############
            nota_aluno = secretaria.find_element(By.XPATH, xpath_nota_aluno) 
            #Verificar se a nota é um número de verdade   
            try:
                nota = int(nota)
            except ValueError:
                nota = 0

            # Colocar a nota no local certo
            sleep(0.3)            
            try:
                nota_aluno.send_keys(nota)
            except:
                pass
            # Levar a barra de rolagem pra baixo e carregar os outros alunos
            loop+=1
            if loop % 13 == 0:
                secretaria.execute_script("window.scrollTo(0,document.body.scrollHeight)")
                nota_aluno.send_keys(Keys.PAGE_DOWN)
        except:
            print(f'Não encontrei {nome_aluno}, talvez o nome esteja diferente do site.')             

        # Salvar as notas
    botao_salvar = secretaria.find_element(By.XPATH, '//*[@id="imagem"]/p/a')       
    botao_salvar.click()
    sleep(0.2)
    try: 
        alert = Alert(secretaria)
        alert.accept()
    except:
        pass
    sleep(1.5)

        # Voltar para as turmas
    secretaria.back()
    sleep(1.3)
    secretaria.back()
    sleep(0.3)
    secretaria.back()
    sleep(0.5)
