from time import sleep
from unidecode import unidecode
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import datetime
import os
import sys


def obtertotalalunos():
    # Obter num total de alunos
    total_alunos = secretaria.find_element(By.XPATH, '//*[@id="frmNotas"]/b/p/font')
    b = ''
    for a in total_alunos.text:
        try:
            b += a
        except:
            continue
    total_de_alunos = int(b)
    return total_de_alunos


def arredonda_nota(numero):
    try:
        numero = float(numero)
        numero_arredondado = round(numero, 1)
        return numero_arredondado
    except ValueError:
        return None


def inicializar():
    print(
        '''
        Esse programa usa tabelas do Excel para automatizar a inserção de notas no portal Objetivo.

        Seu computador precisa ter o Windows e o Google Chrome instalado.

        Certifique-se que o nome dos alunos estão exatamente iguais em sua tabela de notas e no site
        (aqui letras maiúsculas e minúsculas não importam).

        Certifique-se que não há linhas vazias no meio da planilha.

        Certifique-se que cada tabela é um arquivo de uma única turma,
        e que cada bimestre esteja em uma planilha diferente nesse mesmo arquivo.

        O nome da planilha do bimestre deve conter apenas o número do bimestre, mas pode conter outras letras.

        O nome da tabela (arquivo Excel) deve estar EXATAMENTE no formato: 6M1, 9M2, 7M3 (M maiúsculo).

        Em cada bimestre, necessariamente deve existir uma coluna denominada NOME (em letra maiúscula)
        e outra coluna denominada NOTA (em letra maiúscula),
        que conterão os nomes e notas finais dos alunos, respectivamente.

        Evite movimentar o mouse durante o processamento do programa, ele vai levar poucos minutos
        '''
    )


inicializar()
input('Aperte qualquer tecla para iniciar')
# Inicializar algumas varáveis
num_matricula = 212895
senha_user = input('Escreva sua senha: ')
check_bimestre = False
ano_atual = datetime.datetime.now().year
while not check_bimestre:
    bimestre = int(input('Insira o bimestre desejado: '))
    if bimestre in range(1, 5):
        check_bimestre = True
if not check_bimestre:
    input('O bimestre digitado não é válido. Pressione qualquer tecla para fechar o programa')
    sys.exit()

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
secretaria.set_window_size(1920, 1080)
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
    secretaria.set_window_size(1920, 1080)
except:
    input('Sua senha parece estar errada, aperte enter para reiniciar.')
# Acessar a lista de turmas
sleep(2)
selecionar_unidade = secretaria.find_element(By.XPATH, '/html/body/div/center/table/tbody/tr/td/div/center/table[3]/tbody/tr[5]/td/div/center/p/input')
selecionar_unidade.click()
try:
    alert = Alert(secretaria)
    alert.accept()
except:
    pass
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
sleep(1)
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
for i in range(0, total_de_turmas):
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
    # Escolher a tabela correta da turma correspondente na pasta de tabelas
    for nome_arquivo in os.listdir(pasta_arquivos):
        if nome_turma in nome_arquivo:
            caminho_arquivo = os.path.join(pasta_arquivos, nome_arquivo)
            if caminho_arquivo.endswith('.xlsx'):
                planilha = pd.read_excel(caminho_arquivo)
            else:
                print(f'Não consegui encontrar a tabela de notas do {nome_turma}')
    # Pegar as planilhas de cada bimestre
    try:
        planilhas = pd.ExcelFile(caminho_arquivo).sheet_names
    except:
        print(f'Não encontrei a planilha da turma {nome_turma} na pasta de tabelas')
        continue
    tabela1 = f'{nome_turma}.xlsx'
    # Selecionar o bimestre
    turma.click()
    dropdown_bim = secretaria.find_element(By.XPATH,
                                           '//*[@id="frmNotas"]/b/table/tbody/tr/td[3]/select')
    opcao_bim = Select(dropdown_bim)
    opcao_bim.select_by_value(str(bimestre))
    num_de_alunos = obtertotalalunos()
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
    for i in range(1, num_de_alunos + 1):
        nota = 0
        # Encontrar o aluno e deixar o nome minúsculo
        try:
            xpath_aluno = f'//*[@id="tableDiv_General"]/div/div[1]/div[2]/table/tbody/tr[{i}]/td[2]'
            aluno = secretaria.find_element(By.XPATH, xpath_aluno)
            nome_aluno = aluno.text[:-14]
            nome_aluno_ok = unidecode(nome_aluno).lower()
            sleep(0.3)
            # Procurar o nome e achar a nota
            for u, nome in enumerate(coluna_alunos):
                if nome_aluno_ok in unidecode(nome).lower():  # colocando o nome em minuscula
                    nota_raw = coluna_nota[u]
                    # arredondar a nota para uma casa decimal
                    nota = arredonda_nota(nota_raw)
                    # Verificar se a nota é um número de verdade
                    try:
                        nota = float(nota)
                        if nota > 10:
                            nota = 10
                        if nota < 0:
                            nota = 0
                    except ValueError:
                        nota = 0
                    break
            # Achar onde colocar a nota no site
            if bimestre == 1 or bimestre == 3:
                xpath_nota_aluno = f'//*[@id="Open_Text_General"]/tbody/tr[{i}]/td[3]/input'
            if bimestre == 2 or bimestre == 4:
                xpath_nota_aluno = f'//*[@id="Open_Text_General"]/tbody/tr[{i}]/td[2]/input'
            nota_aluno = secretaria.find_element(By.XPATH, xpath_nota_aluno)
            # Colocar a nota no local certo
            sleep(0.3)
            try:
                nota_aluno.send_keys(nota)
            except:
                pass
            # Levar a barra de rolagem pra baixo e carregar os outros alunos
            loop += 1
            if loop % 13 == 0:
                secretaria.execute_script("window.scrollTo(0,document.body.scrollHeight)")
                nota_aluno.send_keys(Keys.PAGE_DOWN)
        except:
            print(f'Não encontrei {nome_aluno}, talvez o nome esteja diferente do site.')
    # Salvar as notas
    botao_salvar = secretaria.find_element(By.XPATH, '//*[@id="imagem"]/p/a')
    botao_salvar.click()
    print(f'{nome_turma} foi inserida')
    sleep(0.2)
    try:
        alert = Alert(secretaria)
        alert.accept()
    except:
        pass
    sleep(1.5)
    # Voltar para as turmas
    wait = WebDriverWait(secretaria, 10)
    bot_notas = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmNotas"]')))
    secretaria.back()
    sleep(1.3)
    secretaria.back()
    sleep(0.3)
    secretaria.back()
    sleep(0.5)
