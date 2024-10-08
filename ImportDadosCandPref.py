from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import time
from datetime import datetime


driver = webdriver.Chrome()


urls = [
    'https://divulgacandcontas.tse.jus.br/divulga/#/candidato/SUDESTE/SP/2045202024/250002078851/2024/71072',
    'https://divulgacandcontas.tse.jus.br/divulga/#/candidato/SUDESTE/SP/2045202024/250002355541/2024/71072',
    'https://divulgacandcontas.tse.jus.br/divulga/#/candidato/SUDESTE/SP/2045202024/250002180213/2024/71072',
    'https://divulgacandcontas.tse.jus.br/divulga/#/candidato/SUDESTE/SP/2045202024/250001926547/2024/71072',
    'https://divulgacandcontas.tse.jus.br/divulga/#/candidato/SUDESTE/SP/2045202024/250002362195/2024/71072',
    'https://divulgacandcontas.tse.jus.br/divulga/#/candidato/SUDESTE/SP/2045202024/250001884312/2024/71072',
    'https://divulgacandcontas.tse.jus.br/divulga/#/candidato/SUDESTE/SP/2045202024/250001978066/2024/71072',
    'https://divulgacandcontas.tse.jus.br/divulga/#/candidato/SUDESTE/SP/2045202024/250002098117/2024/71072',
    'https://divulgacandcontas.tse.jus.br/divulga/#/candidato/SUDESTE/SP/2045202024/250002031025/2024/71072',
    'https://divulgacandcontas.tse.jus.br/divulga/#/candidato/SUDESTE/SP/2045202024/250002163891/2024/71072'
]


wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Dados Extraídos"
ws.append(['URL', 'NOME', 'DATA DE NASCIMENTO', 'IDADE', 'GENERO', 'COR / RACA', 'ESCOLARIDADE', 'PARTIDO', 'LIMITE LEGAL DE GASTOS 1° TURNO', 'TOTAL LÍQUIDO DE RECURSOS RECEBIDOS', 'TOTAL DE DESPESAS'])


for url in urls:
    print(f"Acessando: {url}")
    driver.get(url)

    e
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "basicInformationSection"))
        )
    except Exception as e:
        print(f"Erro ao carregar a página para {url}: {e}")
        continue  

    dados_extraidos = {
        'URL': url,
        'NOME': "Dado não encontrado",
        'DATA DE NASCIMENTO': "Dado não encontrado",
        'IDADE': "Dado não encontrado",
        'GENERO': "Dado não encontrado",
        'COR / RACA': "Dado não encontrado",
        'ESCOLARIDADE': "Dado não encontrado",
        'PARTIDO': "Dado não encontrado",
        'LIMITE LEGAL DE GASTOS 1° TURNO': "Dado não encontrado",
        'TOTAL LÍQUIDO DE RECURSOS RECEBIDOS': "Dado não encontrado",
        'TOTAL DE DESPESAS': "Dado não encontrado"
    }

    
    xpaths = {
        'NOME': '//*[@id="basicInformationSection"]//label[contains(text(), "Nome Completo:")]/following-sibling::label',
        'DATA DE NASCIMENTO': '//*[@id="basicInformationSection"]//label[contains(text(), "Data de Nascimento:")]/following-sibling::label',
        'GENERO': '//*[@id="basicInformationSection"]//label[contains(text(), "Gênero:")]/following-sibling::label',
        'COR / RACA': '//*[@id="basicInformationSection"]//label[contains(text(), "Cor / Raça:")]/following-sibling::label',
        'ESCOLARIDADE': '//*[@id="basicInformationSection"]//label[contains(text(), "Grau de Instrução:")]/following-sibling::label',
        'PARTIDO': '/html/body/dvg-root/main/dvg-canditado-detalhe/div/div/div[1]/dvg-candidato-header/div/div/div/span/label[2]',
        'LIMITE LEGAL DE GASTOS 1° TURNO': '//*[@id="basicInformationSection"]//label[contains(text(), "Limite Legal de Gastos 1º Turno:")]/following-sibling::span',
        'TOTAL LÍQUIDO DE RECURSOS RECEBIDOS': '/html/body/dvg-root/main/dvg-canditado-detalhe/dvg-prestacao-conta-candidato/div/div/div/div[1]/div[2]/dvg-receita-prestacao-contas/div/div/div[1]/div/dvg-receita-prestacao-contas-item/div/p[1]',
        'TOTAL DE DESPESAS': '/html/body/dvg-root/main/dvg-canditado-detalhe/dvg-prestacao-conta-candidato/div/div/div/div[1]/div[3]/dvg-despesa-prestacao-contas/div/div/dvg-despesa-item[2]/div/div/p[1]'
    }

    
    for key, xpath in xpaths.items():
        try:
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            dados_extraidos[key] = element.text.strip()
        except Exception as e:
            print(f"Erro ao extrair {key} de {url}: {e}")

    
    if dados_extraidos['DATA DE NASCIMENTO'] != "Dado não encontrado":
        try:
            data_nascimento = datetime.strptime(dados_extraidos['DATA DE NASCIMENTO'], "%d/%m/%Y")
            idade = (datetime.now() - data_nascimento).days // 365  # Calcula a idade em anos
            dados_extraidos['IDADE'] = idade
        except ValueError:
            print(f"Erro ao converter a data de nascimento: {dados_extraidos['DATA DE NASCIMENTO']}")

 
    try:
        
        driver.refresh()
        time.sleep(2)  

        
        for key, xpath in xpaths.items():
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                dados_extraidos[key] = element.text.strip()
            except Exception as e:
                print(f"Erro ao extrair {key} de {url} após refresh: {e}")

    except Exception as e:
        print(f"Erro ao extrair TOTAL LÍQUIDO DE RECURSOS RECEBIDOS ou TOTAL DE DESPESAS de {url}: {e}")

    ws.append(list(dados_extraidos.values()))
    time.sleep(2)  

wb.save('PREF_SP.xlsx')
print("Dados salvos em 'PREF_SP.xlsx' com sucesso!")

driver.quit()
