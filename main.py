from selenium import webdriver as driver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
import pyautogui as auto

# Nome do arquivo Excel
dados = "dados.xlsx"

# Carregando a planilha Excel
planilha = load_workbook(dados)
seleciona = planilha.active 

# Inicializando o navegador Chrome
navegador = driver.Chrome()
navegador.get('https://cadastro-produtos-devaprender.netlify.app/index.html')
auto.sleep(2)

# Obtendo o número de linhas na planilha
numero_de_linhas = len(seleciona['A'])

# Loop para iterar sobre as linhas na planilha
for linha in range(2, numero_de_linhas + 1):
    auto.sleep(2)
    
    # Obtendo os dados da planilha para cada coluna
    nome_do_produto = seleciona['A%d' % linha].value
    descricao = seleciona['B%d' % linha].value
    categoria = seleciona['C%d' % linha].value
    codigo_produto = seleciona['D%d' % linha].value
    peso = seleciona['E%d' % linha].value
    dimensoes = seleciona['F%d' % linha].value
    preco = seleciona['G%d' % linha].value
    qtd_estoque = seleciona['H%d' % linha].value
    validade = seleciona['I%d' % linha].value
    cor = seleciona['J%d' % linha].value
    tamanho = seleciona['K%d' % linha].value
    material = seleciona['L%d' % linha].value
    fabricante = seleciona['M%d' % linha].value
    pais_origem = seleciona['N%d' % linha].value
    obervacoes = seleciona['O%d' % linha].value
    codigo_barras = seleciona['P%d' % linha].value
    loc_armazem = seleciona['Q%d' % linha].value

    # Preenchendo os campos do formulário no site
    navegador.find_element(By.XPATH, '//*[@id="product_name"]').send_keys(nome_do_produto)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="description"]').send_keys(descricao)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="category"]').send_keys(categoria)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="product_code"]').send_keys(codigo_produto)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="weight"]').send_keys(peso)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="dimensions"]').send_keys(dimensoes)
    auto.sleep(2)

    navegador.find_element(By.XPATH, '/html/body/div/form/div[7]/button[1]').send_keys(Keys.ENTER)
    auto.sleep(2)

    navegador.find_element(By.XPATH, '//*[@id="price"]').send_keys(preco)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="stock"]').send_keys(qtd_estoque)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="expiry_date"]').send_keys(validade)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="color"]').send_keys(cor)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="size"]').send_keys(tamanho)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="material"]').send_keys(material)
    auto.sleep(2)

    navegador.find_element(By.XPATH, '/html/body/div/form/div[7]/button[1]').send_keys(Keys.ENTER)
    auto.sleep(2)

    navegador.find_element(By.XPATH, '//*[@id="manufacturer"]').send_keys(fabricante)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="country"]').send_keys(pais_origem)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="remarks"]').send_keys(obervacoes)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="barcode"]').send_keys(codigo_barras)
    auto.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="warehouse_location"]').send_keys(loc_armazem)
    auto.sleep(2)

    navegador.find_element(By.XPATH, '/html/body/div/form/div[6]/button[1]').send_keys(Keys.ENTER)
    auto.sleep(2)

    auto.press('enter')
    auto.sleep(2)

    navegador.find_element(By.XPATH, '/html/body/div/div/button').send_keys(Keys.ENTER)
    auto.sleep(2)    

# Fechando o navegador ao final do processo
navegador.quit()

