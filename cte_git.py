from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By  # Importe para permitir a seleção por 'name'
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import pyautogui
import openpyxl
import pyperclip
import locale
import keyboard
import threading

# Variável para controlar o encerramento do código
encerrar_codigo = False

# Função para aguardar a tecla "ESC" ser pressionada
def aguardar_esc():
    global encerrar_codigo
    keyboard.wait("esc")
    print("Tecla 'ESC' pressionada. Encerrando o código.")
    encerrar_codigo = True

# Inicializar a thread para aguardar a tecla "ESC"
esc_thread = threading.Thread(target=aguardar_esc)
esc_thread.start()

# Defina a localização para 'pt_BR'
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Caminho completo para o arquivo da planilha Excel
caminho_arquivo_excel = 'X:/Samuel/XML CT-e/Plan CTE/PLAN.CTE.xlsx'

# Abra o arquivo da planilha Excel
workbook = openpyxl.load_workbook(caminho_arquivo_excel)

# Selecione a planilha com a qual você deseja trabalhar
sheet = workbook['PLAN']

# Acesse o valor da célula K5
valor_celula = sheet['F7'].value

# Copie o valor da célula para a área de transferência
pyperclip.copy(valor_celula)

# Feche o arquivo da planilha Excel
workbook.close()

# Configurações do Chrome
options = Options()
options.add_argument("--start-maximized")

# Substitua o caminho abaixo pelo caminho real da sua extensão
caminho_extensao = r"C:\Users\Cigaax\AppData\Local\Google\Chrome\User Data\Default\Extensions\cfmflaihjcfpoakojijhcpcgpfjbkogn\1.3.8.3_0"
options.add_argument(f"load-extension={caminho_extensao}")

# Inicializar o driver do Chrome
driver = webdriver.Chrome(options=options)

driver.get("https://erp.nfservice.com.br/users/login")

# Espere até que o campo de entrada seja visível
campo_username = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.NAME, 'data[User][username]'))
)
# Encontre o campo de entrada pelo nome de usuario
campo_username = driver.find_element(By.NAME, 'data[User][username]')

# Encontre o campo de entrada de senha pelo nome
campo_senha = driver.find_element(By.NAME, 'data[User][password]')

# Preencha o campo de entrada com o texto desejado
texto_para_preencher = 'email'
campo_username.send_keys(texto_para_preencher)

# Preencha o campo de senha
texto_senha = 'senha'
campo_senha.send_keys(texto_senha)

# Encontre o botão "Entrar no Sistema" pelo ID
botao_entrar = driver.find_element(By.ID, 'btnLogar')

# Clique no botão para fazer login
botao_entrar.click()

# Localize o elemento pelo texto parcial do link
elemento_faturamento = driver.find_element(By.PARTIAL_LINK_TEXT, 'Faturamento')

# Clique na opção de faturamento
elemento_faturamento.click()

# Localize o elemento pelo texto parcial do link
elemento_cte = driver.find_element(By.PARTIAL_LINK_TEXT, 'CT-e')

# Clique na opção CT-e
elemento_cte.click()

# Localize o elemento pelo texto parcial do link
elemento_cte = driver.find_element(By.PARTIAL_LINK_TEXT, 'Emitir Novo')

# Clique na opção CT-e
elemento_cte.click()

# Localize o botão pelo ID "refNfe" e o valor (texto) "Referenciar NF-e"
elemento_referenciar_nfe = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, '//input[@id="refNfe" and @value="Referenciar NF-e"]'))
)

# Clique no botão
elemento_referenciar_nfe.click()

# Aguarde alguns segundos para garantir que a página esteja totalmente carregada
time.sleep(2)

# Clica em escolher arquivo
pyautogui.moveTo(280, 245)
pyautogui.sleep(1)
pyautogui.click()

# Pasta XML CT-e
pyautogui.moveTo(84, 250)
pyautogui.sleep(1)
pyautogui.click()

# Campo vazio
pyautogui.moveTo(209, 414)
pyautogui.sleep(1)
pyautogui.click()

# Cola o numero da nota
pyautogui.hotkey('ctrl', 'v')

# Clica em abrir
pyautogui.moveTo(510, 440)
pyautogui.click()
pyautogui.sleep(2)

# Clica em fechar área de escolha do arquivo XML
pyautogui.moveTo(1084, 397)
pyautogui.click()

# Clica no campo "Valor de prestação"
pyautogui.sleep(4)
pyautogui.moveTo(180, 535)
pyautogui.doubleClick()

workbook = openpyxl.load_workbook(caminho_arquivo_excel)
sheet = workbook['PLAN']
célula_com_fórmula = sheet['E7']
# Obtém o valor da célula como um número
valor_numérico = célula_com_fórmula.value

pyperclip.copy(célula_com_fórmula)

workbook.close()

pyautogui.hotkey('ctrl', 'v')

# Preenche campo "Tipo CTE"
pyautogui.moveTo(579, 487)
pyautogui.click()
pyautogui.typewrite('5353')
pyautogui.hotkey('enter')

# Preenche condição pagamento
pyautogui.moveTo(686, 489)
pyautogui.click()
pyautogui.hotkey('down')
pyautogui.hotkey('enter')

# Aba "Carga"
pyautogui.moveTo(498, 341)
pyautogui.click()

# Produto Predominante
pyautogui.moveTo(246, 416)
pyautogui.doubleClick()
pyautogui.typewrite('IOGURTE')

# Volume Carga
pyautogui.moveTo(225, 458)
pyautogui.click()

# Modal
pyautogui.moveTo(545, 340)
pyautogui.click()

pyautogui.moveTo(74, 419)
pyautogui.click()
pyautogui.typewrite('a')
pyautogui.sleep(1)
pyautogui.moveTo(100, 523)
pyautogui.click()

pyautogui.moveTo(252, 415)
pyautogui.click()
pyautogui.typewrite('a')
pyautogui.sleep(1)
pyautogui.hotkey('enter')

pyautogui.scroll(-250)

pyautogui.moveTo(230, 629)
pyautogui.click()
pyautogui.sleep(5)

pyautogui.moveTo(476, 632)
pyautogui.sleep(2)
pyautogui.click()

# Espera aparecer a tela do safenet
pyautogui.sleep(23)

#Move até o safenet
pyautogui.moveTo(657, 404)
pyautogui.click()
pyautogui.typewrite("senha")
pyautogui.moveTo(812, 480)
pyautogui.click()


driver.quit()