import pyautogui
import pyperclip
import openpyxl
X_BROWSER = 1007
Y_BROWSER = 1038
link = "https://cadastroprodutos-devaprender.netlify.app/"


#def para abrir o navegador
def OPEN_BROWSER():
    pyautogui.click(X_BROWSER, Y_BROWSER)

#def para abrir uma guia, colar o link e entrar no site

def PASTE_LINK():
    pyautogui.hotkey('ctrl','n')
    pyperclip.copy(link)
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')

#def para carregar a planilha

def LOAD_WORKBOOK():
    global sheet_produtos
    workbook = openpyxl.load_workbook(r'C:\Users\crist\OneDrive\Área de Trabalho\python\bootcamp\aula02\produtos.xlsx')
    sheet_produtos = workbook['produtos']


#def para carregar os dados, colar os dados nos campos respectivos e registrar o produto

def PASTE_VALUES():
    for linha in sheet_produtos.iter_rows(min_row = 2, max_row = 501):
        produto = linha[0].value
        fornecedor = linha[1].value
        categoria = linha[2].value
        valor_unitario = linha[4].value
        notificar_venda = linha[5].value
        pyautogui.click(706, 458, duration = 0.4)
        pyautogui.write(produto)
        pyautogui.click(1035,464, duration = 0.4)
        pyautogui.write(fornecedor)
        pyautogui.click(737, 577, duration = 0.4)
        pyautogui.write(categoria)
        pyautogui.click(1104, 583, duration = 0.4)
        pyautogui.write(valor_unitario)
        if notificar_venda == 'Sim':
                pyautogui.click(621, 689)
        else:
                pyautogui.click(756,684)
                pyautogui.click(721,773)
                pyautogui.sleep(0.8)
                pyautogui.click(1185,200)
