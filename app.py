import pyautogui
from actions import OPEN_BROWSER, PASTE_LINK,LOAD_WORKBOOK, PASTE_VALUES
'''
Sequência de passos:

1 - Abrir o navegador.

2 - Entrar no site "https://cadastroprodutos-devaprender.netlify.app/".

3 - Abrir a planilha "produtos.xlsx".

4 - Copiar os dados dos campos respectivos e colar no formulário do site.

5 - Repetir o processo para cada linha da planilha até o final.
'''

#-Abrindo o navegador
OPEN_BROWSER()
pyautogui.sleep(1.5)
#-Abrindo nova aba, colando o link destino e entrando no site
PASTE_LINK()
pyautogui.sleep(1.5)
#-Carregando a planilha e os valores
LOAD_WORKBOOK()
pyautogui.sleep(1.5)
#-Colando os valores nos campos respectivos e cadastrando
PASTE_VALUES()


