# Instalando biblioteca 
# pip intall pyautogui
# pip intall pandas
# pip intall openpyxl


# Importando Bibliotecas
import pyautogui
import time
import pandas
import openpyxl

# Pause a aplicação por 0.5 segundos

pyautogui.PAUSE = 0.5

# Passo 1 -  Abrir Sistema da Empresa

pyautogui.press('win')
pyautogui.typewrite('edge')
pyautogui.press('enter')
pyautogui.typewrite('https://dlp.hashtagtreinamentos.com/python/intensivao/login')
pyautogui.press('enter')

# Espera programada

time.sleep(3)

# Passo 2 - Login do sistema (https://dlp.hashtagtreinamentos.com/python/intensivao/login)

pyautogui.click(866, 363)
pyautogui.typewrite('dev.rodrigopires@gmail.com')
pyautogui.press('tab')
pyautogui.typewrite('123456')
pyautogui.press('enter')

# Passo 3 - Acessar o menu de funcionários

tabela = pandas.read_csv("produtos.csv")
print(tabela)

time.sleep(2)
# Passo 4 - Cadastrar produtos

for linha in tabela.index:
    pyautogui.click(757, 244)

    # código
    codigo = tabela.loc[linha, "codigo"]
    pyautogui.typewrite(str(codigo))
    pyautogui.press('tab')

    # Marca
    marca = tabela.loc[linha, "marca"]
    pyautogui.typewrite(str(marca))
    pyautogui.press('tab')

    # Tipo
    tipo = tabela.loc[linha, "tipo"]
    pyautogui.typewrite(str(tipo))  
    pyautogui.press("tab")

    # Categoria
    categoria = tabela.loc[linha, "categoria"]
    pyautogui.typewrite(str(categoria))
    pyautogui.press("tab")

    # Preço Unitário
    preco_unitario = tabela.loc[linha, "preco_unitario"]
    pyautogui.typewrite(str(preco_unitario))
    pyautogui.press("tab")

    # Custo
    custo = tabela.loc[linha, "custo"]
    pyautogui.typewrite(str(custo))
    pyautogui.press("tab")

    # Obs
    obs = str(tabela.loc[linha, "obs"])
    if obs != "nan":
        pyautogui.typewrite(obs)
    pyautogui.press("tab")

    # Salvar
    pyautogui.press("enter")

# Passo 5 - Cadastrar produtos