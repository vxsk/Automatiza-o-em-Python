import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na panilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informaçao de um campo e colar no campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1483,345, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1460,436, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1466,553, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    cod_produto = linha[3].value
    pyperclip.copy(cod_produto)
    pyautogui.click(1543,658, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1469,733, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1494,817, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    # Repetir esses passos para outros campos ate preencher campos daquela pagina
    
    pyautogui.click(1425,874, duration=1)   # Clicar em proximo
    sleep(3)
    
    # Repetir os mesmos passos e ir para proxima pagina(pagina2)
    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1482,366, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    qtd = linha[7].value
    pyperclip.copy(qtd)
    pyautogui.click(1486,457, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    validade = linha[8].value
    pyperclip.copy(validade)
    pyautogui.click(1480,545, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1480,612, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    #Ler info da planilha
    # Se for "Pequeno", clicar em uma pos
    # Se for "Medio", clicar em uma pos
    # Se for "Grande", clicar em uma pos
    
    
    tamanho = linha[10].value
    pyautogui.click(1563,710, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(1462,740, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1440,763, duration=1)
    else:
        pyautogui.click(1448,779, duration=1)
        
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1492,802, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    pyautogui.click(1438,854, duration=1)  # Clicar em proximo
    sleep(3)
    
    # Repetir os mesmos passaos e finalizar o cadastro daquele produto e clicar em concluir
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1510,386, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    origem = linha[13].value
    pyperclip.copy(origem)
    pyautogui.click(1482,463, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    obs = linha[14].value
    pyperclip.copy(obs)
    pyautogui.click(1485,582, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    cod_barras = linha[15].value
    pyperclip.copy(cod_barras)
    pyautogui.click(1488,687, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.click(1512,784, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    #Clicar em concluir 
    pyautogui.click(1432,838, duration=1)
    
    #Clicar em Ok
    pyautogui.click(1835,167, duration=1)
    
     #Clicar em Ok(2)
    pyautogui.click(1826,164, duration=1)
    
    #Clicar em Add mais um
    pyautogui.click(1659,618, duration=1)





# clicar em ok, para finalizar o processo 
# clicar em ok novamente na mensangem do banco de dados
# clicar em adicionar mais um e repetir o processo ate finalizar 
  