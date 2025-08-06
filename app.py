import openpyxl
import pyperclip
import pyautogui
import time 

workbook = openpyxl.load_workbook('Planilha_Produtoss.xlsx')
sheet_produtos = workbook ['Planilha1']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1225,352,duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1224,484,duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1227,617,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1215,733,duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1212,866,duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1209,986,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.scroll(-100)

    pyautogui.click(1364,957,duration=1)
    time.sleep(3)

    pyautogui.scroll(100)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1225,352,duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1224,484,duration=1)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1227,617,duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1215,733,duration=1)
    pyautogui.hotkey('ctrl','v')


    tamanho = linha[10].value
    pyautogui.click(1373,866,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(1265,756,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1272,784,duration=1)
    else:
        pyautogui.click(1281,823,duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1209,986,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.scroll(-100)

    pyautogui.click(1364,957,duration=1)
    time.sleep(3)

    pyautogui.scroll(100)
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1299,362,duration=1)
    pyautogui.hotkey('ctrl','v')


    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1292,489,duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1295,607,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barra = linha[15].value
    pyperclip.copy(codigo_de_barra)
    pyautogui.click(1291,725,duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1269,857,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1465,941,duration=1)
    time.sleep(3)
    
    pyautogui.click(1501,582,duration=1)
    

