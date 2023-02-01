import pyautogui
import openpyxl
import pyperclip
# ATENÇÃO
# Para o funciomaneto correto o site deve estar do tamanho
# antes do container cinza claro ficar fluid 

# 2 - abrir a planilha
workbook = openpyxl.load_workbook(r'C:\Users\rhuan\OneDrive\Área de Trabalho\PROJETOS\Curso Python\Automação Produtos para site\produtos.xlsx')
sheet_produtos = workbook['produtos']
for linha in sheet_produtos.iter_rows(min_row=2,max_row=501):
    produto = linha[0].value
    fornecedor = linha[1].value
    categoria = linha[2].value
    quantidade = linha[3].value
    valor_unitario = linha[4].value
    notificar_venda = linha[5].value
    # colar dados campo produto
    pyautogui.click(1507,360,duration=1)
    pyautogui.write(produto)
    # colar dados campo fornecedor
    pyautogui.click(1538,487,duration=1)
    pyautogui.write(fornecedor)
    # colar dados campo categoria
    pyautogui.click(1735,360,duration=1)
    pyperclip.copy(categoria)
    pyautogui.hotkey('ctrl','v')
    # colar dados campo valor unitário
    pyautogui.click(1714,490,duration=1)
    pyperclip.copy(valor_unitario)
    pyautogui.hotkey('ctrl','v')
    # se notificar venda for igual a sim, marcar sim
    # se notificar venda for igual a não, marcar não
    if notificar_venda == "Sim":
        pyautogui.click(1621,633, duration=1)
    elif notificar_venda == "Não":
        pyautogui.click(1673,633, duration=1)
    # clicar em registrar produto
    pyautogui.click(1616,675, duration=1)
    # clicar em ok, na mensagem de cadastro com sucesso
    pyautogui.click(1799,165,duration=1)

# 3 -  após ter aberto a planilha eu devo copiar o dado que está dentro do campo produto e colar no seu campo respectivo dentro do sistema web
# 4 - Repetir o mesmo processo para as outras colunas
# 5 - Repetir até chegar ao último cadastro da planilha com os 500 cadastros