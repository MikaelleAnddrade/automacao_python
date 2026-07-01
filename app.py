"""
AUTOMATIZAR MENSAGENS DE COBRANÇAS PARA OS CLIENTES
"""
import openpyxl
from urllib.parse import quote
import webbrowser 


webbrowser.open('https://web.whatsapp.com/')


# Ler planilha e guardar informações

workbook = openpyxl.load_workbook('Book 2.xlsx')

pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

# Criar link personalizados do whatsapp e enviar mensagens para cada cliente
#cliente com base nos dados da planilha.

    mensagem = f'Olá {nome} seu pedido vence no dia {vencimento.strftime('%d/%m/%Y, %H:%M:%S')}. Favor pagar o link a seguir (inserirLink)'
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}' 
    webbrowser.open(link_mensagem_whatsapp)

    input('')

