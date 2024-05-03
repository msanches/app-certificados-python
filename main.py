#from datetime import datetime
from docxtpl import DocxTemplate
import openpyxl

# Carregar o modelo Word
modelo = DocxTemplate('template_certificado.docx')

# Carregar os dados do Excel
wb = openpyxl.load_workbook('dados.xlsx')
planilha = wb['Planilha1']  # nome da planilha onde estão os dados

#hoje = datetime.now().strftime('%d/%m/%Y')

# Itera sobre as linhas da planilha, começando da segunda linha para evitar o cabeçalho
for linha in planilha.iter_rows(min_row=2, values_only=True):
    dados = {
        'TITULO': linha[0],
        'AUTOR1': linha[1],
        'AUTOR2': linha[2],
        'MODALIDADE': linha[3],
        'EVENTO': linha[4],
        'ORGANIZADOR': linha[5],
        'DATA_EVENTO': linha[6],
        'DATA_EMISSAO': linha[7]
        # adicione mais marcadores conforme necessário
    }

    # Renderizar o modelo Word com os dados
    #modelo = modelo_1 if linha[3] == 'xxx' else modelo_2
    modelo.render(dados)

    # Salvar o documento preenchido
    #arquivo = 'arquivos/' + linha[3] + '/' + linha[0].replace('/', '_') + '.docx'
    arquivo = 'arquivos/' + linha[1] + '_' + linha[2] + '.docx'
    
    modelo.save(arquivo)
    print('.', end='')
print('\nArquivos salvos com sucesso!')
