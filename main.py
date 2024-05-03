from docxtpl import DocxTemplate
import openpyxl

# Carregar o modelo Word
modelo = DocxTemplate('template_certificado.docx')

# Carregar os dados do Excel
wb = openpyxl.load_workbook('dados.xlsx')
planilha = wb['Planilha1']  # nome da planilha onde estão os dados

# Itera sobre as linhas da planilha, começando da segunda linha para evitar o cabeçalho
for col in planilha.iter_rows(min_row=2, values_only=True):
    dados = {
        'TITULO': col[0], # coluna A
        'AUTOR1': col[1], # coluna B
        'AUTOR2': col[2], # coluna C
        'MODALIDADE': col[3], # coluna D
        'EVENTO': col[4], # coluna E
        'ORGANIZADOR': col[5], # coluna F
        'DATA_EVENTO': col[6], # coluna G
        'DATA_EMISSAO': col[7] # coluna H
    }

    # Renderiza o modelo Word com os dados da planilha
    modelo.render(dados)

    # Salva o documento preenchido
    arquivo = 'arquivos/' + linha[1] + '_' + linha[2] + '.docx'
    
    modelo.save(arquivo)
    print('.', end='')
print('\nArquivos salvos com sucesso!')
