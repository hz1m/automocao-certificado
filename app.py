from PIL import Image, ImageDraw, ImageFont
import openpyxl

# ABRIR PLANILHA
planilha_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = planilha_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2, max_row=3)):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value
    input('')

    # TRANSFERIR DADOS DA PLANILHA PARA IMAGEM DO CERTIFICADO
    imagem = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(imagem)
    
    tamanho_fonte_nome = 90
    tamanho_fonte_geral = 90
    tamanho_fonte_data = 55

    fonte_nome = ImageFont.truetype('arial.ttf', tamanho_fonte_nome)
    fonte_geral = ImageFont.truetype('arial.ttf', tamanho_fonte_geral)
    fonte_data = ImageFont.truetype('arial.ttf', tamanho_fonte_data)

    desenhar.text((1020,827), nome_participante, fill='black', font=fonte_nome)
    desenhar.text((1060,950), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1435,1065), tipo_participacao, fill='black', font=fonte_geral)
    desenhar.text((1480,1182), str(carga_horaria), fill='black', font=fonte_geral)
    desenhar.text((750, 1770), data_inicio, fill='black', font=fonte_geral)
    desenhar.text((750, 1930), data_final, fill='black', font=fonte_geral)
    desenhar.text((2220, 1930), data_emissao, fill='black', font=fonte_geral)

    imagem.save(f'./{indice}_{nome_participante}_certificado.png')
