# Pegar os dados da planilha

import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Abrir a planilha

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for linha in sheet_alunos.iter_rows(min_row=2):
    # cada celula que contém a info que precisamos
    nome_curso = linha[0].value # Nome do curso
    nome_participante = linha[1].value # nome do participante
    tipo_paticipacao = linha[2].value # tipo de participação
    data_inicio = linha[3].value # data de inicio
    data_final = linha[4].value # data final
    carga_horaria = linha[5].value # carga horaria
    data_emissao = linha[6].value # data de emissão
    
    # Transferir os dados da planilha para a imagem do certificado

    font_nome = ImageFont.truetype('./tahomabd.ttf')
    font_geral = ImageFont.truetype('./tahoma.ttf')

    imagem = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(imagem)

    desenhar.text((200,600),nome_participante,fill='black',font=font_nome)

    imagem.save('./test.png')



