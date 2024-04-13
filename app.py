# Pegar os dados da planilha

import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Abrir a planilha

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # cada celula que contém a info que precisamos
    nome_curso = linha[0].value # Nome do curso
    nome_participante = linha[1].value # nome do participante
    tipo_paticipacao = linha[2].value # tipo de participação
    data_inicio = linha[3].value # data de inicio
    data_final = linha[4].value # data final
    carga_horaria = linha[5].value # carga horaria
    data_emissao = linha[6].value # data de emissão
    
    # Transferir os dados da planilha para a imagem do certificado

    font_nome = ImageFont.truetype('./tahomabd.ttf',90)
    font_geral = ImageFont.truetype('./tahoma.ttf',80)
    font_data = ImageFont.truetype('./tahoma.ttf',55)

    imagem = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(imagem)

    desenhar.text((1020,827),nome_participante,fill='black',font=font_nome)
    desenhar.text((1060,950),nome_curso,fill='black',font=font_geral)
    desenhar.text((1435,1065),tipo_paticipacao,fill='black',font=font_geral)
    desenhar.text((1480,1182),nome_curso,fill='black',font=font_geral)

# datas

    desenhar.text((750,1770),data_inicio,fill='blue',font=font_data)
    desenhar.text((750,1930),data_final,fill='blue',font=font_data)

    desenhar.text((2220,1930),data_emissao,fill='blue',font=font_data)




    imagem.save(f'./certificados completos/{indice} {nome_participante} certificado.png')



