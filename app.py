#Pega os dados da planilha 
import openpyxl 
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    #Acessar cada celular que contem informações 
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value


    fonte_nome = ImageFont.truetype('tahomabd.ttf', 90)
    font_geral = ImageFont.truetype('tahoma.ttf', 80)
    font_data = ImageFont.truetype('tahoma.ttf', 55)

    image = Image.open('certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)


    desenhar.text((1020,827),nome_participante, fill='black', font=fonte_nome)
    desenhar.text((1060, 950), nome_curso, fill='black', font=font_geral)
    desenhar.text((1435, 1065), tipo_participacao, fil='black', font=font_geral)
    desenhar.text((1480, 1182), str(carga_horaria), fill='black', font=font_geral)

    desenhar.text((750, 1770), data_inicio, fill='blue', font=font_data)
    desenhar.text((750, 1930), data_final, fill='blue', font=font_data)

    desenhar.text((2220, 1930), data_emissao, fill='blue', font=font_data)
    image.save(f'{indice} {nome_participante} certificado.png')