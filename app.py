import openpyxl
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participação = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value

    fonte_name = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf',60)
    fonte_datas = ImageFont.truetype('./tahoma.ttf',50)

    image = Image.open('./Certificado_Exemplo.png')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((300,500), nome_participante, fill='white', font=fonte_name)
    desenhar.text((810,645), tipo_participação, fill='white', font=fonte_geral)
    desenhar.text((550,750), nome_curso, fill='white', font=fonte_geral)
    desenhar.text((450,847), str(data_inicio), fill='white', font=fonte_datas)
    desenhar.text((525,919), str(data_final), fill='white', font=fonte_datas)
    desenhar.text((710,1208), str(f'{carga_horaria} horas'), fill='white', font=fonte_geral)
    desenhar.text((1445,1280), str(data_emissao), fill='white', font=fonte_datas)

    image.save(f'./certificados_gerados/{indice} {nome_participante} certificado.png')