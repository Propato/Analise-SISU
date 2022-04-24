import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.comments import Comment
import string

def cores(wb, dados, curso, regiao, corte, nota):

    ws0 = wb['Dicionário de dados']
    ws1 = wb[curso]

    #define cores
    red = PatternFill(patternType='solid', fgColor='ff2800')
    green = PatternFill(patternType='solid', fgColor='5cb800')
    blue = PatternFill(patternType='solid', fgColor='39A7FA')
    lilac = PatternFill(patternType='solid', fgColor='00CCCCFF')

    for i in range(dados.shape[0]):

        if (dados.iloc[i]['DS_REGIAO']).lower() == regiao:
            if corte and nota:
                if ws1[f'Y{i+2}'].value:
                    j=0
                    for A in list(string.ascii_uppercase):
                        ws1[f'{A}{i+2}'].fill = lilac
                        j+=1
                        if j>=dados.shape[1]:
                            break
                    if not j>=dados.shape[1]:
                        for A in list(string.ascii_uppercase):
                            ws1[f'A{A}{i+2}'].fill = lilac
                            j+=1
                            if j>=dados.shape[1]:
                                break
            else:
                j=0
                for A in list(string.ascii_uppercase):
                    ws1[f'{A}{i+2}'].fill = lilac
                    j+=1
                    if j>=dados.shape[1]:
                        break

            ws1[f'G{i+2}'].fill = blue

        if corte and nota:
            if ws1[f'Y{i+2}'].value:
                ws1[f'Y{i+2}'].fill = green
                ws1[f'C{i+2}'].fill = green
            else:
                ws1[f'Y{i+2}'].fill = red
                ws1[f'C{i+2}'].fill = red

    print('Cores processadas.')

def comment(corte, nota, igc):
    comentarios = ['Código da Instituição de Ensino Superior', 'Nome da Instituição de Ensino Superior', 'Sigla da Instituição de Ensino Superior', 'Nome do Campus', 'Nome do município do Campus', 'Estado da instituição', 'Região da instituição', 'Código do curso', 'Nome do curso', 'Grau do curso', 'Turno do curso', 'Periodicidade da abertura de vagas através do SISU', 'Número de vagas totais', 'Número de vagas na cota nesta edição']
    
    if corte[0]:
        comentarios = comentarios + [f'Número de vagas na cota em {corte[1]}/{corte[2]}', f'Número de inscritos totais na cota em {corte[1]}/{corte[2]}']

    comentarios = comentarios + ['Tipo de cota', 'Peso da Redação', 'Peso de linguagens', 'Peso de matemática', 'Peso das ciências humanas', 'Peso das ciências da natureza']
    
    if corte[0]:
        comentarios.append(f'Nota de corte para a chamada regular na cota em {corte[1]}/{corte[2]}')

    if nota:
        comentarios.append('Suas notas ponderadas')
        if corte[0]:
            comentarios.append(f'Aprovação na vaga, se as suas notas são maiores (VERDADEIRO) ou menores (FALSO) que a nota de corte de {corte[1]}/{corte[2]}')
        
    if igc:
        comentarios = comentarios + ['Nota da avaliação do IGC (Contínuo) de acordo com o MEC', 'Nota da avaliação do IGC (Faixa) de acordo com o MEC']
    
    return comentarios


def layout(wb, dados, dic_dados, curso, corte, nota, igc):

    ws0 = wb['Dicionário de dados']
    ws1 = wb[curso]

    borda = Border(right=Side(border_style='thin', color='00000000'), bottom=Side(border_style='thin', color='00000000'))
    alinhamento = Alignment(horizontal='center', vertical='center')

    comentarios = comment(corte, nota, igc)

    j=0
    for A in list(string.ascii_uppercase):
    
        if type(ws1[f'{A}2'].value) == float or type(ws1[f'{A}2'].value) == int:
            ws1.column_dimensions[A].width = 12
        else:
            ws1.column_dimensions[A].width = 20

        ws1.column_dimensions['B'].width = 42
        ws1.column_dimensions['D'].width = 30

        h = len(comentarios[j])*2
        compri = len(comentarios[j])*4

        if h < 50:
            h=50
        if h>110:
            h=110
        if compri < 130:
            compri=130

        ws1[f'{A}1'].comment = Comment(comentarios[j], 'David Propato', h, compri)
    
        for i in range(1, dados.shape[0]+2):
            ws1[f'{A}{i}'].border = borda
            ws1[f'{A}{i}'].alignment = alinhamento

        j+=1
        if j>=dados.shape[1]:
            break
    
    if not j>=dados.shape[1]:
        for A in list(string.ascii_uppercase):
    
            if type(ws1[f'A{A}2'].value) == float or type(ws1[f'A{A}2'].value) == int:
                ws1.column_dimensions[f'A{A}'].width = 12
            else:
                ws1.column_dimensions[f'A{A}'].width = 20

            h = len(comentarios[j])*2
            compri = len(comentarios[j])*4

            if h < 50:
                    h=50
            if h>110:
                h=110
            if compri < 130:
                compri=130

            ws1[f'A{A}1'].comment = Comment(comentarios[j], 'David Propato', h, compri)

            for i in range(1, dados.shape[0]+2):
                ws1[f'A{A}{i}'].border = borda
                ws1[f'A{A}{i}'].alignment = alinhamento

            j+=1
            if j>=dados.shape[1]:
                break

    alin = Alignment(horizontal='left', vertical='distributed')
    borda = Border(right=Side(border_style='thin', color='00000000'))

    for i in range(1, dic_dados.shape[0]+2):
        ws0[f'A{i}'].alignment = alinhamento
        ws0[f'B{i}'].alignment = alin

        if i > 1:
            ws0[f'A{i}'].border = borda
            ws0[f'B{i}'].border = borda

        h = int(len(ws0[f'B{i}'].value)/183)

        if len(ws0[f'B{i}'].value)%183:
            h+=1
        '    '
        ws0.row_dimensions[i].height = 15*h

    ws0[f'A{dic_dados.shape[0]+2}'].border = Border(top=Side(border_style='thin', color='00000000'))
    ws0[f'B{dic_dados.shape[0]+2}'].border = Border(top=Side(border_style='thin', color='00000000'))

    ws0.column_dimensions['A'].width = 32
    ws0.column_dimensions['B'].width = 165

    print('Configurações de layout processadas.')