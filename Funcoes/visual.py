from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.comments import Comment
import string

def cores(wb, df, curso, regiao):
    aprovado = 0
    tentativas = 0

    ws = wb[curso]

    #define cores
    red = PatternFill(patternType='solid', fgColor='ff2800')
    green = PatternFill(patternType='solid', fgColor='5cb800')
    blue = PatternFill(patternType='solid', fgColor='39A7FA')
    lilac = PatternFill(patternType='solid', fgColor='00CCCCFF')

    for i in range(df.shape[0]):

        if (df.iloc[i]['DS_REGIAO']).lower() == regiao:
            if ('NU_NOTACORTE' in df.columns and 'NOTAS' in df.columns):
                if ws[f'X{i+2}'].value:
                    j=0
                    for A in list(string.ascii_uppercase):
                        ws[f'{A}{i+2}'].fill = lilac
                        j+=1
                        if j>=df.shape[1]:
                            break
                    if not j>=df.shape[1]:
                        for A in list(string.ascii_uppercase):
                            ws[f'A{A}{i+2}'].fill = lilac
                            j+=1
                            if j>=df.shape[1]:
                                break
            else:
                j=0
                for A in list(string.ascii_uppercase):
                    ws[f'{A}{i+2}'].fill = lilac
                    j+=1
                    if j>=df.shape[1]:
                        break

            ws[f'G{i+2}'].fill = blue

        if ('NU_NOTACORTE' in df.columns and 'NOTAS' in df.columns):
            tentativas+=1
            if ws[f'X{i+2}'].value:
                ws[f'X{i+2}'].fill = green
                ws[f'C{i+2}'].fill = green
                aprovado+=1
            else:
                ws[f'X{i+2}'].fill = red
                ws[f'C{i+2}'].fill = red
    
    if ('NU_NOTACORTE' in df.columns and 'NOTAS' in df.columns):
        ws['X1'].value = f'APROVAÇÃO ({aprovado/tentativas:.2%})'

    print('Cores processadas.')

def comment(corte, nota, notas):
    comentarios = ['Código da Instituição de Ensino Superior', 'Nome da Instituição de Ensino Superior', 'Sigla da Instituição de Ensino Superior', 'Nome do Campus', 'Estado da Instituição de Ensino Superior','Município do Campus',  'Região da Instituição de Ensino Superior', 'Código do curso', 'Nome do curso', 'Grau do curso', 'Turno do curso', 'Periodicidade da abertura de vagas através do SISU', 'Número de vagas totais', 'Número de vagas na cota']
    
    if corte: comentarios.append('Número de inscritos totais na cota')

    comentarios.append('Tipo de cota')
    if nota:
        comentarios = comentarios + [f'Peso da Redação\nNota: {notas[0]}', f'Peso de linguagens\nNota: {notas[1]}', f'Peso de matemática\nNota: {notas[2]}', f'Peso das ciências humanas\nNota: {notas[3]}', f'Peso das ciências da natureza\nNota: {notas[4]}']
    else:
        comentarios = comentarios + ['Peso da Redação', 'Peso de linguagens', 'Peso de matemática', 'Peso das ciências humanas', 'Peso das ciências da natureza']
    
    if corte: comentarios.append(f'Nota de corte para a chamada regular na cota')

    if nota:
        comentarios.append('Suas notas ponderadas')
        if corte:
            comentarios.append(f'Aprovação na vaga, se as suas notas são maiores (VERDADEIRO) ou menores (FALSO) que a nota de corte')
    
    return comentarios


def layout(wb, df, dic_df, curso, notas):

    ws0 = wb['Dicionário de dados']
    ws1 = wb[curso]

    borda = Border(right=Side(border_style='thin', color='00000000'), bottom=Side(border_style='thin', color='00000000'))
    alinhamento = Alignment(horizontal='center', vertical='center')

    comentarios = comment(('NU_NOTACORTE' in df.columns), ('NOTAS' in df.columns), notas)

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
    
        for i in range(1, df.shape[0]+2):
            ws1[f'{A}{i}'].border = borda
            ws1[f'{A}{i}'].alignment = alinhamento

        j+=1
        if j>=df.shape[1]:
            break
    
    if not j>=df.shape[1]:
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

            for i in range(1, df.shape[0]+2):
                ws1[f'A{A}{i}'].border = borda
                ws1[f'A{A}{i}'].alignment = alinhamento

            j+=1
            if j>=df.shape[1]:
                break

    alin = Alignment(horizontal='left', vertical='distributed')
    borda = Border(right=Side(border_style='thin', color='00000000'))

    for i in range(1, dic_df.shape[0]+2):
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

    ws0[f'A{dic_df.shape[0]+2}'].border = Border(top=Side(border_style='thin', color='00000000'))
    ws0[f'B{dic_df.shape[0]+2}'].border = Border(top=Side(border_style='thin', color='00000000'))

    ws0.column_dimensions['A'].width = 32
    ws0.column_dimensions['B'].width = 165

    print('Configurações de layout processadas.')