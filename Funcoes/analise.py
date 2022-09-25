import pandas as pd
import openpyxl
import sys
import Funcoes.visual as visual

def abre_excel(ano, semestre):
    try:
        dic_dados = pd.read_excel(f'Dados/Vagas ofertadas/Portal Sisu_Sisu {ano}-{semestre}_Vagas ofertadas.xlsx', sheet_name=0)
        dados = pd.read_excel(f'Dados/Vagas ofertadas/Portal Sisu_Sisu {ano}-{semestre}_Vagas ofertadas.xlsx', sheet_name=1, usecols='C:E, H:P, R, S, V, W, Y, AA, AC, AE')    
    except Exception as e:
        print(e)
        print('Erro ao abrir arquivo.')
        sys.exit()
    
    try:
        dic_cortes = pd.read_excel(f'Dados/Inscrições e notas de corte/Portal Sisu_Sisu {ano}-{semestre}_Inscrições e notas de corte.xlsx', sheet_name=0)
        cortes = pd.read_excel(f'Dados/Inscrições e notas de corte/Portal Sisu_Sisu {ano}-{semestre}_Inscrições e notas de corte.xlsx', sheet_name=1, usecols='L, M, Q, T, U')
    
        return dic_dados, dados, dic_cortes, cortes
    except Exception as e:
        print(e)
        return dic_dados, dados, False, False 

def gera_relatorio(df, dic_df, ano, semestre, curso, regiao, notas):
    writer = pd.ExcelWriter(f'Relatorios/{curso}_{ano}_{semestre}.xlsx', engine='openpyxl')

    dic_df.to_excel(writer, index=False, sheet_name='Dicionário de dados')
    df.to_excel(writer, index=False, sheet_name=curso)

    writer.save()

    wb = openpyxl.load_workbook(f"Relatorios/{curso}_{ano}_{semestre}.xlsx")
    ws0 = wb['Dicionário de dados']
    
    ### SIMPLIFICA NOME PARA OS PESOS NO DIC ###
    areas = ['REDACAO', 'LINGUAGENS', 'MATEMATICA', 'CIENCIAS_HUMANAS', 'CIENCIAS_NATUREZA']
    j=0
    for i in range(1, dic_df.shape[0]+2):
        if areas[j] in ws0[f'A{i}'].value:
            ws0[f'A{i}'].value = areas[j]
            j+=1
        if j>=5:
            break

    ### ATUALIZA O VISUAL DA TABELA NO EXCEL ###
    visual.cores(wb, df, curso, regiao)
    visual.layout(wb, df, dic_df, curso, notas)

    wb.save(f'Relatorios/{curso}_{ano}_{semestre}.xlsx')

def media_ponderada(notas, pesos):
    soma_pesos = 0
    soma_notas = 0

    for i in range(len(pesos)):
        soma_pesos += pesos[i]
        soma_notas += notas[i]*pesos[i]
    return round(soma_notas/soma_pesos, 2)

def insere_medias(df, dic_df, notas):
    df['NOTAS'] = (media_ponderada(notas, (df['REDACAO'], df['LINGUAGENS'], df['MATEMATICA'], df['CIENCIAS_HUMANAS'], df['CIENCIAS_NATUREZA']))).to_frame()

    dic_df = pd.concat([dic_df, pd.DataFrame([['NOTAS', 'Médias ponderadas para cada curso']],columns=[dic_df.columns[0], dic_df.columns[1]])])
    
    if ('NU_NOTACORTE' in df.columns and 'NOTAS' in df.columns):
        df['APROVAÇÃO'] = ((df['NU_NOTACORTE'] - df['NOTAS'])<0).to_frame()

        dic_df = pd.concat([dic_df, pd.DataFrame([['APROVAÇÃO', 'Indica se as suas notas são maiores (escrito VERDADEIRO) ou menores (escrito FALSO) que a nota de corte']],columns=[dic_df.columns[0], dic_df.columns[1]])])
    
    return df, dic_df

def limpa_dic(dic, df):
    i=0
    while i < dic.shape[0]:
        if dic.iloc[i]['Nome da coluna'] not in df.columns:
            dic = dic.drop(index=i)
            dic.reset_index(drop=True, inplace=True)
            i-=1
        i+=1
    return dic

def filtra_excel(dados, cortes, curso):
    
    cotas_invalidas = ['autodeclarados', 'pretos', 'negros', 'pardos', 'índios', 'indígenas', 'indígena,', 'quilombolas', 'transexuais', 'vulnerabilidade', 'deficiência', 'necessidades', 'carência', 'inferior', 'regional', 'região', 'microrregiões', 'mesorregiões', 'membros de comunidade', 'residentes', 'residem', 'no estado de pernambuco', 'localizadas', 'baixa', 'natal', 'de até um salário-mínimo']

    dados = dados.loc[dados['NO_CURSO']==curso].reset_index(drop=True)
    if not type(cortes) == bool: cortes = cortes.loc[curso==cortes['NO_CURSO']].reset_index(drop=True)

    ### CHECAGEM - CURSO ###
    if (dados.shape[0] == 0) or (not(type(cortes) == bool) and cortes.shape[0] == 0):
        print(f'Curso {curso.upper()} inválido, sem ofertas de vagas ou pesos.')
        sys.exit()

    ### FILTRA POR COTAS ###
    if not type(cortes) == bool:
        i=0
        while i < cortes.shape[0]:
            for j in cotas_invalidas:
                if(j in cortes.iloc[i]['DS_MOD_CONCORRENCIA'].lower()):
                    cortes = cortes.drop(index=i)
                    cortes.reset_index(drop=True, inplace=True)
                    i-=1
                    break
            i+=1

    i=0
    while i < dados.shape[0]:
        for j in cotas_invalidas:
            if(j in dados.iloc[i]['DS_MOD_CONCORRENCIA'].lower()):
                dados = dados.drop(index=i)
                dados.reset_index(drop=True, inplace=True)
                i-=1
                break
        i+=1

    if not type(cortes) == bool:
        dados.sort_values(by='DS_MOD_CONCORRENCIA', kind='stable', inplace=True, ignore_index=True)
        dados.sort_values(by='CO_IES_CURSO', kind='stable', inplace=True, ignore_index=True)
        
        cortes.sort_values(by='DS_MOD_CONCORRENCIA', kind='stable', inplace=True, ignore_index=True)
        cortes.sort_values(by='CO_IES_CURSO', kind='stable', inplace=True, ignore_index=True)
        
        ### CHECAGEM - COTAS ###
        if not ((dados['CO_IES_CURSO'] == cortes['CO_IES_CURSO']).all() and (dados['DS_MOD_CONCORRENCIA'] == cortes['DS_MOD_CONCORRENCIA']).all()):
            print('Dados dos cursos e cortes não batem, favor checar as planilhas e tentar novamente.')
            sys.exit()

        ### INSERE DADOS DE INSCRIÇÕES NO DF PRINCIPAL ###
        dados.insert(14, 'QT_INSCRICAO',  cortes['QT_INSCRICAO'])
        dados.insert(21, 'NU_NOTACORTE',  cortes['NU_NOTACORTE'])

    return dados, cortes

def dados_sisu(curso, ano, semestre, regiao, notas):
    print('='*127)
    print('=Executando código para processar dados do SISU de 2022 a respeito do curso desejado para as cotas de ampla e escolas públicas=')
    print('='*127)

    dic_dados, dados, dic_cortes, cortes = abre_excel(ano, semestre)

    print('Tabelas lida.')

    dados, cortes = filtra_excel(dados, cortes, curso)

    ### REORGANIZA DIC_DADOS ###
    dic_dados = limpa_dic(dic_dados, dados)
    ### REORGANIZA DIC_CORTES ###
    if not type(cortes) == bool: 
        dic_cortes = limpa_dic(dic_cortes, cortes)

        for i in range(3):
            dic_cortes = dic_cortes.drop(index=0)
            dic_cortes.reset_index(drop=True, inplace=True)

        ### FINALIZA DIC_DADOS ###
        dic_dados1 = dic_dados[:14]
        dic_dados2 = dic_dados[14:]

        dic_dados1 = pd.concat([dic_dados1, dic_cortes[1:]])
        dic_dados = pd.concat([dic_dados1, dic_dados2])
        dic_dados = pd.concat([dic_dados, dic_cortes[:1]])

        dic_dados.reset_index(drop=True, inplace=True)

    ### SIMPLIFICA NOME PARA COLUNA DOS PESOS ###
    dados.rename({'PESO_REDACAO': 'REDACAO'}, axis = 1, inplace=True)
    dados.rename({'PESO_LINGUAGENS': 'LINGUAGENS'}, axis = 1, inplace=True)
    dados.rename({'PESO_MATEMATICA': 'MATEMATICA'}, axis = 1, inplace=True)
    dados.rename({'PESO_CIENCIAS_HUMANAS': 'CIENCIAS_HUMANAS'}, axis = 1, inplace=True)
    dados.rename({'PESO_CIENCIAS_NATUREZA': 'CIENCIAS_NATUREZA'}, axis = 1, inplace=True)

    print(f'Dados de {curso} processados.')

    ### INSERE INFORMAÇÕES ADICIONAIS ###
    dados, dic_dados = insere_medias(dados, dic_dados, notas)
    print(f'Notas de {curso} processadas.')

    gera_relatorio(dados, dic_dados, ano, semestre, curso, regiao, notas)

    print('Pronto!')