import pandas as pd
import openpyxl
import sys
import Funcoes.visual as visual

def abre_excel(ano, semestre):
    try:
        dic_dados = pd.read_excel(f'Dados/Vagas ofertadas/Portal Sisu_Sisu {ano}-{semestre}_Vagas ofertadas.xlsx', sheet_name=0)
        dados = pd.read_excel(f'Dados/Vagas ofertadas/Portal Sisu_Sisu {ano}-{semestre}_Vagas ofertadas.xlsx', sheet_name=1, usecols='C:E, H:P, R, S, V, W, Y, AA, AC, AE')    
        
        dic_cortes = pd.read_excel(f'Dados/Inscrições e notas de corte/Portal Sisu_Sisu {ano}-{semestre}_Inscrições e notas de corte.xlsx', sheet_name=0)
        cortes = pd.read_excel(f'Dados/Inscrições e notas de corte/Portal Sisu_Sisu {ano}-{semestre}_Inscrições e notas de corte.xlsx', sheet_name=1, usecols='E, L, M, Q, T, U')
    
        return dic_dados, dados, dic_cortes, cortes
    
    except Exception as e:
        print(e)
        print('Erro ao abrir arquivo.')
        sys.exit()

def gera_relatorio(df, dic_df, ano, semestre, cursoInstituicao, regiao, notas):
    print("Curso | Instituicao:", cursoInstituicao + ".")
    writer = pd.ExcelWriter(f'Relatorios/{cursoInstituicao}_{ano}_{semestre}.xlsx', engine='openpyxl')

    dic_df.to_excel(writer, index=False, sheet_name='Dicionário de dados')
    df.to_excel(writer, index=False, sheet_name=cursoInstituicao)

    writer.save()

    wb = openpyxl.load_workbook(f"Relatorios/{cursoInstituicao}_{ano}_{semestre}.xlsx")
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
    visual.cores(wb, df, cursoInstituicao, regiao)
    visual.layout(wb, df, dic_df, cursoInstituicao, notas)

    wb.save(f'Relatorios/{cursoInstituicao}_{ano}_{semestre}.xlsx')
    print(f'Arquivo: Relatorios/{cursoInstituicao}_{ano}_{semestre}.xlsx')

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

def filtra_excel(dados, cortes, instituicao, curso):
    
    cotas_invalidas = ['autodeclarados', 'pretos', 'negros', 'pardos', 'índios', 'indígenas', 'indígena,', 'quilombolas', 'ciganos',  'transexuais', 'vulnerabilidade', 'deficiência', 'necessidades', 'carência', 'inferior', 'regional', 'região', 'microrregiões', 'mesorregiões', 'membros de comunidade', 'residentes', 'residem', 'residam', 'no estado de pernambuco', 'localizadas', 'baixa', 'natal', 'de até um salário-mínimo']

    if(curso):
        dados = dados.loc[dados['NO_CURSO']==curso].reset_index(drop=True)
        cortes = cortes.loc[curso==cortes['NO_CURSO']].reset_index(drop=True)
    if(instituicao):
        dados = dados.loc[dados['SG_IES']==instituicao].reset_index(drop=True)
        cortes = cortes.loc[cortes['SG_IES']==instituicao].reset_index(drop=True)
    
    ### CHECAGEM - CURSO ###
    if (dados.shape[0] == 0) or (cortes.shape[0] == 0):
        print(f'Curso ou Instituição inválidos, sem ofertas de vagas ou pesos.')
        sys.exit()

    ### FILTRA POR COTAS ###
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

def dados_sisu(curso, ano, semestre, regiao, instituicao, notas):
    print('='*127)
    print('=Executando código para processar dados do SISU de 2022 para as cotas de ampla e escolas públicas=')
    print('='*127)

    dic_dados, dados, dic_cortes, cortes = abre_excel(ano, semestre)

    print('Tabelas lida.')

    dados, cortes = filtra_excel(dados, cortes, instituicao, curso)

    ### REORGANIZA DIC_DADOS ###
    dic_dados = limpa_dic(dic_dados, dados)
    ### REORGANIZA DIC_CORTES ###
    dic_cortes = limpa_dic(dic_cortes, cortes)

    for i in range(4):
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

    if(curso and instituicao):
        print(f'Dados de {curso}/{instituicao} processados.')
    elif(curso):
        print(f'Dados de {curso} processados.')
    else:
        print(f'Dados de {instituicao} processados.')

    ### INSERE INFORMAÇÕES ADICIONAIS ###
    dados, dic_dados = insere_medias(dados, dic_dados, notas)
    if(curso and instituicao):
        print(f'Notas de {curso}/{instituicao} processadas.')
    elif(curso):
        print(f'Notas de {curso} processadas.')
    else:
        print(f'Notas de {instituicao} processadas.')
    
    if(curso): info = curso
    elif(instituicao): info = instituicao
    else: info = "TUDO"
    gera_relatorio(dados, dic_dados, ano, semestre, info, regiao, notas)

    print('Pronto!')