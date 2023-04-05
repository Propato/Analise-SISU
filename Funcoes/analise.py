import pandas as pd
import openpyxl
import sys
from Funcoes.visual import cores, layout2

def abre_excel(ano, semestre):
    try:
        dic_dados = pd.read_excel(f'data/Vagas ofertadas/Portal Sisu_Sisu {ano}-{semestre}_Vagas ofertadas.xlsx', sheet_name=0)
        dados = pd.read_excel(f'data/Vagas ofertadas/Portal Sisu_Sisu {ano}-{semestre}_Vagas ofertadas.xlsx', sheet_name=1, usecols='C:E, H:P, R, S, V, W, Y, AA, AC, AE')    
        
        dic_cortes = pd.read_excel(f'data/Inscrições e notas de corte/Portal Sisu_Sisu {ano}-{semestre}_Inscrições e notas de corte.xlsx', sheet_name=0)
        cortes = pd.read_excel(f'data/Inscrições e notas de corte/Portal Sisu_Sisu {ano}-{semestre}_Inscrições e notas de corte.xlsx', sheet_name=1, usecols='E, L, M, Q, T, U')
    
        return dic_dados, dados, dic_cortes, cortes
    
    except Exception as e:
        print(e)
        print('Erro ao abrir arquivo.')
        sys.exit()

def gera_relatorio(df, dic_df, ano, semestre, cursoInstituicao, regiao, notas):
    print("Curso | Instituicao:", cursoInstituicao + ".")
    
    path = f'Relatorios/{cursoInstituicao}_{ano}_{semestre}.xlsx'

    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df = cores(df, regiao)
        # dic_df.to_excel(writer, index=False, sheet_name='Dicionário de dados')
        df.to_excel(writer, index=False, sheet_name=cursoInstituicao)
        
        workbook = writer.book

            # Get the worksheet object from the workbook
        worksheet = workbook[cursoInstituicao]
        layout2(df, worksheet, notas)

        # wb = openpyxl.load_workbook(f"Relatorios/{cursoInstituicao}_{ano}_{semestre}.xlsx")

        ### ATUALIZA O VISUAL DA TABELA NO EXCEL ###
        # df_colors.to_excel(writer, index=False, sheet_name=cursoInstituicao)
    
    # ### ATUALIZA O VISUAL DA TABELA NO EXCEL ###
    # cores(wb, df, cursoInstituicao, regiao)

    # layout(wb, df, dic_df, cursoInstituicao, notas)
        # df_visual = layout2(df_colors)
        # df_visual.to_excel(writer, index=False, sheet_name=cursoInstituicao)

    # wb.save(path)
    print('Arquivo:', path)

def media_ponderada(notas, pesos):
    soma_pesos = 0
    soma_notas = 0

    for i in range(len(pesos)):
        soma_pesos += pesos[i]
        soma_notas += notas[i]*pesos[i]
    return round(soma_notas/soma_pesos, 2)

def insere_medias(df, dic_df, notas):
    df['NOTAS'] = media_ponderada(notas, [df['REDACAO'], df['LINGUAGENS'], df['MATEMATICA'], df['CIENCIAS_HUMANAS'], df['CIENCIAS_NATUREZA']])
    df['APROVAÇÃO'] = ((df['NU_NOTACORTE'] - df['NOTAS'])<0)
    
    dic_df = pd.concat([dic_df, pd.DataFrame([['NOTAS', 'Médias ponderadas para cada curso'], ['APROVAÇÃO', 'Indica se as suas notas são maiores (escrito VERDADEIRO) ou menores (escrito FALSO) que a nota de corte']],columns=[dic_df.columns[0], dic_df.columns[1]])])
    # dic_df = pd.concat([dic_df, pd.DataFrame([],columns=[dic_df.columns[0], dic_df.columns[1]])])
    
    return df, dic_df

def limpa_dic(dic, df):
    conjunto = set()
    for index in dic.index:
        if dic.loc[index, 'Nome da coluna'] not in df.columns:
            conjunto.add(index)
    dic = dic.drop(labels=list(conjunto))
    dic.reset_index(drop=True, inplace=True)

    return dic

def filtra_excel(dados, cortes, instituicao, curso):
    
    cotas_invalidas = ['autodeclarados', 'índios', 'indígena', 'quilombolas', 'ciganos',  'transexuais', 'vulnerabilidade', 'deficiência', 'necessidades', 'carência', 'inferior', 'regional', 'região', 'regiões', 'membros de comunidade', 'residentes', 'residem', 'residam', 'no estado de pernambuco', 'localizadas', 'baixa', 'natal', 'de até um salário-mínimo']

    if(curso and instituicao):
        dados = dados.loc[(dados['NO_CURSO']==curso) & (dados['SG_IES']==instituicao)].reset_index(drop=True)
        cortes = cortes.loc[(cortes['NO_CURSO']==curso) & (cortes['SG_IES']==instituicao)].reset_index(drop=True)
    elif(curso):
        dados = dados.loc[dados['NO_CURSO']==curso].reset_index(drop=True)
        cortes = cortes.loc[cortes['NO_CURSO']==curso].reset_index(drop=True)
    elif(instituicao):
        dados = dados.loc[dados['SG_IES']==instituicao].reset_index(drop=True)
        cortes = cortes.loc[cortes['SG_IES']==instituicao].reset_index(drop=True)
        
    ### CHECAGEM - CURSO ###
    if (dados.shape[0] == 0) or (cortes.shape[0] == 0):
        print(f'Curso ou Instituição inválidos, sem ofertas de vagas ou pesos.')
        sys.exit()

    ### FILTRA POR COTAS ###
    # cortes = cortes[~cortes['DS_MOD_CONCORRENCIA'].str.lower().isin(cotas_invalidas)]
    conjunto = set()
    for elem in cotas_invalidas:
        for index in cortes.index:
            if elem in cortes.loc[index, 'DS_MOD_CONCORRENCIA'].lower():
                conjunto.add(index)
    cortes = cortes.drop(labels=list(conjunto))
    cortes.reset_index(drop=True, inplace=True)
    
    conjunto = set()
    for elem in cotas_invalidas:
        for index in dados.index:
            if elem in dados.loc[index, 'DS_MOD_CONCORRENCIA'].lower():
                conjunto.add(index)
    dados = dados.drop(labels=list(conjunto))
    dados.reset_index(drop=True, inplace=True)
    
    # merge the two dataframes based on the 'id' column, and only include the 'name' and 'age' columns in the result
    merged_df = pd.merge(dados, cortes[['QT_INSCRICAO', 'NU_NOTACORTE', 'CO_IES_CURSO', 'DS_MOD_CONCORRENCIA']], on=['CO_IES_CURSO', 'DS_MOD_CONCORRENCIA'])

    merged_df = merged_df[['CO_IES', 'NO_IES', 'SG_IES', 'NO_CAMPUS', 'NO_MUNICIPIO_CAMPUS', 'SG_UF_CAMPUS', 'DS_REGIAO', 'CO_IES_CURSO', 'NO_CURSO', 'DS_GRAU', 'DS_TURNO', 'DS_PERIODICIDADE', 'NU_VAGAS_AUTORIZADAS', 'QT_VAGAS_CONCORRENCIA', 'QT_INSCRICAO', 'DS_MOD_CONCORRENCIA', 'PESO_REDACAO', 'PESO_LINGUAGENS', 'PESO_MATEMATICA', 'PESO_CIENCIAS_HUMANAS', 'PESO_CIENCIAS_NATUREZA', 'NU_NOTACORTE']]

    return merged_df, cortes

def dados_sisu(curso, ano, semestre, regiao, instituicao, notas):
    print('='*127)
    print('=Executando código para processar dados do SISU de 2022 para as cotas de ampla e escolas públicas=')
    print('='*127)

    dic_dados, dados, dic_cortes, cortes = abre_excel(ano, semestre)

    print('Tabelas lida.')

    dados, cortes = filtra_excel(dados, cortes, instituicao, curso)

    ### REORGANIZA DICs ###
    dic_dados = limpa_dic(dic_dados, dados)
    dic_cortes = limpa_dic(dic_cortes, cortes)

    dic_dados = pd.concat([dic_dados[:14], dic_cortes[5:], dic_dados[14:], dic_cortes[4:5]]).reset_index(drop=True)
    
    ### SIMPLIFICA NOME PARA COLUNA DOS PESOS ###
    rename_dict = {
    'PESO_REDACAO': 'REDACAO',
    'PESO_LINGUAGENS': 'LINGUAGENS',
    'PESO_MATEMATICA': 'MATEMATICA',
    'PESO_CIENCIAS_HUMANAS': 'CIENCIAS_HUMANAS',
    'PESO_CIENCIAS_NATUREZA': 'CIENCIAS_NATUREZA'
    }

    dados.rename(columns=rename_dict, inplace=True)
    for i, (old_name, new_name) in enumerate(rename_dict.items()):
        dic_dados.loc[i + 15, "Nome da coluna"] = new_name

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