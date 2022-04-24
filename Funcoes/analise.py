import pandas as pd
import openpyxl
import sys
import Funcoes.infos as infos
import Funcoes.visual as visual

def limpa_dic(dic, df):
    i=0
    while i < dic.shape[0]:
        if dic.iloc[i]['Nome da coluna'] not in df.columns:
            dic = dic.drop(index=i)
            dic.reset_index(drop=True, inplace=True)
            i-=1
        i+=1
    return dic

def abre_checa_excel(ano, semestre, corte):
    dic_dados = pd.read_excel(f'Dados/Vagas ofertadas/Portal Sisu_Sisu {ano}-{semestre}_Vagas ofertadas.xlsx', sheet_name=0)
    dados = pd.read_excel(f'Dados/Vagas ofertadas/Portal Sisu_Sisu {ano}-{semestre}_Vagas ofertadas.xlsx', sheet_name=1, usecols='C:E, H:P, R, S, V, W, Y, AA, AC, AE')    
    dados.sort_values(by='CO_IES_CURSO', kind='stable', inplace=True, ignore_index=True)

    if corte[0]:
        dic_cortes = pd.read_excel(f'Dados/Inscrições e notas de corte/Portal Sisu_Sisu {corte[1]}-{corte[2]}_Inscrições e notas de corte.xlsx', sheet_name=0)
        cortes = pd.read_excel(f'Dados/Inscrições e notas de corte/Portal Sisu_Sisu {corte[1]}-{corte[2]}_Inscrições e notas de corte.xlsx', sheet_name=1, usecols='L, Q, S:U')
        cortes.sort_values(by='CO_IES_CURSO', kind='stable', inplace=True, ignore_index=True)
        
        ### CHECAGEM ###
        if not (dados['CO_IES_CURSO'] == cortes['CO_IES_CURSO']).all():
            print('Dados dos cursos e cortes não batem, fazer checar as planilhas e tentar novamente.')
            sys.exit()

        pesos = pd.read_excel(f'Dados/Vagas ofertadas/Portal Sisu_Sisu {corte[1]}-{corte[2]}_Vagas ofertadas.xlsx', sheet_name=1, usecols='L, W, Y, AA, AC, AE')
        pesos.sort_values(by='CO_IES_CURSO', kind='stable', inplace=True, ignore_index=True)
        
        ### CHECAGEM ###
        if (dados['CO_IES_CURSO'] == pesos['CO_IES_CURSO']).all():
            if not (dados['PESO_REDACAO'] == pesos['PESO_REDACAO']).all() and (dados['PESO_LINGUAGENS'] == pesos['PESO_LINGUAGENS']).all() and (dados['PESO_MATEMATICA'] == pesos['PESO_MATEMATICA']).all() and (dados['PESO_CIENCIAS_HUMANAS'] == pesos['PESO_CIENCIAS_HUMANAS']).all() and (dados['PESO_CIENCIAS_NATUREZA'] == pesos['PESO_CIENCIAS_NATUREZA']).all():
                print(f'Pesos de {ano}/{semestre} e {corte[1]}/{corte[2]} diferem, programa encerrado.')
                sys.exit()

        else:
            print(f'Cursos de {ano}/{semestre} e {corte[1]}/{corte[2]} diferem, programa encerrado.')
            sys.exit()
    
        return dic_dados, dados, dic_cortes, cortes

    return dic_dados, dados, False, False

def dados_sisu(curso, ano, semestre, regiao, notas, igc, corte):
    print('='*127)
    print('=Executando código para processar dados do SISU de 2022 a respeito do curso desejado para as cotas de ampla e escolas públicas=')
    print('='*127)

    dic_dados, dados, dic_cortes, cortes = abre_checa_excel(ano, semestre, corte)

    print('Tabelas lida.')

    cotas_invalidas = ['autodeclarados', 'pretos', 'negros', 'pardos', 'índios', 'indígenas', 'indígena,', 'quilombolas', 'transexuais', 'vulnerabilidade', 'deficiência', 'necessidades', 'carência', 'inferior', 'regional', 'região', 'microrregiões', 'mesorregiões', 'membros de comunidade', 'residentes', 'residem', 'no estado de pernambuco', 'localizadas', 'baixa', 'natal', 'de até um salário-mínimo']
    
    ### INSERE DADOS DE INSCRIÇÕES NO DF PRINCIPAL ###
    if corte[0]:
        if (dados['DS_MOD_CONCORRENCIA'] == cortes['DS_MOD_CONCORRENCIA']).all():
            dados.insert(14, f'QT_VAGAS_{ano}/{semestre}',  cortes['QT_VAGAS_CONCORRENCIA'])
            dados.insert(15, 'QT_INSCRICAO',  cortes['QT_INSCRICAO'])
            dados.insert(22, 'NU_NOTACORTE',  cortes['NU_NOTACORTE'])
        else:
            print('aaaaaaaaaaaa')
            sys.exit()

    dados = dados.loc[dados['NO_CURSO']==curso].reset_index(drop=True)
    
    ### CHECAGEM ###
    if dados.shape[0] == 0:
        print(f'Curso {curso.upper()} inválido.')
        sys.exit()

    ### FILTRA POR COTAS ###
    i=0
    while i < dados.shape[0]:
        for j in cotas_invalidas:
            if(j in dados.iloc[i]['DS_MOD_CONCORRENCIA'].lower()):
                dados = dados.drop(index=i)
                dados.reset_index(drop=True, inplace=True)
                i-=1
                break
        i+=1

    ### REORGANIZA DIC_DADOS ###
    dic_dados = limpa_dic(dic_dados, dados)

    ### REORGANIZA DIC_CORTES ###
    if corte[0]:
        dic_cortes = limpa_dic(dic_cortes, cortes)

        for i in range(3):
            dic_cortes = dic_cortes.drop(index=0)
            dic_cortes.reset_index(drop=True, inplace=True)

        ### FINALIZA DIC_DADOS ###
        dic_dados1 = dic_dados[:14]
        dic_dados2 = dic_dados[14:]

        dic_dados1 = pd.concat([dic_dados1, pd.DataFrame({f'{dic_dados.columns[0]}': [f'QT_VAGAS_{ano}/{semestre}'], f'{dic_dados.columns[1]}': [f'Quantidade de vagas ofertadas naquela modalidade em {ano}/{semestre}']})])

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
    if float == type(notas[0]):
        dados, dic_dados = infos.insere_medias(dados, dic_dados, notas, corte[0])
        print(f'Notas de {curso} processadas.')
    if igc:
        dados, dic_dados = infos.insere_top(dados, dic_dados)

    writer = pd.ExcelWriter(f'Relatorios/{curso}_{ano}_{semestre}.xlsx', engine='openpyxl')

    dic_dados.to_excel(writer, index=False, sheet_name='Dicionário de dados')
    dados.to_excel(writer, index=False, sheet_name=curso)

    writer.save()

    wb = openpyxl.load_workbook(f"Relatorios/{curso}_{ano}_{semestre}.xlsx")
    ws0 = wb['Dicionário de dados']
    
    ### SIMPLIFICA NOME PARA OS PESOS NO DIC ###
    areas = ['REDACAO', 'LINGUAGENS', 'MATEMATICA', 'CIENCIAS_HUMANAS', 'CIENCIAS_NATUREZA']
    j=0
    for i in range(1, dic_dados.shape[0]+2):
        if areas[j] in ws0[f'A{i}'].value:
            ws0[f'A{i}'].value = areas[j]
            j+=1
        if j>=5:
            break

    ### ATUALIZA O VISUAL DA TABELA NO EXCEL ###
    visual.cores(wb, dados, curso, regiao, corte[0], float == type(notas[0]))
    visual.layout(wb, dados, dic_dados, curso, corte, float == type(notas[0]), igc)

    wb.save(f'Relatorios/{curso}_{ano}_{semestre}.xlsx')

    print('Pronto!')