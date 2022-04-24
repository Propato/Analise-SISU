import pandas as pd
import sys

def media_ponderada(notas, pesos):
    soma_pesos = 0
    soma_notas = 0

    for i in range(len(pesos)):
        soma_pesos += pesos[i]
        soma_notas += notas[i]*pesos[i]
    return soma_notas/soma_pesos

def insere_medias(df, dic_df, notas, aprovacao):
    df['NOTAS'] = (media_ponderada(notas, (df['REDACAO'], df['LINGUAGENS'], df['MATEMATICA'], df['CIENCIAS_HUMANAS'], df['CIENCIAS_NATUREZA']))).to_frame()

    dic_df = pd.concat([dic_df, pd.DataFrame([['NOTAS', 'Médias ponderadas para cada curso']],columns=[dic_df.columns[0], dic_df.columns[1]])])
    
    if aprovacao:
        df['APROVAÇÃO'] = ((df['NU_NOTACORTE'] - df['NOTAS'])<0).to_frame()

        dic_df = pd.concat([dic_df, pd.DataFrame([['APROVAÇÃO', 'Indica se as suas notas são maiores (escrito VERDADEIRO) ou menores (escrito FALSO) que a nota de corte']],columns=[dic_df.columns[0], dic_df.columns[1]])])
    
    return df, dic_df

def insere_top(df2, dic_df2):

    df1 = pd.read_excel('dados/IGC_2019.xlsx', usecols='B, O, P')

    df1.sort_values(by=' Código da IES', kind='stable', inplace=True, ignore_index=True)
    df2.sort_values(by='CO_IES', kind='stable', inplace=True, ignore_index=True)

    i=j=0
    while i < df1.shape[0] and j < df2.shape[0]:

        if j>0 and df1.iloc[j][' Código da IES'] < df1.iloc[j-1][' Código da IES']:
            print(f'Fim da filtragem do IGC.')
            break
    
        if df1.iloc[j][' Código da IES'] == df2.iloc[i]['CO_IES']:
            j+=1
            i+=1
            continue
        elif df1.iloc[j][' Código da IES'] < df2.iloc[i]['CO_IES']:
            df1=df1.drop(index=j).reset_index(drop=True)
            continue
        elif df1.iloc[j][' Código da IES'] > df2.iloc[i]['CO_IES']:
            if j>0:
                k = j-1
                while k>0 and df1.iloc[k][' Código da IES'] > df2.iloc[i]['CO_IES']:
                    k-=1

                if df1.iloc[k][' Código da IES'] == df2.iloc[i]['CO_IES']:
                    df1 = pd.concat([df1, df1.iloc[[k]]], ignore_index=True)
                    i+=1
                    continue
                else:
                    df1 = pd.concat([df1, pd.DataFrame([[df2.iloc[i]['CO_IES'], 0, 0]],columns=[' Código da IES', ' IGC (Contínuo)', ' IGC (Faixa)'])], ignore_index=True)
                    i+=1
                    continue

            else:
                df1 = pd.concat([df1, pd.DataFrame([[df2.iloc[i]['CO_IES'], 0, 0]],columns=[' Código da IES', ' IGC (Contínuo)', ' IGC (Faixa)'])], ignore_index=True)
                i+=1
                continue

    while i < df2.shape[0]:
        if df1.iloc[j-1][' Código da IES'] == df2.iloc[i]['CO_IES']:
            df1 = pd.concat([df1, df1.iloc[[j-1]]], ignore_index=True)
            i+=1
        else:
            df1 = pd.concat([df1, pd.DataFrame([[df2.iloc[i]['CO_IES'], 0, 0]],columns=[' Código da IES', ' IGC (Contínuo)', ' IGC (Faixa)'])], ignore_index=True)
            i+=1

    df1.reset_index(drop=True, inplace=True)
    df1.sort_values(by=' Código da IES', kind='stable', inplace=True, ignore_index=True)

    if (df1[' Código da IES'] == df2['CO_IES']).all() == False:
        print(f'Erro na filtragem do IGC, favor checar e tentar novamente depois de corrigido o problema.')

    df2['IGC (Contínuo)'] = df1[' IGC (Contínuo)']
    df2['IGC (Faixa)'] = df1[' IGC (Faixa)']

    dic_df2 = pd.concat([dic_df2, pd.DataFrame([['IGC (Contínuo)', 'Nota da avaliação do IGC (Contínuo) de acordo com o MEC. Cursos com 0 significam apenas que não havia informações sobre eles no banco de dados do ICG 2019 do MEC, NÃO significam que possuem nota 0'], ['IGC (Faixa)', 'Nota da avaliação do IGC (Faixa) de acordo com o MEC. Cursos com 0 significam apenas que não havia informações sobre eles no banco de dados do ICG 2019 do MEC, NÃO significam que possuem nota 0']],columns=[dic_df2.columns[0], dic_df2.columns[1]])])
    
    print('Dados do IGC processados.')

    return df2, dic_df2