from Funcoes.analise import dados_sisu
import Funcoes.select as select

print('_'*30, 'ANALISE DE DADOS DO SISU EM FUNÇÃO DOS CURSOS/INSTITUIÇÕES', '_'*30, end='\n\n')

###################################  ###################################
print('Insira as informações pedidas para realizar a analise de dados a respeito dos relatórios disponibilidados pelo SISU.\n')

print('Informe o curso, ano, semestre e regiao que deseja analisar, inserindo também suas notas.')

curso_var, ano_var, semestre_var, regiao_var, instituicao_var, notas_var = select.teste2()
dados_sisu(curso_var, ano_var, semestre_var, regiao_var, instituicao_var, notas_var)