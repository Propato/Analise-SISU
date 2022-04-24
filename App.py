import sys
from Funcoes.analise import dados_sisu
import Funcoes.select as select

print('_'*30, 'ANALISE DE DADOS DO SISU EM FUNÇÃO DOS CURSOS', '_'*30, end='\n\n')
print('='*49, 'INÍCIO', '='*50, end='\n\n')

###################################  ###################################
print('Insira as informações pedidas para realizar a analise de dados a respeito dos relatórios disponibilidados pelo SISU.\n')

print('Informe o ano e semestre que deseja analisar.')
ano, semestre, curso, regiao, igc, notas, corte = select.auto()
dados_sisu(curso, ano, semestre, regiao, notas, igc, corte)