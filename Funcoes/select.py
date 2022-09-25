import sys

def end(erro):
    print(f'{erro} inválido.\nPrograma encerrado.')
    sys.exit()

def curso():
    curso = input('Insira curso a ser analisado: ')
    return curso;

def ano():
    ano = int(input('Ano: '))
    
    if 2019 <= ano <= 2022:
        return ano
    end('Ano')

def semestre():
    semestre = int(input('Semestre (1 ou 2): '))
    
    if 1 == semestre or semestre == 2:
        return semestre
    end('Semestre')

def regiao():
    print('Deseja analisar destacando uma região? S/N')
    desejo = input().lower()

    if desejo == 'n' or desejo == 'na' or desejo == 'nã' or desejo == 'nao'  or desejo == 'não':
        return 'nenhuma'
    
    print('\nSelecione região desejada:')
    print('    1 - Sul')
    print('    2 - Sudeste')
    print('    3 - Centro-Oeste')
    print('    4 - Nordeste')
    print('    5 - Norte')
    opcao = input('Opção: ').lower()

    if opcao == '1' or opcao == 'sul':
        return 'sul'
    elif opcao == '2' or opcao == 'sudeste':
        return 'sudeste'
    elif opcao == '3' or opcao == 'centro-oeste':
        return 'centro-oeste'
    elif opcao == '4' or opcao == 'nordeste':
        return 'nordeste'
    elif opcao == '5' or opcao == 'norte':
        return 'norte'
    else:
        end('Região')

def notas():
    print('Deseja analisar calculando suas notas ponderadas? S/N')
    desejo = input().lower()

    if desejo == 'n':
        return [False]
    elif desejo == 's':
        print('Insira as notas do TRI:')
        notas = [float(input('Redação: ')), float(input('Linguagens: ')), float(input('Matemática: ')), float(input('Humanas: ')), float(input('Natureza: '))]
        return notas

def all():
    curso_var = curso()
    ano_var = ano()
    semestre_var = semestre()
    regiao_var = regiao()
    notas_var = notas()

    return ano_var, semestre_var, curso_var, regiao_var, igc_var, notas_var, corte_var

### FUNÇÃO AUXILIAR ###
def teste():
    curso_var = 'MEDICINA'
    ano_var = 2022
    semestre_var = 1
    regiao_var = 'sudeste'
    notas_var = [960, 708, 870, 700, 730]

    return curso_var, ano_var, semestre_var, regiao_var, notas_var