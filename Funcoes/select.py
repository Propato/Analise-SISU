import sys

def end(erro):
    print(f'{erro} inválido.\nPrograma encerrado.')
    sys.exit()

def curso():
    curso = input('Insira curso a ser analisado: ').upper()
    if(curso == ""):
        end("Curso")
    return curso

def ano():
    ano = int(input('Ano: '))
    
    if 2019 <= ano <= 2023:
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

    if desejo[0] == 'n':
        return None
    
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

def instituicao():
    print('Deseja analisar apenas uma Instituição de Ensino? S/N')
    desejo = input().lower()

    if desejo[0] == 'n':
        return None

    opcao = input('IE: ').upper()
    return opcao

def notas():
    print('Insira as notas do TRI:')
    notas = [float(input('Redação: ')), float(input('Linguagens: ')), float(input('Matemática: ')), float(input('Humanas: ')), float(input('Natureza: '))]
    return notas

def all():
    curso_var = curso()
    ano_var = ano()
    semestre_var = semestre()
    regiao_var = regiao()
    instituicao_var = instituicao()
    notas_var = notas()

    return ano_var, semestre_var, curso_var, regiao_var, instituicao_var,notas_var

### FUNÇÔES AUXILIAR ###

def teste1():
    curso_var = 'MEDICINA'
    ano_var = 2021
    semestre_var = 2
    regiao_var = 'sudeste'
    instituicao_var = "UFES"
    notas_var = [980, 690, 840, 700, 720]

    return curso_var, ano_var, semestre_var, regiao_var, instituicao_var, notas_var

def teste2():
    curso_var = "ODONTOLOGIA"
    ano_var = 2023
    semestre_var = 1
    regiao_var = 'sudeste'
    instituicao_var = "UFES"
    # Redação, Linguagens, Matemática, Humanas, Natureza
    notas_var = [960, 667.2, 814.7, 711.2, 659.7]

    return curso_var, ano_var, semestre_var, regiao_var, instituicao_var, notas_var