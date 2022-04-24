import sys

def end(erro):
    print(f'{erro} inválido.\nPrograma encerrado.')
    sys.exit()

def ano():
    ano = int(input('Ano: '))
    
    if 2019 <= ano <= 2022:
        return ano
    end('Ano')

def semestre():
    semestre = int(input('Semestre (1 ou 2): '))
    
    if 1 <= semestre <= 2:
        return semestre
    end('Semestre')

def regiao():
    print('Deseja analisar destacando uma região? S/N')
    desejo = input().lower()

    if desejo == 'n':
        return 'nenhuma'
    
    print('\nSelecione região desejada:')
    print('    1 - Sul')
    print('    2 - Sudeste')
    print('    3 - Centro-Oeste')
    print('    4 - Nordeste')
    print('    5 - Norte')
    opcao = input('Opção: ')

    if opcao == '1':
        return 'sul'
    elif opcao == '2':
        return 'sudeste'
    elif opcao == '3':
        return 'centro-oeste'
    elif opcao == '4':
        return 'nordeste'
    elif opcao == '5':
        return 'norte'
    else:
        end('Região')

def IGC():
    print('Deseja analisar e inserir dados do IGC? S/N')
    print('    IGC: Índice Geral de Cursos')
    desejo = input().lower()

    if desejo == 'n':
        return False
    elif desejo == 's':
        return True
    else:
        end('IGC')

def notas():
    print('Deseja analisar calculando suas notas ponderadas? S/N')
    desejo = input().lower()

    if desejo == 'n':
        return [False]
    elif desejo == 's':
        print('Insira as notas do TRI:')
        notas = [float(input('Redação: ')), float(input('Linguagens: ')), float(input('Matemática: ')), float(input('Humanas: ')), float(input('Natureza: '))]
        return notas

def corte():
    print('Deseja analisar junto das notas de corte? S/N')
    desejo = input().lower()

    if desejo == 'n':
        return [False]
    elif desejo == 's':
        print('Insira ano e semestre das notas de corte desejadas:')
        ano_var = ano()
        semestre_var = semestre()
        return [True, ano_var, semestre_var]

def all():
    ano_var = ano()
    semestre_var = semestre()
    curso_var = input('Curso: ').upper()
    regiao_var = regiao()
    igc_var = IGC()
    notas_var = notas()
    corte_var = corte()

    return ano_var, semestre_var, curso_var, regiao_var, igc_var, notas_var, corte_var

### FUNÇÃO AUXILIAR ###
def auto():
    ano_var = 2022
    semestre_var = 1
    curso_var = 'MEDICINA'
    regiao_var = 'sudeste'
    igc_var = True
    notas_var = [1000.0, 700.0, 815.0, 765.0, 735.0]
    corte_var = [True, 2022, 1]

    return ano_var, semestre_var, curso_var, regiao_var, igc_var, notas_var, corte_var