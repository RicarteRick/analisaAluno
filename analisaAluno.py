import xlrd
import datetime

def orgDatas(arq):
    plano = arq.sheet_by_name('Datas')
    datas = []
    tdatas = []
    for i in range(1, 7):  # linhas
        for j in range(1, 5):  # colunas
            dataxlrd = plano.cell(i, j).value
            datapy = xlrd.xldate.xldate_as_datetime(dataxlrd, arq.datemode)  # convertendo para data do python
            data = datapy.strftime("%d/%m/%Y")  # formatando a data pro nosso padrão e com barras, pra ficar mais bonito no arquivo
            datas.append(data)
        tdatas.append(datas[:])  # copiando a lista anterior dentro dessa lista, que contém as datas de todas as matérias separadas por listas
        datas.clear()

    return tdatas


# segue o mesmo raciocínio das datas
def orgValores(arq):
    plano = arq.sheet_by_name('Valores')
    valores = []
    tvalores = []
    for i in range(1, 7):
        for j in range(1, 4):  # não tem motivo pra pegar a última coluna aqui
            campo = plano.cell(i, j).value
            valores.append(campo)
        tvalores.append(valores[:])
        valores.clear()

    return tvalores


# segue o mesmo raciocínio das datas
def orgNotas(arq):
    plano = arq.sheet_by_name('Notas')
    notas = []
    tnotas = []
    for i in range(1, 7):
        for j in range(1, 5):
            campo = plano.cell(i, j).value
            if campo == '':  # se for vazio
                campo = 0.0
            notas.append(campo)
        tnotas.append(notas[:])
        notas.clear()

    return tnotas


def calcNotaRestante(numMat, valores, notas):
    notaAtual = notas[numMat][0] + notas[numMat][1] + notas[numMat][2]
    valorTotal = valores[numMat][0] + valores[numMat][1] + valores[numMat][2]
    restante = (valorTotal * 0.6) - notaAtual   # porque precisa de 60% pra passar

    return restante


def calcDiasRestantes(numMat, datas):
    dataBarnabe = datetime.datetime.strptime('14/06/2019', '%d/%m/%Y')
    dataP3 = datetime.datetime.strptime(datas[numMat][2], '%d/%m/%Y')
    restante = abs((dataP3 - dataBarnabe).days)

    return restante


def calcDiasSub(numMat, datas):
    dataBarnabe = datetime.datetime.strptime('14/06/2019', '%d/%m/%Y')
    dataSub = datetime.datetime.strptime(datas[numMat][3], '%d/%m/%Y')
    restante = abs((dataSub - dataBarnabe).days)

    return restante


def calcPontosSub(numMat, restante, notas):
    pontos = 0
    menor = 0
    if numMat == 2 or numMat == 4:
        pontos = restante - 100
    elif numMat == 0 or numMat == 5:
        retornaMenor(numMat, restante, notas)
        pontos = restante + menor   # como a nota menor será substituída, entende-se que o que falta para passar seja o que faltava antes + a nota que Barnabé tirou

    return pontos


# fiz essa função pois o valor também será usado mais pra frente, não só na função acima ^
def retornaMenor(numMat, restante, notas):
    if notas[numMat][0] <= notas[numMat][1]:
        menor = notas[numMat][0]
    else:
        menor = notas[numMat][1]

    return menor


def orgArquivo(arq):
    dataAtual = datetime.datetime.strptime('14/06/2019', '%d/%m/%Y')
    dataCabecalho = dataAtual.strftime("%d/%m/%Y")
    datas = orgDatas(arq)
    valores = orgValores(arq)
    notas = orgNotas(arq)

    fp = open('Status_Barnabe.txt', 'w')    # abrindo arquivo em modo de criação/escrita

    fp.write(f'-=-=-= NOME: Barnabé =-=-=-= UNIVERSIDADE FEDERAL DOS DEV-BACKEND =-=-=-= DATA: {dataCabecalho} =-=-=-\n\n')

    for numMat in range(0, 6):
        restante = calcNotaRestante(numMat, valores, notas)
        dias = calcDiasRestantes(numMat, datas)

        fp.write(f'Matéria {numMat + 1}:\n')
        fp.write(f'Valores necessários: P1: {valores[numMat][0]}, P2: {valores[numMat][1]}, P3: {valores[numMat][2]}')
        if numMat == 2 or numMat == 4:
            fp.write(' (média das 3 notas)')

        fp.write(f'\nNotas de Barnabé: P1: {notas[numMat][0]}, P2: {notas[numMat][1]}, P3: {notas[numMat][2]}\n')

        fp.write('Estilo da prova substitutiva: ')
        if numMat == 0 or numMat == 5:
            fp.write('Substitui a prova cuja menor nota foi tirada por outra de mesmo valor, refazendo a média depois.\n')
        elif numMat == 2 or numMat == 4:
            fp.write('Prova de valor 100, refazendo a média juntando-a com a média das demais provas.\n')
        else:
            fp.write('Não possui prova substitutiva.\n')

        if restante <= 0:
            fp.write('Já está aprovado!!')
        elif (numMat == 1 and restante > 30) or (numMat == 3 and restante > 40):
            fp.write('Já está reprovado.')
        else:
            fp.write(f'Ainda faltam {restante} pontos para passar nessa matéria.')
            if numMat == 2 or numMat == 4:
                fp.write(' (para atingir os 180)')  # para não precisar fazer outras condições e cálculos para converter essas notas

        fp.write(f'\nFaltam {dias} dias para a próxima prova(P3) dessa matéria, que ocorrerá no dia {datas[numMat][2]}\n')
        if (numMat == 0 and restante > 40) or (numMat == 5 and restante > 30) or (numMat == 2 and restante > 100) or (numMat == 4 and restante > 100):
            sub = calcDiasSub(numMat, datas)
            pontos = calcPontosSub(numMat, restante, notas)
            if numMat == 2 or numMat == 4:
                if sub < 100:
                    fp.write(f'Faltam {sub} dias para a prova substitutiva. Você deverá tirar no mínimo {pontos} pontos nela, levando em conta que tire 100% na P3.\n')
                else:
                    fp.write('Mesmo com a sub, estará reprovado.\n')
            else:
                menor = retornaMenor(numMat, restante, notas)
                if sub < menor:
                    fp.write(f'Faltam {sub} dias para a prova substitutiva. Você deverá tirar no mínimo {pontos} pontos nela, substituindo a prova cuja nota fora a menor.\n')
                else:
                    fp.write('Mesmo com a sub, estará reprovado.\n')
        fp.write('\n')

    fp.close()


arq = xlrd.open_workbook('provas.xls')  # abrindo o arquivo com os dados

orgArquivo(arq)
