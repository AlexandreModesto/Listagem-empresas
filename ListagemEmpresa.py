import pandas as pd
import openpyxl
import PySimpleGUI as sg
file = open("Importação Sacados.txt","w")
text = sg.popup_get_file('Coloque a planilha Importação Sacados',title='Importação Sacados')
planilha = pd.read_excel(text,engine='openpyxl')
df = pd.DataFrame(planilha)
indice=1
for indice,row in df.iterrows():
    concatenando =df['NÚMERO CPF / CNPJ'].values[indice]
    concatenandoConv = str(concatenando)
    concatenandoConv= concatenandoConv.upper()
    tamanho= len(concatenandoConv)
    isCPF = str(0)
    isCNPJ=str(1)
    if tamanho == 11:
        concatenadoCPF=isCPF +'000'+concatenandoConv

    elif tamanho == 10:
        concatenadoCPF =isCPF + '0000' + concatenandoConv

    elif tamanho == 9:
        concatenadoCPF = isCPF+'00000' + concatenandoConv

    elif tamanho == 14:
        concatenadoCPF = isCNPJ+''+ concatenandoConv

    elif tamanho == 13:
        concatenadoCPF = isCNPJ+'0'+ concatenandoConv

    elif tamanho == 12:
        concatenadoCPF= isCNPJ+ '00'+concatenandoConv

    # *************************************************************************
    concatenando = df['NOME CLIENTE / RAZÃO SOCIAL'].values[indice]
    concatenandoConv = str(concatenando)
    concatenandoConv= concatenandoConv.upper()
    concatenandoConv = concatenandoConv.replace('Ã', 'A')
    concatenandoConv = concatenandoConv.replace('Á', 'A')
    concatenandoConv = concatenandoConv.replace('Ó', 'O')
    concatenandoConv = concatenandoConv.replace('Õ', 'O')
    concatenandoConv = concatenandoConv.replace('É', 'E')
    concatenandoConv = concatenandoConv.replace('Ç', 'C')
    concatenandoConv = concatenandoConv.replace('Í', 'I')
    concatenandoConv = concatenandoConv.replace('Â', 'A')
    concatenandoConv = concatenandoConv.replace('Ê', 'E')
    concatenandoConv = concatenandoConv.replace('Ú', 'U')
    tamanho = len(concatenandoConv)
    if tamanho <=50:
        newTamanho = 50 - tamanho
        rEspacos  = (' ') * newTamanho
        finalNome = concatenandoConv + rEspacos
    else:
        finalNome = concatenandoConv[:50]
    # *************************************************************************
    concatenando = df['ENDEREÇO'].values[indice]
    concatenandoConv = str(concatenando)
    concaend= concatenandoConv.upper()
    concaend=concaend.replace('Ã','A')
    concaend=concaend.replace('Á','A')
    concaend = concaend.replace('Ó','O')
    concaend = concaend.replace('Õ','O')
    concaend = concaend.replace('É','E')
    concaend = concaend.replace('Ç', 'C')
    concaend = concaend.replace('Í','I')
    concaend = concaend.replace('Â','A')
    concaend = concaend.replace('Ê','E')
    concaend = concaend.replace('Ú', 'U')

    tamanho = len(concaend)
    finalENde=' '
    if tamanho <=40:
        newTamanho = 40 - tamanho
        eEspacos  = (' ') * newTamanho
        finalENde = concaend + eEspacos
    else:
        finalENde = concaend[:40]
    # ****************************************************************************
    concatenando = df['BAIRRO'].values[indice]
    concatenandoConv = str(concatenando)
    concatenandoConv= concatenandoConv.upper()
    concatenandoConv = concatenandoConv.replace('Ã', 'A')
    concatenandoConv = concatenandoConv.replace('Á', 'A')
    concatenandoConv = concatenandoConv.replace('Ó', 'O')
    concatenandoConv = concatenandoConv.replace('Õ', 'O')
    concatenandoConv = concatenandoConv.replace('É', 'E')
    concatenandoConv = concatenandoConv.replace('Ç', 'C')
    concatenandoConv = concatenandoConv.replace('Í', 'I')
    concatenandoConv = concatenandoConv.replace('Â', 'A')
    concatenandoConv = concatenandoConv.replace('Ê', 'E')
    concatenandoConv = concatenandoConv.replace('Ú', 'U')
    tamanho = len(concatenandoConv)
    concatatenadoBAIRRO = ''
    if tamanho <=15:
        newTamanho = 15 - tamanho
        espacos = espacos = (' ') * newTamanho
        concatatenadoBAIRRO = concatenandoConv + espacos
    else:
        concatatenadoBAIRRO = concatenandoConv[:15]
    # ***************************************************************************
    concatenando = df['CIDADE'].values[indice]
    concatenandoConv = str(concatenando)
    concatCidade= concatenandoConv.upper()
    concatCidade = concatCidade.replace('Ã', 'A')
    concatCidade = concatCidade.replace('Á', 'A')
    concatCidade = concatCidade.replace('Ó', 'O')
    concatCidade = concatCidade.replace('Õ', 'O')
    concatCidade = concatCidade.replace('É', 'E')
    concatCidade = concatCidade.replace('Ç', 'C')
    concatCidade = concatCidade.replace('Í', 'I')
    concatCidade = concatCidade.replace('Â', 'A')
    concatCidade = concatCidade.replace('Ê', 'E')
    concatCidade = concatCidade.replace('Ú', 'U')
    tamanho = len(concatCidade)
    if tamanho <=40:
        newTamanho = 40 - tamanho
        cespacos = cespacos = (' ') * newTamanho
        finalCIDADE = concatCidade + cespacos
    else:
        finalCIDADE = concatCidade[:40]
    # ********************************************************************************
    uf = df['UF'].values[indice]
    cep = df['CEP (SEM TRAÇO)'].values[indice]
    newCEP = str(cep)
    newCEP= newCEP.upper()
    newCEP = newCEP.replace('-', '')
    newCEP = newCEP.replace('Ã', 'A')
    newCEP = newCEP.replace('Á', 'A')
    newCEP = newCEP.replace('Ó', 'O')
    newCEP = newCEP.replace('Õ', 'O')
    newCEP = newCEP.replace('É', 'E')
    newCEP = newCEP.replace('Ç', 'C')
    newCEP = newCEP.replace('Í', 'I')
    newCEP = newCEP.replace('Â', 'A')
    newCEP = newCEP.replace('Ê', 'E')
    newCEP = newCEP.replace('Ú', 'U')
    tamanho = len(newCEP)
    if tamanho <= 7:
        sei = uf +'0'+ newCEP
    else:
        sei = uf+newCEP
    # *******************************************************************************
    econcatenado = df['E-MAIL'].values[indice]
    econcatenandoConv = str(econcatenado)
    econcatenandoConv= econcatenandoConv.upper()
    econcatenandoConv = econcatenandoConv.replace('Ã', 'A')
    econcatenandoConv = econcatenandoConv.replace('Á', 'A')
    econcatenandoConv = econcatenandoConv.replace('Ó', 'O')
    econcatenandoConv = econcatenandoConv.replace('Õ', 'O')
    econcatenandoConv = econcatenandoConv.replace('É', 'E')
    econcatenandoConv = econcatenandoConv.replace('Ç', 'C')
    econcatenandoConv = econcatenandoConv.replace('Í', 'I')
    econcatenandoConv = econcatenandoConv.replace('Â', 'A')
    econcatenandoConv = econcatenandoConv.replace('Ê', 'E')
    econcatenandoConv = econcatenandoConv.replace('Ú', 'U')
    etamanho = len(econcatenandoConv)
    ecaracs = 200 - etamanho
    eEspacos = ' ' * ecaracs
    if etamanho ==0:
        final3 = (' ') * 200
    else:final3 = econcatenandoConv + eEspacos
    ultimo = sei + final3
    df['new_NÚMERO CPF / CNPJ']='s'
    df['new_NOME CLIENTE / RAZÃO SOCIAL']='s'
    df['new_ENDEREÇO']='s'
    df['new_BAIRRO']='s'
    df['new_CIDADE']='s'
    df['new_UF']='s'

    df['new_NÚMERO CPF / CNPJ'].values[indice] = concatenadoCPF
    df['new_NOME CLIENTE / RAZÃO SOCIAL'].values[indice]=finalNome
    df['new_ENDEREÇO'].values[indice]=finalENde
    df['new_BAIRRO'].values[indice]=concatatenadoBAIRRO
    df['new_CIDADE'].values[indice]=finalCIDADE
    df['new_UF'].values[indice]=ultimo
    filler6 = (' ')*6
    filler20 =(' ')*20
    filler15=(' ')*15
    real = concatenadoCPF + finalNome + finalENde +filler6+filler20 + concatatenadoBAIRRO+filler15 + finalCIDADE + ultimo+'\n'

    file.write(real)

    indice+=1



file.close()