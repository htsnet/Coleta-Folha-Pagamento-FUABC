#!/usr/bin/env python
# coding: utf-8
# ## Rotina para localizar nome e salários do mês da FUABC

#define ano, mês a área bem como o nome dos arquivos PDF baixados
ano = 2020
mes = 3
area = 90 #constante para indicar a FUABC no banco de dadaos do Cidade em Números https://cidadeemnumeros.com.br/
apagaAnteriores = False

print(ano, mes, area)

#pede o ano e mês da Folha
ano = int(input("Digite o ano de referência: "))
mes = int(input("Digite o mes de referência: "))

#importando o módulo
import pdfminer
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from io import StringIO
from pdfminer.layout import LAParams
from pdfminer.converter import TextConverter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
import pandas as pd
import os
import glob
import conexaoMysql



#para controlar leitura de registros antes só no primeiro arquivo
primeiroArquivo = True

#busca a lista de arquivos com os funcionários (cargos e nomes) na pasta
input_list = glob.glob('*cargos_funcionarios.pdf')

for arq in input_list:
    arquivoNomes = arq
    #busca o correspondente com os salários
    nomeParcial = arq.replace("cargos_funcionarios.pdf", "")
    arqSalarios = glob.glob(nomeParcial + 'cargos_salarios.pdf')
    arquivoSalarios = arqSalarios[0]

    print(arquivoNomes, arquivoSalarios)


    # ## Nomes e matrículas
    #criando um objeto leitor
    parser = PDFParser(open(arquivoNomes, 'rb'))
    document = PDFDocument(parser)


    # Try to parse the document
    if not document.is_extractable:
        raise PDFTextExtractionNotAllowed


    # Create a PDF resource manager object 
    # that stores shared resources.
    rsrcmgr = PDFResourceManager()
    # Create a buffer for the parsed text
    retstr = StringIO()
    # Spacing parameters for parsing
    laparams = LAParams()
    codec = 'utf-8'


    # Create a PDF device object
    device = TextConverter(rsrcmgr, retstr, 
            #codec = codec, 
            laparams = laparams)


    # Create a PDF interpreter object
    interpreter = PDFPageInterpreter(rsrcmgr, device)


    # Process each page contained in the document.
    for page in PDFPage.create_pages(document):
        interpreter.process_page(page)


    records = []
            
    lines = retstr.getvalue().splitlines()
    for line in lines:
        records.append(line)


    pessoas = pd.DataFrame(columns=['EMPRESA','DEPARTAMENTO','MATRICULA','NOME','CARGO','HORAS MES'])

    #prepara e identifica cada coluna para gera um dataframe
    pessoas = pd.DataFrame(columns=['EMPRESA','DEPARTAMENTO','MATRICULA','NOME','CARGO','HORAS MES'])
    pessoas = pessoas.append({'EMPRESA': '','DEPARTAMENTO': '','MATRICULA': 0,'NOME': '','CARGO': '','HORAS MES': ''}, ignore_index=True)
    coluna = -1
    linhaAnterior = 1
    linha = 0
    for campo in records:
        if campo == 'EMPRESA':
            coluna = 0
            linha = linhaAnterior
        elif campo == 'DEPARTAMENTO':
            coluna = 1
            linha = linhaAnterior
        elif campo == 'MATRICULA':
            coluna = 2
            linha = linhaAnterior
        elif campo == 'NOME':
            coluna = 3
            linha = linhaAnterior
        elif campo == 'CARGO':
            coluna = 4
            linha = linhaAnterior
        elif campo == 'HORAS MES':
            coluna = 5
            linha = linhaAnterior
        elif 'SIGA/' in campo:
            linhaAnterior = linha
            coluna = -1
        elif 'Hora:' in campo:
            coluna = -1
            linhaAnterior = linha
        else: 
            if coluna >= 0:
                #print(linha, coluna, campo)
                #print(pessoas.shape[0])
                if pessoas.shape[0] <= linha +1:
                    pessoas = pessoas.append({'EMPRESA': '','DEPARTAMENTO': '','MATRICULA': 0,'NOME': '','CARGO': '','HORAS MES': ''}, ignore_index=True)
                    
                #verifica se é campo matricula com valor branco
                if coluna == 2 and campo == '':
                    pessoas.iloc[linha,coluna] = 0
                else:
                    pessoas.iloc[linha,coluna] = campo
                linha = linha + 1
            
    pessoas.tail()

    pessoas = pessoas.drop(pessoas[(pessoas.MATRICULA == 0)].index)

    pessoas.describe()

    #criando um objeto leitor
    parser = PDFParser(open(arquivoSalarios, 'rb'))
    document = PDFDocument(parser)

    # Try to parse the document
    if not document.is_extractable:
        raise PDFTextExtractionNotAllowed

    # Create a PDF resource manager object 
    # that stores shared resources.
    rsrcmgr = PDFResourceManager()
    # Create a buffer for the parsed text
    retstr = StringIO()
    # Spacing parameters for parsing
    laparams = LAParams()
    codec = 'utf-8'


    # Create a PDF device object
    device = TextConverter(rsrcmgr, retstr, 
            #codec = codec, 
            laparams = laparams)

    # Create a PDF interpreter object
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # Process each page contained in the document.
    for page in PDFPage.create_pages(document):
        interpreter.process_page(page)

    #print(rsrcmgr)        
    #retira trocas de linhas que possam ter entrado
    #lines = lines.replace('\n', '')


    records = []
    lines = retstr.getvalue().splitlines()
    for line in lines:
        records.append(line)

    #print(records)
    
    #prepara e identifica cada coluna para gera um dataframe
    pessoas['SIT. FOLHA'] = ''
    pessoas['SALARIO BASE'] = 0.0
    pessoas['LIQ.A RECEBER'] = 0.0
    pessoas['TOTAL BRUTO'] = 0.0
    pessoas.head()


    def trocaPontoVirgula(valor):
        final = valor.replace('.', '')
        final = final.replace(',', '.')
        return final

    def retiraAspas(valor):
        final = valor.replace("'", "\\'")
        final = final.replace('"', '\\"')
        return final

    #prepara e identifica cada coluna para gera um dataframe
    auxiliarSalarios = pd.DataFrame(columns=['MATRICULA','CARGO','SIT. FOLHA','SALARIO BASE','LIQ.A RECEBER','TOTAL BRUTO'])
    auxiliarSalarios = auxiliarSalarios.append({'MATRICULA': 0,'CARGO': '','SIT. FOLHA': '','SALARIO BASE': 0.0,'LIQ.A RECEBER': 0.0,'TOTAL BRUTO': 0.0}, ignore_index=True)
    coluna = -1
    linhaAnterior = 1
    linha = 0
    maiorLinha = 0
    for campo in records:
        #campo = campo.replace('\n', '')
        #print(campo)
        if campo == 'SIT. FOLHA SALARIO BASE': #atenção para o título, pq juntou com a coluna seguinte
            coluna = 2
            linha = linhaAnterior
        elif campo == 'SALARIO BASE':
            coluna = 3
            linha = linhaAnterior
        elif campo == 'LIQ.A RECEBER':
            coluna = 4
            linha = linhaAnterior
        elif campo == 'TOTAL BRUTO':
            coluna = 5
            linha = linhaAnterior
        elif campo == 'MATRICULA':
            coluna = 0
            linha = linhaAnterior
        elif campo == 'CARGO':
            coluna = 1
            linha = linhaAnterior
        elif 'SIGA/' in campo:
            linhaAnterior = linha
            coluna = -1
        elif 'Hora:' in campo:
            coluna = -1
            linhaAnterior = linha
        elif 'Folha:' in campo:
            coluna = -1
            linhaAnterior = linha        
        else: 
            if coluna >= 0:
                final = campo
                
                #retira o caracter | do campo da matrícula    
                if coluna == 0 and campo != '' and campo is not None:
                    ###print(valor)
                    valor = campo.split('|')
                    ###print(len(valor))
                    if len(valor) > 0:
                        final = valor[1]
                    else:
                        final = 0
                else:
                    final = campo
                    
                #acerta valor numérico (. e ,)    
                if coluna == 3 or coluna == 4 or coluna == 5:
                    final = trocaPontoVirgula(campo)
                    
                #só usa o primeiro grupo deste campo, pois juntou com o salário base
                if coluna == 2:
                    if campo.isalpha() and campo != '':
                        final = campo
                    else:
                        final = -1
                
                
                #se final for branco, não faz nada
                #print(final)
                if final != -1:
                    if auxiliarSalarios.shape[0] <= linha +1:
                        auxiliarSalarios = auxiliarSalarios.append({'MATRICULA': 0,'CARGO': '','SIT. FOLHA': '','SALARIO BASE': 0.0,'LIQ.A RECEBER': 0.0,'TOTAL BRUTO': 0.0}, ignore_index=True)

                    auxiliarSalarios.iloc[linha,coluna] = final
                    linha = linha + 1

    #retira as linhas sem informação na matrícula
    auxiliarSalarios = auxiliarSalarios.drop(auxiliarSalarios[(auxiliarSalarios.MATRICULA == '')].index)
    auxiliarSalarios = auxiliarSalarios.drop(auxiliarSalarios[(auxiliarSalarios.MATRICULA == 0)].index)


    # ## Faz a junção dos 2 dataframes usando a matrícula como chave
    for index, linha in auxiliarSalarios.iterrows():
        matricula = linha['MATRICULA']
        #procura a matrícula na tabela principal (pessoas)
        indice = pessoas[pessoas['MATRICULA'] == matricula].index
        indiceObtido = indice[0]
        #print(auxiliarSalarios.loc[[index]])
        #print(matricula, indiceObtido, index)
        #pega os dados a adicionar
        cargo = auxiliarSalarios.at[index,'CARGO']
        situacao = auxiliarSalarios.at[index, 'SIT. FOLHA']
        #print(auxiliarSalarios.at[index, 'LIQ.A RECEBER'])
        liquido = auxiliarSalarios.at[index, 'LIQ.A RECEBER']
        bruto = auxiliarSalarios.at[index, 'TOTAL BRUTO']
        #print(liquido, bruto)
        #pessoas.at[indiceObtido, 6] = cargo
        pessoas.at[indiceObtido, 'SIT. FOLHA'] = situacao
        pessoas.at[indiceObtido, 'LIQ.A RECEBER'] = liquido
        pessoas.at[indiceObtido, 'TOTAL BRUTO'] = bruto
        
    # ## Grava no Banco de Dados
    # 
    conta = 0
    try:

        
        if primeiroArquivo == True:
            #verifica quantos registros tem antes de iniciar
            sql = """select count(*) from CEN_tVencimentos where VencimentosArea = {0} and VencimentosAno = {1} and VencimentosMes = {2}""".format(area, ano, mes)
            #print(sql)
            cursor = connection.cursor()
            cursor.execute(sql)
            qtde = cursor.fetchone()
            print("Registros antes", qtde)
            primeiroArquivo = False
        
        #insere os registros encontrados
        primeiro = True
        for index, linha in pessoas.iterrows():
            matricula = linha['MATRICULA']
            cargo = retiraAspas(linha['CARGO'])
            empresa = retiraAspas(linha['EMPRESA'])
            departamento = retiraAspas(linha['DEPARTAMENTO'])
            nome = retiraAspas(linha['NOME'])
            horas = linha['HORAS MES'].strip()
            situacao = retiraAspas(linha['SIT. FOLHA'])
            #base = linha['SALARIO BASE']
            liquido = linha['LIQ.A RECEBER']
            bruto = linha['TOTAL BRUTO']
            
            #só no primeiro registro, depois de pegar o nome da empresa
            if primeiro == True and apagaAnteriores == True:
                #apaga os registros desta área e "empresa" do mesmo ano/mês para não duplicar os registros
                sql = """delete from CEN_tVencimentos where VencimentosArea = {0} and VencimentosAno = {1} and VencimentosMes = {2} and VencimentosSecretaria = '{3}'""".format(area, ano, mes, empresa)
                #print(sql)
                print("Deletando os registros anteriores.")
                cursor = connection.cursor()
                cursor.execute(sql)
            
            primeiro = False
            
            
            sql = """insert into CEN_tVencimentos (VencimentosAno, VencimentosMes, VencimentosMatricula, VencimentosNome, VencimentosArea, VencimentosSecretaria,
                VencimentosCargo, VencimentosBruto, VencimentosLiquido, VencimentosHorasBase, VencimentosSituacao) values 
                ({0}, {1}, {2}, '{3}', {4}, '{5}', '{6}', {7}, {8}, {9}, '{10}')""".format(ano, mes, matricula, nome, area, empresa, cargo, bruto, liquido, horas, situacao)
            #print(sql)
            cursor = connection.cursor()
            cursor.execute(sql)
            #print(cursor.rowcount, "Registro inserido")
            conta = conta + 1
            #print(conta)
            

        print("Total de registros inseridos: {0}".format(conta))    
        
        #verifica quantos registros tem depois de executar
        sql = """select count(*) from CEN_tVencimentos where VencimentosArea = {0} and VencimentosAno = {1} and VencimentosMes = {2}""".format(area, ano, mes)
        cursor = connection.cursor()
        cursor.execute(sql)
        qtde = cursor.fetchone()
        print("Registros depois: ", qtde, "\n")
            
    except mysql.connector.Error as error:
        print("Falha na inclusão {}".format(error))        

    connection.commit()
    cursor.close()    

input("Processamento Encerrado. Tecla algo para fechar.")
