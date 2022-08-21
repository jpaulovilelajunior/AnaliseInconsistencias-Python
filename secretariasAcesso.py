import numpy as np
import pandas as pd
import openpyxl
from openpyxl.worksheet.dimensions import ColumnDimension,DimensionHolder
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font,Border,alignment
from openpyxl import load_workbook
import datetime
from acessoBanco import Conecta_Banco

class conexaoSecretarias:

    def __init__(self,nomeExcel):
        self.__nomeExcel = nomeExcel

    #pega a consulta realizada no banco e passa para o dataframe. Salva no excel.
    def __DataFrame_Excel(self,dados,nomeColuna):
        #verifica se os dados não são nulos,se não houve erros
        if ((dados != None) or (nomeColuna != None)):
            #converte em DataFrame do pandas
            dfGeral = pd.DataFrame(data = dados, index= None, columns= nomeColuna)

            #cria a coluna de carga horária executada
            dfGeral.insert(loc=22,column="C.H.A.Exe",value=(dfGeral['carga_hr_comp']-dfGeral['carga_hr_exec']))
            #substitui os valores vazios com 0
            dfGeral["C.H.A.Exe"] = dfGeral["C.H.A.Exe"].fillna(0)

            print("Analisando inconsistências")
            #analisa os componentes e retorna as inconsistências de prestação de contas
            dfAnalise = self.__Verifica_inconsistencias(dataFrame=dfGeral)
            dfPedagogico = self.__Verifica_pedagogico(dataFrame=dfGeral)
            erroEscolasAnalise,erroCidadesAnalise,erroTiposAnalise = self.__Porcentagem_inconsistencias(dataFrame=dfAnalise)
            erroEscolasPeda, erroCidadesPeda, erroTiposPeda = self.__Porcentagem_inconsistencias(
                dataFrame=dfPedagogico)

            print("Salvando arquivo Excel")
            #salva no arquivo
            dfGeral.to_excel(self.__nomeExcel, sheet_name="Base Dados", index=False, float_format='.1f')
            with pd.ExcelWriter(self.__nomeExcel, mode='a',if_sheet_exists='overlay') as writer:
                dfAnalise.to_excel(writer, sheet_name="Analise Metas",index=False, float_format='.1f')
                dfPedagogico.to_excel(writer, sheet_name="Analise Pedag.",index=False, float_format='.1f')
                erroEscolasAnalise.to_excel(writer,sheet_name="% Incons. Metas", index=False)
                erroCidadesAnalise.to_excel(writer,sheet_name="% Incons. Metas", index=False,startcol=3)
                erroTiposAnalise.to_excel(writer,sheet_name="% Incons. Metas", index=False,startcol=6)
                erroEscolasPeda.to_excel(writer,sheet_name="% Incons. Pedagogico", index=False)
                erroCidadesPeda.to_excel(writer,sheet_name="% Incons. Pedagogico", index=False,startcol=3)
                erroTiposPeda.to_excel(writer,sheet_name="% Incons. Pedagogico", index=False,startcol=6)

            print("Arquivo Gerado com sucesso")
        else:
            print("Verifique a conexão do banco de dados informado")

    #verifica as inconsistências de metas físicas das componentes (regra de negócios)
    def __Verifica_inconsistencias(self,dataFrame):

        hoje = datetime.datetime.today()
        hoje = hoje.strftime('%Y/%m/%d')

        #Verifica o componente que está com carga horária executada = 0 e está com status ABERTO
        statusAberCHAE0 = dataFrame[(dataFrame['status_comp'] == 'ABERTO') & (dataFrame['carga_hr_comp'] == dataFrame['carga_hr_exec'])]
        statusAberCHAE0 = statusAberCHAE0.assign(Erro='Componente com status EM ABERTO e CH.Exec finalizada')

        #Verifica se está com status CURSANDO e não possui mas carga horária a ser executado
        statusCursandoCHOk = dataFrame[
            (dataFrame['status_comp'] == 'EM ANDAMENTO') & (dataFrame['carga_hr_comp'] == dataFrame['carga_hr_exec'])]
        statusCursandoCHOk = statusCursandoCHOk.assign(Erro='Componente com status EM ANDAMENTO e CH.Exec Finalizada')

        #Verifica se o status está como CONCLUÍDO e ainda possui carga horária a executar
        statusConclNOK = dataFrame[(dataFrame['status_comp'] == 'CONCLUIDO') & (dataFrame['C.H.A.Exe'] > 0)]
        statusConclNOK = statusConclNOK.assign(Erro='Componente com status CONCLUIDO e CH a executar')

        #verifica se o status está ABERTO e já possuí carga horária executada
        statusAbertoExecutado = dataFrame[
            (dataFrame['status_comp'] == 'ABERTO') & (dataFrame['C.H.A.Exe'] > 0)]
        statusAbertoExecutado = statusAbertoExecutado.assign(Erro='Componente com status ABERTO e com CH executada')

        #verifica os componentes cancelados
        #statusCancelado = valoresExcel[(valoresExcel['status_comp'] == 'CANCELADO')]
        #statusCancelado = statusCancelado.assign(Erro='Componente CANCELADO')

        #verifica o status ABERTO, se tem professor e já deveria estar iniciado
        statusAbertoDataNOK = dataFrame[
            (dataFrame['status_comp'] == 'ABERTO') &
            (pd.to_datetime(dataFrame['previsao_inicio_comp']) < hoje) &
            (dataFrame['qtde_matriculas'] > 0) & (dataFrame['professor_nome'])]
        statusAbertoDataNOK = statusAbertoDataNOK.assign(Erro='Componente já deveria estar Iniciado')

        #verifica se o status está EM ANDAMENTO e já deveria estar fianlizado (e tem carga horária executada)
        statusEmAndamento = dataFrame[
            (dataFrame['status_comp'] == 'EM ANDAMENTO') &
            (pd.to_datetime(dataFrame['previsao_termino_comp']) < hoje) &
            (dataFrame['C.H.A.Exe'] > 0)]
        statusEmAndamento = statusEmAndamento.assign(Erro = 'Componente com status EM ANDAMENTO e já deveria estar com CHE. finalizada')

        #verifica se não tem data inserida nas componentes
        datasNulas = dataFrame[dataFrame['previsao_inicio_comp'].isnull()]
        datasNulas = datasNulas.assign(Erro='Componente sem data de Início/Fim cadastrados')

        #compila os erros em um novo dataFrame
        errosCompilados= pd.concat([statusAberCHAE0,statusCursandoCHOk,statusConclNOK,statusAbertoExecutado,
                                    statusAbertoDataNOK,datasNulas,statusEmAndamento],ignore_index=True)
        return errosCompilados

    #Verifica as inconsistências pedagógicos de planejamento
    def __Verifica_pedagogico(self,dataFrame):

        hoje = datetime.datetime.today()

        #verifica se primeira/ultima aula estão batendo com primeiro/ultimo dia de lançamento
        lancamentoDiario = dataFrame[(dataFrame['modalidade'] == 'PRESENCIAL')
                                        & (dataFrame['status_comp'] == 'CONCLUIDO')
                                        & ((dataFrame['previsao_inicio_comp'] != dataFrame['primeira_aula'])
                                           | (dataFrame['previsao_termino_comp'] != dataFrame['ultima_aula']))]
        lancamentoDiario = lancamentoDiario.assign(Verificar = 'Verificar Primeiro/Ultimo lançamento em diário')

        #verifica a previsão de término, se está com planejamento ok
        valores = dataFrame[(dataFrame['status_comp'] == 'EM ANDAMENTO') &
                                 (dataFrame['carga_hr_exec']>0) &
                                 (pd.to_datetime(dataFrame['previsao_termino_comp'])>hoje)]

        previsaoTermino = valores[((pd.to_datetime(valores['previsao_termino_comp']).apply(lambda x: (x - hoje).days))
                                        < valores['C.H.A.Exe'].apply(lambda x: int(round(x/4))))]
        qntDiasFaltantes = previsaoTermino['C.H.A.Exe'].apply(lambda x: int(round(x / 4))) - pd.to_datetime(previsaoTermino[
            'previsao_termino_comp']).apply(lambda x: (x - hoje).days)
        #monta a string
        qntDiasFaltantes = qntDiasFaltantes.to_frame(name= 'dias')
        qntDiasFaltantes['dias'] = qntDiasFaltantes['dias'].astype(str)
        qntDiasFaltantes['texto1'] = 'Verificar - Prev. Término não condiz com CHA.Executar. Faltando: '
        qntDiasFaltantes['texto2'] = ' dias.'
        previsaoTermino['Verificar'] = qntDiasFaltantes['texto1'] + qntDiasFaltantes['dias'] + qntDiasFaltantes['texto2']

        errosCompilados = pd.concat([lancamentoDiario,previsaoTermino],
                                    ignore_index=True)
        return errosCompilados

    #Faz a porcentagem de inconsistências
    def __Porcentagem_inconsistencias(self,dataFrame):

        erroEscolas = pd.DataFrame(dataFrame.iloc[:,3].value_counts(normalize=True)*100).reset_index()
        erroCidades = pd.DataFrame(dataFrame.iloc[:,4].value_counts(normalize=True)*100).reset_index()
        erroTipos = pd.DataFrame(dataFrame.iloc[:,26].value_counts(normalize=True)*100).reset_index()
        erroEscolas.columns=['Escola','%Erro Escola']
        erroCidades.columns=['Cidade','%Erro Cidade']
        erroTipos.columns=['Erro','%Tipos de Erro']

        return erroEscolas,erroCidades,erroTipos

    def conectar_Sedi(self):
        # realiza a conexão no banco SIGA - Produção SEDI e retorna com os dados.
        consultaComponentes = Conecta_Banco(host='10.6.62.95',user='powerbi',password='VYecctxUF7ytW!',port='3306',database='siga_producao')
        dados,nomeColuna = consultaComponentes.Analisar_sedi()
        self.__DataFrame_Excel(dados=dados,nomeColuna=nomeColuna)

    def conectar_Ser(self):
    # realiza a conexão no banco SIGA - Produção SEDI e retorna com os dados.
        consultaComponentes = Conecta_Banco(host="127.0.0.1",user= "joao_sge",password="fK1ejcOIqvtz59yJnn1h",port="3306",database="c3siga")
        dados,nomeColuna = consultaComponentes.Analisar_ser()
        self.__DataFrame_Excel(dados=dados,nomeColuna=nomeColuna)