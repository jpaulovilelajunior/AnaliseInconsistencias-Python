import mysql.connector
from mysql.connector import errorcode
import paramiko

class Conecta_Banco:
# Inicializando a classe
    def __init__(self,host,user,password,port,database):
        self.__host = host
        self.__user = user
        self.__password = password
        self.__port = port
        self.__database = database

    #Faz a busca no banco de dados da SEDI
    def Analisar_sedi(self):
        #conecta no Banco de dados
        conn = self.__Verifica_Conexao_SQL()
        if conn == None:
            print("impossivel conectar")
            consultaBanco = None
            nomesColunas = None
            return consultaBanco, nomesColunas  # retorna a consulta e o nome das colunas
            pass
        else:
            linkadoBanco = conn.cursor()
            try:
                linkadoBanco.execute(self.__Analisa_ComponentesSEDI()) #executa o sql .txt
                print("Realizando consulta no banco")
            except:
                linkadoBanco.fetchall()
                pass
            nomesColunas = linkadoBanco.column_names #retorna o nome das colunas para inserirmos no excel posteriormente
            consultaBanco = linkadoBanco.fetchall() #fecha a consulta e salva
            print("Consulta realizada com sucesso")
            return consultaBanco,nomesColunas #retorna a consulta e o nome das colunas

    #Faz a busca no banco de dados da SER
    def Analisar_ser(self):
        #tenta conectar ssh
        ssh = self.__Verifica_Conexao_SSH()
        if (ssh != None):
            conn = self.__Verifica_Conexao_SQL()
            if conn == None:
                print("impossivel conectar")
                consultaBanco = None
                nomesColunas = None
                return consultaBanco, nomesColunas  # retorna a consulta e o nome das colunas
                pass
            else:
                linkadoBanco = conn.cursor()
                try:
                    linkadoBanco.execute(self.__Analisa_ComponentesSER()) #executa o sql .txt
                    print("Realizando consulta no banco")
                except:
                    linkadoBanco.fetchall()
                    pass
            nomesColunas = linkadoBanco.column_names #retorna o nome das colunas para inserirmos no excel posteriormente
            consultaBanco = linkadoBanco.fetchall() #fecha a consulta e salva
            print("Consulta realizada com sucesso")
            return consultaBanco,nomesColunas #retorna a consulta e o nome das colunas
        else:
            consultaBanco = None
            nomeColuna = None
            return consultaBanco,nomeColuna

    #faz a leitura do .txt onde está o código sql SEDI e retorna a leitura.
    def __Analisa_ComponentesSEDI(self):
        try:
            with open(r".\Querys\controleOfertaFrequenciaSEDI.txt", encoding= 'utf8') as arquivo:
                query = arquivo.read()
                arquivo.close()
        except (FileNotFoundError,FileExistsError) as erro:
            print(erro)
            print("Verifique caminho/nome do arquivo")
            return None
        return query

    #faz a leitura do .txt onde está o código sql SEDI e retorna a leitura.
    def __Analisa_ComponentesSER(self):
        try:
            with open(r".\Querys\controleOfertaFrequenciaSER.txt", encoding= 'utf8') as arquivo:
                query = arquivo.read()
                arquivo.close()
        except (FileNotFoundError,FileExistsError) as erro:
            print(erro)
            print("Verifique caminho/nome do arquivo")
            return None
        return query

    #tenta realizar a conexão no banco de dados.
    def __Verifica_Conexao_SQL(self):
        try:
            conn = mysql.connector.connect(host=self.host,
                                           user=self.user,
                                           password=self.password,
                                           port=self.__port,
                                           database=self.database)
        except:
            print("Não foi possível conectar ao banco (Verifique VPN ou dados)")
            return None
        print('Conexão realizada')
        return conn

    #tenta realizar conexão via SSH
    def __Verifica_Conexao_SSH(self):
        #tenta realizar conexão via SSH com os dados pré estabelecidos
        #Os dados foram modificados para preservar a conexão SSH
        try:
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect(hostname=locaQueEstáConectado,
            port=numeroPorta,
            username=usuarioSSH,
            password=senhaSSH)
        except:
            print("Verificar dados de conexão SSH")
            ssh = None
            return ssh
        print("Conectando ao tunel SSH")
        return ssh

#Setters e Getters dos atributos privados --------------------------
    @property
    def host(self):
        return self.__host
    @property
    def user(self):
        return self.__user
    @property
    def password(self):
        return self.__password
    @property
    def port(self):
        return self.__port
    @property
    def database(self):
        return self.__database
#Setters e Getters dos atributos privados -------------------------- FIM
