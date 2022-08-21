from secretariasAcesso import conexaoSecretarias
import pandas as pd

if __name__ == '__main__':
    #Base de conexões voltadas para a Secretaria da Retomada de Goiás - SER
    ser = conexaoSecretarias("Relatorio Erros SER.xlsx")
    ser.conectar_Ser()
    #Base de conexões voltadas para a Secretaria de Desenvolvimento de Goiás - SEDI
    sedi = conexaoSecretarias("Relatorio Erros SEDI.xlsx")
    sedi.conectar_Sedi()





