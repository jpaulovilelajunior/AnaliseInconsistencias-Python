from secretariasAcesso import conexaoSecretarias
import pandas as pd

if __name__ == '__main__':
    #ser = conexaoSecretarias("Relatorio Erros SER.xlsx")
    #ser.conectar_Ser()
    sedi = conexaoSecretarias("Relatorio Erros SEDI.xlsx")
    sedi.conectar_Sedi()





