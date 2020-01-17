# 10_create_schema_tb_rms.py
import sqlite3

# conectando...
print ("Conectando ...")
conn = sqlite3.connect('Projetos.db')
# definindo um cursor
cursor = conn.cursor()

# criando a tabela (schema)
cursor.execute("""
CREATE TABLE RMS(
    ID INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
    Numero_RM integer,
    Resumo_Descricao_RM varchar (2000),
    Responsavel varchar (200),
    Data_criacao varchar (10),
    TipoDaMudanca varchar (50),
    Diretoria varchar (50),
    Sistema varchar (200),
    Origem varchar (200),
    DataIniDesenv varchar (10),
    DataFimDesenv varchar (10),
    DataIniQA varchar (10),
    DataFimQA varchar (10),
    DataIniHML varchar (10),
    DataFimHML varchar (10),
    DataIniPrd varchar (10),
    DataFimPrd varchar (10),
    Status varchar (200),
    DataComite varchar (10),
    Motivo varchar (2000),
    CenarioNegAtual varchar(2000),
    CenarioNegDesejado varchar (2000),
    CenarioTecProposto varchar (2000),
    TicketdeSistemaExterno varchar (200)
);
""")

conn.commit()

print('Tabela rms criada com sucesso.')
# desconectando...
conn.close()


