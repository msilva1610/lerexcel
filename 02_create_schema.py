# 02_create_schema.py
import sqlite3

# conectando...
print ("Conectando ...")
conn = sqlite3.connect('Projetos.db')
# definindo um cursor
cursor = conn.cursor()

# criando a tabela (schema)
cursor.execute("""
CREATE TABLE projetos (
        id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
        Nome_Projeto TEXT NOT NULL,
        VP VARCHAR(50),
        Formula_Fase VARCHAR(50),
        an VARCHAR (100),
        gp VARCHAR (50),
        LT VARCHAR(50),
        Lider_Teste VARCHAR (50),
        Gerente_Desenv VARCHAR (50),
        Formula_Status_Projeto VARCHAR(50),
        Descricao TEXT,
        Nome_Arquivo VARCHAR (100),
        Ini_desenv VARCHAR (300),
        Term_desenv VARCHAR (300),
        Completude1 VARCHAR(4),
        Ini_Teste_Integrado VARCHAR (300),
        Term_Teste_Integrado VARCHAR (300),
        Completude2 VARCHAR(4),
        Ini_hml VARCHAR (300),
        Fim_hml VARCHAR (300),
        Completude3 VARCHAR(4),
        Problema_Risco TEXT,
        SubCausa TEXT,
        Plano_Acao TEXT,
        causa TEXT,
        id_arquivo integer
);
""")

print('Tabela criada com sucesso.')
# desconectando...
conn.close()
