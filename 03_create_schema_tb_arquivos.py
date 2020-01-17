# 02_create_schema_tb_arquivos.py
import sqlite3

# conectando...
print ("Conectando ...")
conn = sqlite3.connect('Projetos.db')
# definindo um cursor
cursor = conn.cursor()

# criando a tabela (schema)
cursor.execute("""
CREATE TABLE arquivos(
        id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
        NomeDoArquivo varchar(300) NOT NULL,
        DataCriacaoDoArquivo datetime,
        TotalDeLinhas integer
);
""")

conn.commit()

print('Tabela tb_arquivos criada com sucesso.')
# desconectando...
conn.close()
