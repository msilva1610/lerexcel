# 04_create_schema_historico.py
import sqlite3

# conectando...
print ("Conectando ...")
conn = sqlite3.connect('Projetos.db')
# definindo um cursor
cursor = conn.cursor()

# criando a tabela (schema)
cursor.execute("""
CREATE TABLE historico(
        id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
        id_projeto INTEGER NOT NULL,
        id_arquivo INTEGER NOT NULL,
        cod_evento INTEGER NOT NULL,
        evento varchar(100),
        valor_anterior TEXT,
        novo_valor TEXT
);
""")

conn.commit()

print('Tabela tb_arquivos criada com sucesso.')
# desconectando...
conn.close()
