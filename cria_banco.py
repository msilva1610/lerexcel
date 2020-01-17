# connect_db.py
# 01_create_db.py
import sqlite3
print ("Criando banco....")
conn = sqlite3.connect('Projetos.db')
conn.close()
print ("Banco Projetos criado.")
