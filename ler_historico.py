# -*- coding: utf-8 -*
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
import difflib
import sqlite3
import datetime
import sys
import glob, os
import os.path
import platform
import shutil
from datetime import datetime


def main():
        ler_historico()

def ler_historico():
        EVENTO = 700
        ID_ARQUIVO = 16
        ID_PROJETO = 418

        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

        db_path = os.path.join(BASE_DIR, "Projetos.db")

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute("""
        SELECT id,
        valor_anterior,
        novo_valor
        FROM historico
        where cod_evento = ?
        and id_arquivo = ?
        and id_projeto = ?
        """,(EVENTO,ID_ARQUIVO,ID_PROJETO,))

        for ID, VALOR_ANTERIOR, NOVO_VALOR in cursor.fetchall():
            #valor1 = str(VALOR_ANTERIOR.splitlines())
            #valor2 = str(NOVO_VALOR.splitlines())
            valor1 = str(VALOR_ANTERIOR)
            valor2 = str(NOVO_VALOR)

            pesquisar = valor1[:10]
            posicao_fim = valor2.find(pesquisar)
            texto_novo = valor2[:posicao_fim]
            print (texto_novo)

        print ("Fim...")

        if conn:
                conn.close()

if __name__ == "__main__":
	main()
