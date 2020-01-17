# -*- coding: utf-8 -*
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
import sqlite3
import datetime
import sys
import glob, os
import os.path
import platform
import shutil
from datetime import datetime

def main():
    projeto = "16.0213.1.TN-Automatização da Devolução de Valores para Clientes Prospect - Solução Definitiva (Online)"
    ListaDeRms = RetornaRms(projeto)
    for cadalinha in ListaDeRms:
        print (cadalinha[1])

        
def RetornaRms(Nome_do_Projeto):
    try:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        db_path = os.path.join(BASE_DIR, "Projetos.db")
            
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        Projeto = str(Nome_do_Projeto[0:7]).strip()
        print (Projeto)

        cursor.execute("""
            SELECT 
                   Numero_RM,
                   Resumo_Descricao_RM,
                   Responsavel,
                   Data_criacao,
                   DataComite,
                   Sistema,
                   DataIniQA,
                   DataFimQA,
                   DataIniHML,
                   DataFimHML,
                   Status
              FROM RMS
              where TicketdeSistemaExterno like ?
            """, (str("%" + Projeto + "%"),))

        ListaDeRms = cursor.fetchall()
    
        return ListaDeRms

        conn.commit()
    except Exception as e:
        raise
    finally:
        conn.close()

if __name__ == "__main__":
    main()
