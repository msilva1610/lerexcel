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
        lerdata_identifica_qtde()

def lerdata_identifica_qtde():
        juliana = "Juliana Alves Castro Perez"
        leonardo = "Leonardo Augusto Mendes Leandro"
        requested = "Requested"
        at_risk = "At Risk"
        delayed = "Delayed"
        fora_pipeline = "fora pipeline"        

        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

        db_path = os.path.join(BASE_DIR, "Projetos.db")

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute("""
        SELECT
        Nome_Projeto, VP, Formula_Fase, an, gp, LT, Lider_Teste, Gerente_Desenv, Formula_Status_Projeto, Descricao, Nome_Arquivo, 
        Ini_desenv, Term_desenv,  Completude1, Ini_Teste_Integrado, Term_Teste_Integrado, Completude2, 
        Ini_hml, Fim_hml, Completude3, Problema_Risco, SubCausa, Plano_Acao, causa 
        FROM projetos
        WHERE (Gerente_Desenv = ? OR Gerente_Desenv = ?)
        AND Formula_Status_Projeto not in (?, ?)
        """,(juliana,leonardo,requested,fora_pipeline,))


        wbTemplate = load_workbook('Templ1.xlsx')
        wbTempSource = wbTemplate.active

        thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))

        i = 0
        Linha = 2
        wsCopia = wbTemplate.copy_worksheet(wbTempSource)
        for Nome_Projeto, VP, Formula_Fase, an, gp, LT, Lider_Teste, Gerente_Desenv, Formula_Status_Projeto, Descricao, Nome_Arquivo, Ini_desenv, Term_desenv,  Completude1, Ini_Teste_Integrado, Term_Teste_Integrado, Completude2, Ini_hml, Fim_hml, Completude3, Problema_Risco, SubCausa, Plano_Acao, causa in cursor.fetchall():
                i = i + 1                
                CellProjeto = "A"+str(Linha)
                CellVP = "B"+str(Linha)
                CellFormula_Fase = "C"+str(Linha)
                Cellan = "D"+str(Linha)
                Cellgp = "E"+str(Linha)
                CellLT = "F"+str(Linha)
                CellLider_Teste = "G"+str(Linha)
                CellGerente_Desenv = "I"+str(Linha)
                CellFormula_Status_Projeto = "J"+str(Linha)

                wsCopia[CellProjeto] = Nome_Projeto
                wsCopia[CellVP] = VP
                wsCopia[CellFormula_Fase] = Formula_Fase
                
                wsCopia[CellProjeto].border = thin_border
                wsCopia[CellVP].border = thin_border
                Linha = Linha + 1
        wsCopia.title = "M"+str(i)
        wsCopia.sheet_view.showGridLines = False
        wbTemplate.save("ListaDeProjetos.xlsx")
        print ("Fim...")
                


        if conn:
                conn.close()


if __name__ == "__main__":
	main()
