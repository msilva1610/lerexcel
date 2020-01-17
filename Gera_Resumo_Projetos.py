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
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

        db_path = os.path.join(BASE_DIR, "Projetos.db")

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute("""
                select id, nome_projeto, substr(nome_projeto,1,7), Formula_Fase, Formula_Status_Projeto 
                from projetos as p
                where formula_fase in ("Cancelado", "Fase 1", "Fase 2", "Fase 3", "Fase 4", "Parado", "Sem Fase")
                and id >= ?
                order by 1
        """,("1",))


        wbTemplate = load_workbook('TemplateResumo.xlsx')
        wbTempSource = wbTemplate.active

        i = 0
        Linha = 2
        wsCopia = wbTemplate.copy_worksheet(wbTempSource)

        cursor =  cursor.fetchall()
        
        monta_planilha(wsCopia, cursor)

        wsCopia.title = "M"
        wsCopia.sheet_view.showGridLines = False
        wbTemplate.save("ListaResumoProjetos.xlsx")
        print ("Fim...")

        if conn:
                conn.close()

def monta_planilha(wsCopia, cs):
                
        for id, nome_projeto, subs_nome_projeto, Formula_Fase, Formula_Status_Projeto  in cs.fetchall():
                i = i + 1                
                CellProjeto = "A"+str(Linha)
                CellQTDE_DT_ID = "B"+str(Linha)
                CellQTDE_DT_TD = "C"+str(Linha)
                CellQTDE_DT_ITI	= "D"+str(Linha)
                CellQTDE_DT_TTI	= "E"+str(Linha)
                CellQTDE_DT_IHML = "F"+str(Linha)
                CellQTDE_DT_FHML = "G"+str(Linha)
                CellFase = "H"+str(Linha)
                CellFormula = "I"+str(Linha)
                CellStatusProjeto = "J"+str(Linha)
                CellInicio_Desenvolvimento = "K"+str(Linha)
                CelFimDesenvolvimento = "L"+str(Linha)
                CellInicio_Teste_Integrado = "M"+str(Linha)
                CellTermino_Teste_Integrado = "N"+str(Linha)
                CellInicio_Hoologação = "O"+str(Linha)
                CellFim_Homologação = "P"+str(Linha)
                CellQtde_Sistemas = "Q"+str(Linha)
                CellQtde_Rms = "R"+str(Linha)
                Cell_Gerente =  "S"+str(Linha)

                id_projeto = coluna[1]
                nome_projeto = coluna [2]
                nome_abreviado = colune[3]
                Formula_Fase = colune[4]
                Formula_Status_Projeto = coluna[5]

                wsCopia[CellProjeto]= nome_projeto
                wsCopia[CellFase] = Formula_Fase
                wsCopia[CellFormula] = Formula_Status_Projeto

                #Analisa os Eventos de cada projeto
                ListaEventos = lista_comportamento_projeto(id_projeto)
                for cadaregistro in ListaEventos:
                        if (cadaregistro[1] == 400):
                                wsCopia[CellQTDE_DT_ID] = cadaregistro[1]
               
                Linha = Linha + 1

def lista_comportamento_projeto(id):
        try:
                BASE_DIR = os.path.dirname(os.path.abspath(__file__))

                db_path = os.path.join(BASE_DIR, "Projetos.db")

                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()

                cursor.execute("""
                        select cod_evento,evento, count(*) as qtde_DT_ID
                        from historico
                        where id_projeto = ?
                        group by 1,2
                """,(id,))
                
                return cursor
        
                if conn:
                        conn.close()
        except Exception as e:
                raise
        finally:
                if conn:
                        conn.close()
if __name__ == "__main__":
	main()
