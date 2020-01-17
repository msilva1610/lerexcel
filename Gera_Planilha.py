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


#Cell_NomeDoProjeto = A2
#Cell_VP = D2
#Cell_Formula_Fase = E2
#Cell_AN = F2
#Cell_GP = G2
#Cell_LT = H2
#Cell_Lider_Teste = I2
#Cell_Gerente_Desenvolvimento = J2
#Cell_Formula_Status_Projeto = K2
#Cell_Descricao = A4
#Cell_Nome_Arquivo = A9
#Cell_Inicio_Desenv = C9
#Cell_Termino_Desenv = D9
#Cell_Complitude1 = E9
#Cell_Inicio_Teste_Integrado = F9
#Cell_Termino_Teste_Integrado = G9
#Cell_Complitude2 = H9
#Cell_Inicio_HML = I9
#Cell_Termino_HML = J9
#Cell_Complitude3 = K9
#Cell_Problemas_Riscos = A12
#Cell_SubCausa = E12
#Cell_Plano_de_Acao = G12

def main():
    ler_ultimoArquivo()

def ler_ultimoArquivo():
    print("lendo...")

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(BASE_DIR, "Projetos.db")

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
    SELECT max(id) FROM arquivos
    """)

    total_fetchone = cursor.fetchone()[0]
    print(total_fetchone)

    if conn:
        conn.close()

    ler_projeto_id(total_fetchone)

def ler_projeto_id(id_arquivo):
    Cell_NomeDoProjeto = "A2"
    Cell_VP = "D2"
    Cell_Formula_Fase = "E2"
    Cell_AN = "F2"
    Cell_GP = "G2"
    Cell_LT = "H2"
    Cell_Lider_Teste = "I2"
    Cell_Gerente_Desenvolvimento = "J2"
    Cell_Formula_Status_Projeto = "K2"
    Cell_Descricao = "A4"
    Cell_Nome_Arquivo = "A9"
    Cell_Inicio_Desenv = "C9"
    Cell_Termino_Desenv = "D9"
    Cell_Complitude1 = "E9"
    Cell_Inicio_Teste_Integrado = "F9"
    Cell_Termino_Teste_Integrado = "G9"
    Cell_Complitude2 = "H9"
    Cell_Inicio_HML = "I9"
    Cell_Termino_HML = "J9"
    Cell_Complitude3 = "K9"
    Cell_Problemas_Riscos = "A12"
    Cell_SubCausa = "E12"
    Cell_Plano_de_Acao = "G12"

    
    requested = "Requested"
    at_risk = "At Risk"
    delayed = "Delayed"
    fora_pipeline = "fora pipeline"
    juliana = "Juliana Alves Castro Perez"
    leonardo = "Leonardo Augusto Mendes Leandro"
    thales = "Thales Antonio Silva De Freitas"
    fase1 = "Fase 1"
    fase2 = "Fase 2"
    fase3 = "Fase 3" 
    fase4 = "Fase 4"
    parado = "Parado"
    formulafase = "X"
    LinhaInicial = 14

    print("Lendo projeto ...")
   
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
    WHERE Gerente_Desenv in (?, ?, ?)
    AND Formula_Status_Projeto not in (?, ?)
    """,(juliana,leonardo,thales,requested,fora_pipeline,))

    i = 0

    wbTemplate = load_workbook('Template.xlsx')
    wbTempSource = wbTemplate.active

    thin_border = Border(left=Side(style='thin'), 
                 right=Side(style='thin'), 
                 top=Side(style='thin'), 
                 bottom=Side(style='thin'))

    for Nome_Projeto, VP, Formula_Fase, an, gp, LT, Lider_Teste, Gerente_Desenv, Formula_Status_Projeto, Descricao, Nome_Arquivo, Ini_desenv, Term_desenv,  Completude1, Ini_Teste_Integrado, Term_Teste_Integrado, Completude2, Ini_hml, Fim_hml, Completude3, Problema_Risco, SubCausa, Plano_Acao, causa in cursor.fetchall():
        i = i + 1
        wsTemplate = wbTemplate['Template']
        #print (wbTemplate.get_sheet_names())
        wsCopia = wbTemplate.copy_worksheet(wbTempSource)

        if (str(Formula_Fase).strip() == fase1):
            formulafase = "F1"
        elif (str(Formula_Fase).strip() == fase2):
            formulafase = "F2"
        elif (str(Formula_Fase).strip() == fase3):
            formulafase = "F3"
        elif (str(Formula_Fase).strip() == fase4):
            formulafase = "F4"
        else:
            formulafase = "XX"

        if (Gerente_Desenv == leonardo):
            wsCopia.title = formulafase + "L"+str(i)
        elif (Gerente_Desenv == thales):
            wsCopia.title = formulafase + "T"+str(i)
        else:
            wsCopia.title = formulafase + "J"+str(i)

        
        if (Formula_Status_Projeto == at_risk):
            wsCopia.sheet_properties.tabColor = "FFFF00"
        elif (Formula_Status_Projeto == delayed):
            wsCopia.sheet_properties.tabColor = "FF0000"

        Prob_risco_upper = str(Problema_Risco).upper()
        intPalavraAmbiente = Prob_risco_upper.find("AMBIENTE")
        if (intPalavraAmbiente > 0):
            #pinta de roxo
            wsCopia.sheet_properties.tabColor = "660066"
            print (i,Nome_Projeto) 

        wsCopia[Cell_NomeDoProjeto] = Nome_Projeto
        wsCopia[Cell_VP] = VP
        wsCopia[Cell_Formula_Fase] = Formula_Fase
        wsCopia[Cell_AN] = an
        wsCopia[Cell_GP] = gp
        wsCopia[Cell_LT] = LT
        wsCopia[Cell_Lider_Teste] = Lider_Teste
        wsCopia[Cell_Gerente_Desenvolvimento] = Gerente_Desenv
        wsCopia[Cell_Formula_Status_Projeto] = Formula_Status_Projeto
        wsCopia[Cell_Descricao] = Descricao
        wsCopia[Cell_Nome_Arquivo] = Nome_Arquivo
        wsCopia[Cell_Inicio_Desenv] = Ini_desenv
        wsCopia[Cell_Termino_Desenv] = Term_desenv
        wsCopia[Cell_Complitude1] = Completude1
        wsCopia[Cell_Inicio_Teste_Integrado] = Ini_Teste_Integrado
        wsCopia[Cell_Termino_Teste_Integrado] = Term_Teste_Integrado
        wsCopia[Cell_Complitude2] = Completude2
        wsCopia[Cell_Inicio_HML] = Ini_hml
        wsCopia[Cell_Termino_HML] = Fim_hml
        wsCopia[Cell_Complitude3] = Completude3
        wsCopia[Cell_Problemas_Riscos] = Problema_Risco
        wsCopia[Cell_SubCausa] = SubCausa
        wsCopia[Cell_Plano_de_Acao] = Plano_Acao

        ListaDeRmsParaOProjeto = RetornaRms(Nome_Projeto)
        LinhaInicial = 14
        for row in ListaDeRmsParaOProjeto:
            CellRM = "A"+str(LinhaInicial)
            CellResumo = "B"+str(LinhaInicial)
            CellResp = "C"+str(LinhaInicial)
            CellDataCriacao = "D"+str(LinhaInicial)
            CellDataComite = "E"+str(LinhaInicial)
            CellSistema = "F"+str(LinhaInicial)
            CellDataInicioTI = "G"+str(LinhaInicial)
            CellDataFimTI = "H"+str(LinhaInicial)
            CellDataIniHML = "I"+str(LinhaInicial)
            CellDataFimHML = "J"+str(LinhaInicial)
            CellStatus = "K"+str(LinhaInicial)

            wsCopia[CellRM] = row[0]
            wsCopia[CellResumo] = row[1]
            wsCopia[CellResp] = row[2]
            wsCopia[CellDataCriacao] = str(row[3])[0:10]
            wsCopia[CellDataComite] = str(row[4])[0:10]
            wsCopia[CellSistema] = row[5]
            wsCopia[CellDataInicioTI] = str(row[6])[0:10]
            wsCopia[CellDataFimTI] = str(row[7])[0:10]
            wsCopia[CellDataIniHML] = str(row[8])[0:10]
            wsCopia[CellDataFimHML] = str(row[9])[0:10]
            wsCopia[CellStatus] =row[10]

            wsCopia[CellRM].border = thin_border
            LinhaInicial = LinhaInicial + 1

    wsCopia.sheet_view.showGridLines = False
        #print (wbTemplate.get_sheet_names())
    wbTemplate.save("ProjetosLidos10.xlsx")
    #print (wbTemplate.get_sheet_names())

    if conn:
        conn.close()
    
def RetornaRms(Nome_Projeto):
    try:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        db_path = os.path.join(BASE_DIR, "Projetos.db")
            
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        Projeto = str(Nome_Projeto[0:7]).strip()

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

        #print ("Total de Rms encontradas: {}".format(str(len(ListaDeRms))))
    
        return ListaDeRms

        conn.commit()
    except Exception as e:
        raise
    finally:
        conn.close()
        
if __name__ == "__main__":
    main()
