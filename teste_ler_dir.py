# -*- coding: utf-8 -*
import datetime
import sys
import glob, os
from openpyxl import load_workbook

def main():
    #lerdir01()
    lerdir03()

def lerdir01():
    os.chdir("pendentes")
    for file in glob.glob("*.xlsm"):
        print(file)

def lerdir03():
    os.chdir("pendentes")
    files = glob.glob("*.xlsm")
    #files.sort(key=os.path.getmtime)
    files.sort(key=os.path.getctime)
    print("\n".join(files))

def lerdir04():
    for root, dirs, files in os.walk("pendentes"):
        for file in files:
            if file.endswith(".xlsm"):
                 print(os.path.join(root, file))

if __name__ == "__main__":
	main()

