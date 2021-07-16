import re
import pandas as pd
from openpyxl import load_workbook

class PremiamosTuConfianzaDentalPre():
    excelfiles = []
    patternToExcel = ""
    def __init__(self, excelfiles, patternToExcel):
        excelfiles = self.excelfiles
        patternToExcel = self.patternToExcel

    def validationCommon(self):
        checkNIFandCard = True
        for file in self.excelfiles:
            if re.match(self.patternToExcel, file, re.M | re.I):
                df = pd.read_excel('ExcelFiles/' + file)
                line_number = 2
                for index, row in df.iterrows():
                    # print(row)
                    # print(index)
                    if pd.isna(row['IDIOMA']):
                        print("La fila " + str(line_number) + " no tiene idioma en " + file)
                        workbook = load_workbook(filename='ExcelFiles/' + file)
                        sheet = workbook.active
                        sheet["C" + str(line_number)] = "CAS"
                        workbook.save('ExcelFiles/' + file)
                    if pd.isna(row['N_TARJETA']) and pd.isna(row['NIF']):
                        print("La fila " + str(line_number) + " no tiene número de tarjeta ni NIF en " + file)
                        with open('../validations.txt', 'a') as f:
                            f.write("La fila " + str(line_number) + " no tiene número de tarjeta ni NIF en " + file + "\n")
                        checkNIFandCard = False

                    line_number += 1
            else:
                raise Exception(file + " is not an Excel File !")
        return checkNIFandCard


class VentajasBasico():
    excelfiles = []
    patternToExcel = ""

    def __init__(self, excelfiles, patternToExcel):
        excelfiles = self.excelfiles
        patternToExcel = self.patternToExcel
    #TODO: validaciones de la Campaña 2
    def validationCommon(self):
        pass
