import re
import pandas as pd
from openpyxl import load_workbook
pathToFiles = '../FTP/'

class PremiamosTuConfianzaDentalPre():
    excelfiles = []
    patternToExcel = ""
    def __init__(self, excelfiles, patternToExcel):
        excelfiles = self.excelfiles
        patternToExcel = self.patternToExcel

    def executeValidations(self):
        checkNIFandCard = True
        for file in self.excelfiles:
            if re.match(self.patternToExcel, file, re.M | re.I):
                df = pd.read_excel(pathToFiles + file)
                line_number = 2
                for index, row in df.iterrows():
                    # print(row)
                    # print(index)
                    if pd.isna(row['IDIOMA']):
                        print("La fila " + str(line_number) + " no tiene idioma en " + file)
                        workbook = load_workbook(filename=pathToFiles + file)
                        sheet = workbook.active
                        sheet["C" + str(line_number)] = "CAS"
                        workbook.save(pathToFiles + file)
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

    def executeValidations(self):
        global checkIfCardExist
        for file in self.excelfiles:
            checkIfCardExist = True
            #Creamos la columna de TARGET para agregar los datos que se piden
            workbook = load_workbook(filename=pathToFiles + file)
            sheet = workbook.active
            sheet["Z1"] = 'TARGET'
            workbook.save(pathToFiles + file)
            if re.match(self.patternToExcel, file, re.M | re.I):
                df = pd.read_excel(pathToFiles + file)
                line_number = 2
                # validacion de si en target corresponde JOV o GRAL
                if '_JOV_' in file:
                    print("En el archivo " + file + " el target debe ser JOVEN")
                    targetToAdd = 'JOVEN'
                else:
                    targetToAdd = 'GRAL'
                    print("En el archivo " + file + " el target debe ser GRAL")
                for index, row in df.iterrows():
                    #validación del idioma
                    if pd.isna(row['IDIOMA']):
                        print("La fila " + str(line_number) + " no tiene idioma en " + file)
                        workbook = load_workbook(filename=pathToFiles + file)
                        sheet = workbook.active
                        sheet["C" + str(line_number)] = "CAS"
                        workbook.save('ExcelFiles/' + file)
                    #validación de si el numero de tarjeta está en blanco
                    if pd.isna(row['N_TARJETA']):
                        print("La fila " + str(line_number) + " no tiene número de tarjeta en " + file)
                        with open('../validations.txt', 'a') as f:
                            f.write("La fila " + str(line_number) + " no tiene número de tarjeta en " + file + "\n")
                        checkIfCardExist = False
                    #validación de si el codigo de promo esta en blanco o no es el especificado
                    if 'CADB1' != row['CODIGO_PROMO'] or pd.isna(row['CODIGO_PROMO']):
                        print("La fila " + str(line_number) + " no tiene el codigo de promoción correcto en " + file + ". Se establece CADB1 por defecto")
                        workbook = load_workbook(filename=pathToFiles + file)
                        sheet = workbook.active
                        sheet["Y" + str(line_number)] = 'CADB1'
                        workbook.save(pathToFiles + file)
                    #Escribimos en la columna TARGET su valor correspondiente
                    workbook = load_workbook(filename=pathToFiles + file)
                    sheet = workbook.active
                    sheet["Z" + str(line_number)] = targetToAdd
                    workbook.save(pathToFiles + file)
                    line_number += 1
            else:
                raise Exception(file + " is not an Excel File !")
        return checkIfCardExist

class VentajasPlena():
    excelfiles = []
    patternToExcel = ""

    def __init__(self, excelfiles, patternToExcel):
        excelfiles = self.excelfiles
        patternToExcel = self.patternToExcel

    def executeValidations(self):
        for file in self.excelfiles:
            checkIfCardExist = True
            targetToAdd = 'GRAL'
            workbook = load_workbook(filename=pathToFiles + file)
            sheet = workbook.active
            sheet["Z1"] = 'TARGET'
            workbook.save(pathToFiles + file)
            if re.match(self.patternToExcel, file, re.M | re.I):
                df = pd.read_excel(pathToFiles + file)
                line_number = 2
                for index, row in df.iterrows():
                    #validación del idioma
                    if pd.isna(row['IDIOMA']):
                        print("La fila " + str(line_number) + " no tiene idioma en " + file)
                        workbook = load_workbook(filename=pathToFiles + file)
                        sheet = workbook.active
                        sheet["C" + str(line_number)] = "CAS"
                        workbook.save(pathToFiles + file)
                    # validación de si el numero de tarjeta está en blanco
                    if pd.isna(row['N_TARJETA']):
                        print("La fila " + str(line_number) + " no tiene número de tarjeta en " + file)
                        with open('../validations.txt', 'a') as f:
                            f.write("La fila " + str(line_number) + " no tiene número de tarjeta en " + file + "\n")
                        checkIfCardExist = False
                        # validación de si el codigo de promo esta en blanco o no es el especificado
                    if 'CAD2' != row['CODIGO_PROMO'] or pd.isna(row['CODIGO_PROMO']):
                        print("La fila " + str(line_number) + " no tiene el codigo de promoción correcto en " + file + ". Se establece CAD2 por defecto")
                        workbook = load_workbook(filename=pathToFiles + file)
                        sheet = workbook.active
                        sheet["Y" + str(line_number)] = 'CAD2'
                        workbook.save(pathToFiles + file)
                    #Escribimos en la columna TARGET el valor especificado
                    workbook = load_workbook(filename=pathToFiles + file)
                    sheet = workbook.active
                    sheet["Z" + str(line_number)] = targetToAdd
                    workbook.save(pathToFiles + file)
                    line_number += 1
            else:
                raise Exception(file + " is not an Excel File !")