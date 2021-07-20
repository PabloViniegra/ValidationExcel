import re
import datetime
import pandas as pd
from openpyxl import load_workbook

pathToFiles = '../FTP/'


class ValidacionesManuales:
    excelfiles = []
    patternToExcel = ""

    def __init__(self, excelfiles, patternToExcel):
        self.excelfiles = excelfiles
        self.patternToExcel = patternToExcel

    def ejecutarValidaciones(self):
        checkToReturn = True
        for file in self.excelfiles:
            if re.match(self.patternToExcel, file, re.M | re.I):
                print("Analizando " + file + " ....")
                df = pd.read_excel(pathToFiles + file)
                line_number = 2
                for index, row in df.iterrows():
                    # validación de idioma
                    if pd.isna(row['IDIOMA']):
                        print("La fila " + str(
                            line_number) + " no tiene idioma en " + file + " .Se establece por defecto en CAS")
                        workbook = load_workbook(filename=pathToFiles + file)
                        sheet = workbook.active
                        sheet["C" + str(line_number)] = "CAS"
                        workbook.save(pathToFiles + file)
                        with open('../OK.txt', 'a') as f:
                            f.write("La fila " + str(
                                line_number) + " no tiene idioma en " + file + " .Se establece por defecto en CAS" + "\n")
                    # validación si el numero de tarjeta o el nif vienen vacíos
                    if pd.isna(row['N_TARJETA']) and pd.isna(row['NIF']):
                        print("La fila " + str(line_number) + " no tiene número de tarjeta ni NIF en " + file)
                        with open('../validations.txt', 'a') as f:
                            f.write(str(datetime.datetime.now()) + ": La fila " + str(
                                line_number) + " no tiene número de tarjeta ni NIF en " + file + "\n")
                        checkToReturn = False
                    # validación de la vía de impacto
                    if pd.isna(row['CORREO_CLIENTE']):
                        print("Fila " + str(line_number) + " no hay correo. Se revisan los telefonos")
                        if pd.isna(row['TELEFONO_MOVIL_SMS']) and pd.isna(row['TELEFONO_MOVIL_SMS_2']):
                            print("La fila " + str(line_number) + " no tiene ni correo ni teléfono para ese cliente")
                            with open('../validations.txt', 'a') as f:
                                f.write(str(datetime.datetime.now()) + ": La fila " + str(
                                    line_number) + " no tiene ni correo ni teléfono para ese cliente en " + file + "\n")
                            checkToReturn = False
                        else:
                            if re.fullmatch("(666|696)+[ \d]?", str(row['TELEFONO_MOVIL_SMS']), re.M) or re.fullmatch(
                                    "(666|696)+[ \d]?", str(row['TELEFONO_MOVIL_SMS_2']), re.M):
                                print(
                                    "La fila " + str(line_number) + " no tiene números de teléfono válidos")
                                with open('../validations.txt', 'a') as f:
                                    f.write(str(datetime.datetime.now()) + ": En la fila " + str(
                                        line_number) + " los números de teléfono no son válidos en " + file + "\n")
                                checkToReturn = False
                            else:
                                print("Fila " + str(line_number) + " vía de impacto SMS")
                    else:
                        print("Fila " + str(line_number) + " vía de impacto CORREO")
                    line_number += 1
            else:
                raise Exception(file + " is not an Excel File !")
        return checkToReturn

