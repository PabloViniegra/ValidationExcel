import os
import sys
import validations

#Cuando ejecutemos el script tiene que venir como argumento el nombre de la campaña para saber que validacion hacer. Con sys.argv recogemos una lista de los argumentos donde el [0] es el propio nombre del script
if len(sys.argv) == 1:
    path = '../FTP/'
    excelfiles = os.listdir(path)
    patternToExcel = '.+\.(xlsx|xls|xlsm)'

    validacion = validations.ValidacionesManuales(excelfiles, patternToExcel)
    check = validacion.ejecutarValidaciones()
    if check:
        print("La campaña ha pasado las validaciones correctamente")
        with open('../OK.txt', 'a') as f:
            f.write("Se han realizado las validaciones sin problema")
    else:
        print("No han pasado las validaciones. Revise el archivo validaciones.txt")

else:
    print("Los argumentos no son correctos")



