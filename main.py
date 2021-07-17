import os, sys
import validations
#Cuando ejecutemos el script tiene que venir como argumento el nombre de la campaña para saber que validacion hacer. Con sys.argv recogemos una lista de los argumentos donde el [0] es el propio nombre del script
if len(sys.argv) == 2:
    path = '../FTP/'
    excelfiles = os.listdir(path)
    patternToExcel = '.+\.(xlsx|xls|xlsm)'
    if 'Premiamos_Tu_Confianza_Dental_Pre' == sys.argv[1]:
        #Instanciamos el parser1 y ejecutamos su validación
        parser1 = validations.PremiamosTuConfianzaDentalPre(excelfiles, patternToExcel)
        check = parser1.executeValidations()
        if check:
            print(sys.argv[1] + " ha pasado las validaciones correctamente")
            open('../OK.txt', 'w')
        else:
            print("No han pasado las validaciones. Revise el archivo validaciones.txt")
    elif 'Ventajas_Basico' == sys.argv[1]:
        #Instancia de parser de la campaña 2
        parser2 = validations.VentajasBasico(excelfiles, patternToExcel)
        check = parser2.executeValidations()
        if check:
            print(sys.argv[1] + " ha pasado las validaciones correctamente")
            open('../OK.txt', 'w')
        else:
            print("No han pasado las validaciones. Revise el archivo validaciones.txt")
    elif 'Ventajas_Plena' == sys.argv[1]:
        #Instancia de parser de la campaña 3
        parser3 = validations.VentajasPlena(excelfiles, patternToExcel)
        check = parser3.executeValidations()
        if check:
            print(sys.argv[1] + " ha pasado las validaciones correctamente")
            open('../OK.txt', 'w')
        else:
            print("No han pasado las validaciones. Revise el archivo validaciones.txt")
    elif 'Revision_Medica' == sys.argv[1]:
        #Instancia de parser de la campaña 4
        parser4 = validations.RevisionMedica(excelfiles, patternToExcel)
        check = parser4.executeValidations()
        if check:
            print(sys.argv[1] + " ha pasado las validaciones correctamente")
            open('../OK.txt', 'w')
        else:
            print("No han pasado las validaciones. Revise el archivo validaciones.txt")
    elif 'SAC_Adeslas_Basico' == sys.argv[1] or 'SAC_Adeslas_Plena' == sys.argv[1]:
        #Instancia de parser de la campaña 5 y 6
        parser5_6 = validations.SACAdeslasBasicaYPlena(excelfiles, patternToExcel)
        check = parser5_6.executeValidations()
        if check:
            print(sys.argv[1] + " ha pasado las validaciones correctamente")
            open('../OK.txt', 'w')
        else:
            print("No han pasado las validaciones. Revise el archivo validaciones.txt")
    elif 'Caixa_Accidentes' == sys.argv[1]:
        #Instancia de parser de la campaña 7
        parser7 = validations.CaixaAccidentes(excelfiles, patternToExcel)
        check = parser7.executeValidations()
        if check:
            print(sys.argv[1] + " ha pasado las validaciones correctamente")
            open('../OK.txt', 'w')
        else:
            print("No han pasado las validaciones. Revise el archivo validaciones.txt")
    elif 'Premiamos_Tu_Confianza_Dental_Post' == sys.argv[1]:
        #Instancia de parser de la campaña 8
        parser8 = validations.PremiamosTuConfianzaDentalPost(excelfiles, patternToExcel)
        check = parser8.executeValidations()
        if check:
            print(sys.argv[1] + " ha pasado las validaciones correctamente")
            open('../OK.txt', 'w')
        else:
            print("No han pasado las validaciones. Revise el archivo validaciones.txt")
    else:
        print("No se reconoce ese nombre de Campaña")
else:
    print("Los argumentos no son correctos")



