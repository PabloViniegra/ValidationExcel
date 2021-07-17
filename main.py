import os, sys
import validations
path = '../FTP'
excelfiles = os.listdir(path)
patternToExcel = '.+\.(xlsx|xls|xlsm)'
#Cuando ejecutemos el script tiene que venir como argumento el nombre de la campaña para saber que validacion hacer. Con sys.argv recogemos una lista de los argumentos donde el [0] es el propio nombre del script
if sys.argv == 2:

    if 'Premiamos_Tu_Confianza_Dental_Pre' == sys.argv[1]:
        #Instanciamos el parser1 y ejecutamos su validación
        parser1 = validations.PremiamosTuConfianzaDentalPre(excelfiles, patternToExcel)
        check = parser1.executeValidations()
        if check:
            print("Las validaciones han sido correctas")
            open('../OK.txt', 'w')
        else:
            print("No han pasado las validaciones")
    elif 'Ventajas_Basico' == sys.argv[1]:
        #Instancia de parser de la campaña 2
        parser2 = validations.VentajasBasico(excelfiles, patternToExcel)
        check = parser2.executeValidations()
        if check:
            print("Las validaciones han sido correctas")
            open('../OK.txt', 'w')
        else:
            print("No han pasado las validaciones")
    elif 'Ventajas_Plena' == sys.argv[1]:
        #Instancia de parser de la campaña 3
        pass
else:
    print("Los argumentos no son correctos")


#TODO: desarrollar y probar mas parsers
