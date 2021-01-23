import smtplib as sm
import os

while True:
    os.system('cls')
    print('1 - Ingreso')
    print('2 - Ver cantidad de personas')
    print('3 - Salir')
    n = input('>>>')
    if n == '1':
        os.system('cls')
        print('Opción 1')
        input('presione enter para continuar...')
    elif n == '2':
        os.system('cls')
        print('Opcion 2')
        input('presione enter para continuar...')
    elif n == '3':
        os.system('cls')
        print('Gracias por usar el programa')
        exit()
