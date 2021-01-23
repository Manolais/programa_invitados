import smtplib as sm
import os

personas = []

txt = open('personas.txt','a')

while True:
    os.system('cls')
    print('1 - Ingreso')
    print('2 - Ver cantidad de personas')
    print('3 - Salir y enviar')
    n = input('>>>')
    if n == '1':
        while True:
            os.system('cls')
            per = input('Ingrese el nombre: ')
            tel = input('Ingrese número telefónico: ')
            try:
                tel = int(tel)
                break
            except ValueError:
                print('Ingrese un valor numérico')
                input('Presione enter para continuar...')
        mail = input('Ingrese el mail:')
                

    elif n == '2':
        os.system('cls')
        print('Opcion 2')
        input('presione enter para continuar...')
    elif n == '3':
        os.system('cls')
        print('Gracias por usar el programa')
        exit()
