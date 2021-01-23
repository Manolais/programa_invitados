import smtplib as sm
import os
from openpyxl import workbook,load_workbook

personas = []

txt = open('personas.txt','a')

while True:
    os.system('cls')
    print('1 - Ingreso')
    print('2 - Ver cantidad de personas')
    print('3 - Salir y enviar')
    n = input('>>> ')
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
        mail = input('Ingrese el mail: ')
        personas.append({'Nombre': per,'Teléfono': tel,'Email': mail})
        print('¡Invitado ingresado correctamente!')
        input('Presione enter para continuar...')

    elif n == '2':
        os.system('cls')
        for i in range(len(personas)):
            print('Invitado',i+1)
            print()
            print('Nombre:',personas[i]['Nombre'])
            print('Teléfono:',personas[i]['Teléfono'])
            print('Email:',personas[i]['Email'])
            print()
        input('Presione enter para continuar')
    elif n == '3':
        os.system('cls')
        print('Gracias por usar el programa, guardando invitados...')
        wb = load_workbook('personas.xlsx')
        sheet = wb.get_sheet_by_name('Sheet1')
        for i in range(len(personas)):
            sheet['A'+str(len(sheet['A'])+1)] = personas[i]['Nombre']
            sheet['B'+str(len(sheet['B']))] = personas[i]['Teléfono']
            sheet['C'+str(len(sheet['C']))] = personas[i]['Email']
        wb.save('personas.xlsx')
        
        exit()
