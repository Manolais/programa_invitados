import smtplib as sm
import os
from openpyxl import workbook,load_workbook

personas = []

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
        while True:
            os.system('cls')
            print('1 - Ver invitados en ram')
            print('2 - Ver invitados en excel')
            p = input('>>>')
            if p == '1':
                print('--------------------')
                for i in range(len(personas)):
                    print('')
                    print('Invitado',i+1)
                    print()
                    print('Nombre:',personas[i]['Nombre'])
                    print('Teléfono:',personas[i]['Teléfono'])
                    print('Email:',personas[i]['Email'])
                    print('--------------------')
                input('Presione enter para continuar')
            if p == '2':
                wb = load_workbook('personas.xlsx')
                sheet1 = wb.get_sheet_by_name('Sheet1')
                if type(sheet1['A2'].value) == None:
                    os.system('cls')
                    print('Excel vacio, vuelva cuando guarde los cambios')
                    input('Presione enter para continuar...')
                elif type(sheet1['A2'].value) == str:
                    os.system('cls')
                    print('--------------------')
                    for i in range(len(sheet1['A'])-1):
                        print('Invitado', i+1)
                        print()
                        print(sheet1['A'+str(i+2)].value)
                        print(sheet1['B'+str(i+2)].value)
                        print(sheet1['C'+str(i+2)].value)
                        print('--------------------')
                    input('Presione enter para continuar...')
            break


    elif n == '3':
        wb = load_workbook('personas.xlsx')
        sheet = wb.get_sheet_by_name('Sheet1')
        for i in range(len(personas)):
            sheet['A'+str(len(sheet['A'])+1)] = personas[i]['Nombre']
            sheet['B'+str(len(sheet['B']))] = personas[i]['Teléfono']
            sheet['C'+str(len(sheet['C']))] = personas[i]['Email']
        wb.save('personas.xlsx')
        sv = sm.SMTP('smtp.gmail.com',587)
        sv.starttls()
        sv.login('programainvitados@gmail.com','prog_invit')
        sv.sendmail('programainvitados@gmail.com','bsasmanuelsilva@gmail.com',str(personas))
        os.system('cls')
        print('Gracias por usar el programa, guardando invitados...')
        print('Correo enviado correctamente!')
        exit()
    else:
        print('Ingrese una opción valida')
        input('Presione enter para continuar...')