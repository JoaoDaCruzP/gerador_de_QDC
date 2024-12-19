import time
import _tkinter as tk

# FUNÇÃO DE ESPERA
def loading():
    for _ in range(15):
        print('=', end='', flush=True)
        time.sleep(.1)
    print('\n')

def nome_item():
    name = ''
    potencia = ''
    fp = ''
    fd = ''
    
    print(' LISTA DE CARGAS   \n')
    escolha = str(input('[1] SPLIT\n[2] ILUMINAÇÃO\n[3]TOMADA\n[4] CHUVEIRO ELETRICO\n[5] MOTOR DE PORTÃO\n\n DIGITE O NUMERO DO ITEM:  '))
    if escolha == str('1'):
        escolha = str(input('LISTA DE CARGAS\n\n[1] SPLIT 9\n[2] SPLIT 12\n[3] SPLIT 18\n\nDIGITE A POTENCIA DO SPLIT: '))
        if escolha == str('1'):
            name = 'SPLIT 9.000 BTUS'
            potencia = float(0.99)
            fp = float(0.85)
            fd = float(0.92)
        elif escolha == str('2'):
            name = 'SPLIT 12.000 BTUS'
            potencia = float(1.26)
            fp = float(0.85)
            fd = float(0.92)

        elif escolha == str('3'):
            name = 'SPLIT 18.000 BTUS'
            potencia = float(1.98)
            fp = float(0.85)
            fd = float(0.92)
        else:
            print('Não foi possivel adicinar o valor')
    elif escolha == str('2'):
        name = 'ILUMINAÇÃO'
        potencia = float(0.030)
        fp = float(0.90)
        fd = float(0.45)

    elif escolha == str('3'):
        name = 'TOMADA'
        potencia = float(0.1)
        fp = float(0.85)
        fd = float(0.45)

    elif escolha == str('4'):
        name = 'CHUVEIRO ELETRICO'
        potencia = float(4.5)
        fp = float(1)
        fd = float(0.40)
    
    elif escolha == str('5'):
        name = 'MOTOR PORTÃO'
        potencia = float(0.4)
        fp = float(0.70)
        fd = float(0.40)

    return name, potencia, fd, fp

    

def quant_item():
    quant = int(input('DIGITE A QUANTIDADE DESEJADA: '))
    return quant

def msg_colorida(cor='', text='', bg=''):
    blue = '\033[1;34m'
    red = '\033[1;31m'
    green = '\033[1;32m'
    bg_azul = '\033[44m'
    bg_vermelho = '\033[41m'
    bg_green = '\033[42m'
    reset = '\033[m'

    if cor == 'blue':
        print(blue,text,reset)
    elif cor == 'red':
        print(red,text,reset)
    elif cor == 'green':
        print(green,text,reset)
    else:
        if bg == 'blue':
            print(bg_azul,text,reset)
        elif bg == 'red':
            print(bg_vermelho,text,reset)
        elif bg == 'green':
            print(bg_green,text,reset)
        
            
'''msg_colorida(text='MEU AMIGO, ISSO É PROGRAMAÇÃO DE PONTA',cor='blue')
msg_colorida(text='OLHA SÓ ESSA FUNÇÃO QUE EU CRIEI',cor='red')
msg_colorida(text='DO ZERO MAN, CE TA LOUCO!!!',cor='green')
msg_colorida(text='TU SO..... ', bg='blue')
msg_colorida(text='PODE ESTAR DE ....', bg='red')
msg_colorida(text='SACANAGEM!!!!KKKKKKKKK',bg='green')'''