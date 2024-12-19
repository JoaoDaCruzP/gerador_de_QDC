from leitor_arquivos import selecionar_celula, salva_arquivo
from funcoes import loading, msg_colorida

# VARIAVEIS QUE CONTROLAM AS LINHAS DA PLANILHA E O LAÇO DE REPETIÇÃO
lin = 1
iniciar = 's'
decisao = 's'
# APRSENTAÇÃO DO PROGRAMA
#print('BEM VINDO AO GERADOR DE QUADRO DE CARGAS')
msg_colorida(text="BEM VINDO, AO GERADOR DE QUADRO DE CARGAS.", bg="blue")
loading()

# LAÇO DE REPETIÇÃO QUE CHAMA A FUNÇÃO DE ADICIONA ITEM NA PLANILHA

while (iniciar == 's'):
    if iniciar == 's':
        lin = lin + 1

        selecionar_celula(lin)

        iniciar = input(f'\nADICIONAR MAIS ITENS?\n/ [S]SIM  [N]NAO: ').lower()
        if iniciar !=  's':
            decisao = input('DESEJA SALVAR O ARQUIVO?: ').lower()
            if decisao == 's':
                salvar = salva_arquivo()
        
            elif decisao != 's':   
                decisao = input('DESEJA MESMO SAIR SEM SALVAR O ARQUIVO?: SIM[S] NÃO[N]: ').lower()
                if decisao != 's':
                    salvar = salva_arquivo()
                else:
                    for _ in range(20):
                        print("=", end="", flush = True)
                    msg_colorida(text="\n ARQUIVO DESCARTADO. ", cor="red")
                    for _ in range(20):
                        print("=", end="", flush = True)
            else:
                msg_colorida(text=" ARQUIVO DESCARTADO. ", cor="red")
    else:
        break
