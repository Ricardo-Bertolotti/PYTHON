# IMPORTAÇÕES ----------------------------------------------------------------------------------------------------------

import os
import sys
import socket
import tkinter as tk

# VARIÁVEIS FIXAS ------------------------------------------------------------------------------------------------------

preto = str("#000000") # VARIÁVEL PARA A COR PRETO
preto_claro = str("#171717") # VARIÁVEL PARA A COR PRETO CLARO
cinza = str("#999999") # VARIÁVEL PARA A COR CINZA

Arquivo_caminho_pdf = r"C:\Users\r1c4r\OneDrive\Área de Trabalho\PROJ_PY\ARQUIVOS\AC ENG.pdf" # CAMINHO DO ARQUIVO .PDF
Icone_caminho = r"C:\Users\r1c4r\OneDrive\Área de Trabalho\PROJ_PY\ARQUIVOS\Impressora_img.ico" # CAMINHO DO ÍCONE
Printer_IP = str("") # VARIÁVEL PARA POSTERIORMENTE DEFINIR O VALOR DO ENDEREÇO IP

# FUNÇÕES --------------------------------------------------------------------------------------------------------------

def Titulo_janela(Janela_0, titulo): # FUNÇÃO PARA DEFINIR O TÍTULO DA JANELA
    Janela_0.title(titulo)

def Cor_janela(Janela_0, cor): # FUNÇÃO PARA DEFINIR A COR DE FUNDO DA JANELA
    Janela_0.configure(bg=cor)

def Dimensionar_janela(Janela_0, largura, altura): # FUNÇÃO PARA DEFINIR AS DIMENSÕES DA JANELA
    Janela_0.geometry(f"{largura}x{altura}")

def Centralizar_janela(Janela_0): # FUNÇÃO PARA CENTRALIZAR A JANELA NA TELA
    Janela_0.update_idletasks()
    X = (Janela_0.winfo_screenwidth() - Janela_0.winfo_width()) // 2
    Y = (Janela_0.winfo_screenheight() - Janela_0.winfo_height()) // 2
    Janela_0.geometry(f"{Janela_0.winfo_width()}x{Janela_0.winfo_height()}+{X}+{Y}")

def Definir_Printer_IP(): # FUNÇÃO PARA DEFINIR O ENDEREÇO IP MEDIANTE A ESCOLHA DA IMPRESSORA

    global Printer_IP

    IP_local = Escolha_0.get()

    if IP_local == "ADM":
        Printer_IP = str("192.168.107.128")
    elif IP_local == "AK":
        Printer_IP = str("192.168.107.104")
    elif IP_local == "MOBILE":
        Printer_IP = str("192.168.107.119")
    elif IP_local == "NOTEBOOK":
        Printer_IP = str("192.168.107.131")

def Click_Button_0(event): # FUNÇÃO PARA ALTERAR A COR AO SIMULAR/CLICAR NO Button_0
    if event.type == "7":
        Button_0.config(bg="white")
    elif event.type == "8":
        Button_0.config(bg=cinza)

def Imprimir(): # FUNÇÃO PARA IMPRIMIR O ARQUIVO NA IMPRESSORA ESCOLHIDA POR CONEXÃO SOCKET

    Definir_Printer_IP()

    Printer_port = 9100

    # ABRIR O ARQUIVO E FAZER A SUA LEITURA EM BINÁRIO
    with open(f'{Arquivo_caminho_pdf}', 'rb') as file:
        Data = file.read()

    # CRIAR UM SOCKET DO TIPO TCP/IP PARA SE COMUNICAR COM A IMPRESSORA
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.connect((Printer_IP, Printer_port))

    # ESTRUTURA DE REPETIÇÃO PARA ENVIAR O ARQUIVO DE ACORDO COM A QUANTIDADE DE CÓPIAS DO ARQUIVO
    for i in range(int(Spinbox_0.get())):

        # ENVIAR O CONTEÚDO DO ARQUIVO PARA A IMPRESSORA
        sock.sendall(Data)

# INICIO DO CÓDIGO -----------------------------------------------------------------------------------------------------

Janela_0 = tk.Tk() # INICIA A JANELA

# OBTEM O CAMINHO PARA O DIRETÓRIO TEMPORÁRIO DO PYINSTALLER
if getattr(sys, 'frozen', False):
    # EXECUTANDO NO PYINSTALLER
    dir_path = sys._MEIPASS
else:
    # EXECUTANDO A PARTIR DO CÓDIGO-FONTE
    dir_path = os.path.dirname(os.path.abspath(__file__))

# CONCATENA O CAMINHO PARA O ARQUIVO
arq = os.path.join(dir_path, "AC ENG.pdf")

# CONCATENA O CAMINHO PARA O ÍCONE
Icone = os.path.join(dir_path, "Impressora_img.ico")

# FUNÇÃO PARA DEFINIR O NOME DA JANELA
Titulo_janela(Janela_0, "IMPRESSÃO - A/C ENGENHARIA")

# FUNÇÃO PRA DEFINIR A COR DE FUNDO DA JANELA
Cor_janela(Janela_0, preto)

# FUNÇÃO PRA DEFINIR O ÍCONE DA JANELA
#Janela_0.iconbitmap(Icone_caminho)

# FUNÇÃO PARA DEFINIR AS DIMENSÕES DA JANELA
Dimensionar_janela(Janela_0, 380, 400)

Centralizar_janela(Janela_0) # ATIVA A FUNÇÃO DE CENTRALIZAR A JANELA NA TELA DO USUÁRIO

Janela_0.resizable(False, False)  # DESABILITA A CAPACIDADE DE REDIMENSIONAR A JANELA

# FRAMES
Frame_0 = tk.Frame(Janela_0, bg=cinza, height=50)
Frame_0.place(x=20,y=20, relwidth=0.88, relheight=0.28)

Frame_1 = tk.Frame(Janela_0, bg=cinza, height=50)
Frame_1.place(x=20,y=150, relwidth=0.88, relheight=0.28)

Frame_2 = tk.Frame(Janela_0, bg=cinza, height=50)
Frame_2.place(x=20,y=305, relwidth=0.88, relheight=0.15)

# LABELS

Label_0 = tk.Label(Frame_0,  text=" SELECIONE O NÚMERO DE CÓPIAS ", fg="white", font="bold", bg=preto_claro)
Label_0.place(x=0,y=0, relwidth=1, relheight=0.40)

Label_1 = tk.Label(Frame_1,  text=" SELECIONE A IMPRESSORA ", fg="white", font="bold", bg=preto_claro)
Label_1.place(x=0,y=0, relwidth=1, relheight=0.40)

# SPINBOX

Spinbox_0 = tk.Spinbox(Frame_0, from_=1, to=10, font=("Arial", 25, "bold"), fg="white", buttonbackground=preto_claro, bg=cinza)
Spinbox_0.place(x=0,y=40, relwidth=1, relheight=0.65)
Spinbox_0.configure(justify='center')

# OPTION MENU

Escolha_0 = tk.StringVar(Frame_1)
Escolha_0 .set("NOTEBOOK")
Lista_escolhas = ["ADM",  "AK", "MOBILE", "NOTEBOOK"]
Escolha_principal = tk.OptionMenu(Frame_1 , Escolha_0, *Lista_escolhas, ) # O "*" DESCOMPACTA CADA ITEM DA LISTA "Lista_escolhas"
Escolha_principal.place(x=0,y=40, relwidth=1, relheight=0.65)
Escolha_principal.config(font=("Arial", 25, "bold"), fg="white", bg=cinza)

# BUTTOM

Button_0 = tk.Button(Frame_2,text="IMPRIMIR", font=("Arial", 25, "bold"), fg=preto_claro, bg=cinza, command=Imprimir)
Button_0.pack(fill=tk.BOTH, expand=True)

Button_0.bind("<Enter>", Click_Button_0)  # Quando o mouse entra ou sai do botão
Button_0.bind("<Leave>", Click_Button_0)

Janela_0.mainloop() # MANTÉM A JANELA ABERTA ATÉ SER FECHADA

# FIM DO CÓDIGO --------------------------------------------------------------------------------------------------------
