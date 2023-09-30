# IMPORTAÇÕES ----------------------------------------------------------------------------------------------------------

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import sys
import socket
import tempfile
from docx2pdf import convert

# FUNÇÕES --------------------------------------------------------------------------------------------------------------
def Titulo_janela(janela, titulo): # FUNÇÃO PARA DEFINIR O TÍTULO DA JANELA
    janela.title(titulo)

def Cor_janela(janela, cor): # FUNÇÃO PARA DEFINIR A COR DE FUNDO DA JANELA
    janela.configure(bg=cor)

def Dimensionar_janela(janela, largura, altura): # FUNÇÃO PARA DEFINIR AS DIMENSÕES DA JANELA
    janela.geometry(f"{largura}x{altura}")

def Centralizar_janela(janela): # FUNÇÃO PARA CENTRALIZAR A JANELA NA TELA
    janela.update_idletasks()
    X = (janela.winfo_screenwidth() - janela.winfo_width()) // 2
    Y = (janela.winfo_screenheight() - janela.winfo_height()) // 2
    janela.geometry(f"{janela.winfo_width()}x{janela.winfo_height()}+{X}+{Y}")

def Botao1_buscador(): # FUNÇÃO PARA OBTER O CAMINHO E O NOME DE UM ARQUIVO SELECIONADO
    global Arquivo_caminho
    Arquivo_caminho = filedialog.askopenfilename()
    if Arquivo_caminho:
        Nome_arquivo = os.path.basename(Arquivo_caminho)
        Label1['text'] = Nome_arquivo
    else:
        messagebox.showerror("ERRO", "Nenhum arquivo foi selecionado!")

def Botao2_limpar(Entry): # FUNÇÃO PARA LIMPAR OS DADOS INSERIDOS EM UM ENTRY
    Entry.delete(0, tk.END)

def Botao3_confirmar(): # FUNÇÃO PARA CONFIRMAR E VALIDAR A ENTRADA DOS DADOS DO "Entry1"
        Entry1_data = Entry1.get()
        try:

            Entry1_data = int(Entry1_data)
            if Entry1_data <= 0:
                icon = "warning"  # info, error, warning, question, (caminho_icone)
                messagebox.showerror("ERRO", "ERRO : O VALOR INSERIDO É INVÁLIDO!", icon=icon)
            else:
                messagebox.showerror("CONFIRMAÇÃO", "O NÚMERO FOI REGISTRADO COM SUCESSO!", icon="info")

        except ValueError as e:
            messagebox.showerror("ERRO", "ERRO : O VALOR INSERIDO É INVÁLIDO!",
                                 icon="error")
            return

def Botao4_imprimir(): # FUNÇÃO PARA SOLICITAR A IMPRESSÃO DE UM ARQUIVO

    # DEFINIR O ENDEREÇO IP E A PORTA DA IMPRESSORA
    printer_ip = '192.168.107.131'
    # 192.168.107.128 ADM MG-ADM-COLOR
    # 192.168.107.131 NOTEBOOK
    # 192.168.107.104 AK
    # 192.168.107.XXX MOBILE

    printer_port = 9100

    # VARIÁVEL PARA OBTER A EXTENSÃO DO ARQUIVO
    extension = os.path.splitext(Arquivo_caminho)[1]

    # CONDIÇÃO PARA VALIDAR SE OS ARQUIVOS ESTÃO EM FORMATO .PDF
    if extension != ".pdf":

        # COMANDO PARA CONVERTER .DOCX EM .PDF
        #Arquivo_caminho_pdf = r"C:\Users\r1c4r\OneDrive\Área de Trabalho\test.pdf"
        Arquivo_caminho_pdf = os.path.join(tempfile.gettempdir(), 'Py_pdf_spool.pdf')
        convert(Arquivo_caminho, Arquivo_caminho_pdf)
        remove = bool(True)

    else:

        remove = bool(False)
        Arquivo_caminho_pdf = Arquivo_caminho

    # ABRIR O ARQUIVO E FAZER A SUA LEITURA EM BINÁRIO
    with open(f'{Arquivo_caminho_pdf}', 'rb') as file:
        Data = file.read()

    # CRIAR UM SOCKET DO TIPO TCP/IP PARA SE COMUNICAR COM A IMPRESSORA
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.connect((printer_ip, printer_port))

    # VARIÁVEL QUE DEFINE A QUANTIDADE DE CÓPIAS DO ARQUIVO
    Entry1_data = Entry1.get()

    # ESTRUTURA DE REPETIÇÃO PARA ENVIAR O ARQUIVO DE ACORDO COM A QUANTIDADE DE CÓPIAS DO ARQUIVO
    for i in range(int(Entry1_data)):

        # ENVIAR O CONTEÚDO DO ARQUIVO PARA A IMPRESSORA
        sock.sendall(Data)

    if remove is True:
        os.remove(Arquivo_caminho_pdf)

    messagebox.showerror("CONFIRMAÇÃO", "A IMPRESSÃO FOI SOLICITADA COM SUCESSO!", icon="info")

    # Fechar a conexão com a impressora
    sock.close()

def Botao5_calibrar(): # FUNÇÃO PARA SOLICITAR A CALIBRAÇÃO DA IMPRESSORA
    print("CALIBRAR")

def Botao6_sair(janela): # FUNÇÃO PARA FECHAR UMA JANELA
    janela.destroy()

def Enter_confirmar(func): # FUNÇÃO PARA INTERLIGAR O "ENTER" COM OUTRA FUNÇÃO, AO SER PRESSIONADO
    Entry1.bind('<Return>', lambda event: func())

# INICIO DO CÓDIGO -----------------------------------------------------------------------------------------------------

Interface = tk.Tk() # INICIA A JANELA

# FUNÇÃO PARA DEFINIR O NOME DA JANELA
Titulo_janela(Interface, "IMPRESSÃO DE ETIQUETAS OFFLINES")

# FUNÇÃO PRA DEFINIR A COR DE FUNDO DA JANELA
Cor_janela(Interface, "black")

# FUNÇÃO PARA DEFINIR AS DIMENSÕES DA JANELA
Dimensionar_janela(Interface, 380, 550)

Centralizar_janela(Interface)

# FRAMES
Frame1 = tk.Frame(Interface, bg="yellow")
Frame1.pack(fill="both", expand=True, padx=20, pady=20)

Frame2 = tk.Frame(Frame1, bg="black")
Frame2.pack(fill="both", expand=True, padx=5, pady=5)

Frame3 = tk.Frame(Frame2, bg="yellow", height=50)
Frame3.place(x=20,y=20, relwidth=0.88, relheight=0.28)

Frame4 = tk.Frame(Frame2, bg="yellow", height=50)
Frame4.place(x=20,y=180, relwidth=0.88, relheight=0.28)

Frame5 = tk.Frame(Frame2, bg="yellow", height=50)
Frame5.place(x=20,y=340, relwidth=0.88, relheight=0.28)

Frame6 = tk.Frame(Frame3, bg="black")
Frame6.pack(fill="both", expand=True, padx=5, pady=5)

Frame7 = tk.Frame(Frame4, bg="black")
Frame7.pack(fill="both", expand=True, padx=5, pady=5)

Frame8 = tk.Frame(Frame5, bg="black")
Frame8.pack(fill="both", expand=True, padx=5, pady=5)

# BOTÕES
Botao1 = tk.Button(Frame6, text="BUSCAR ARQUIVO", bg="navy", fg="white", command=Botao1_buscador)
Botao1.pack(fill="both", padx=20 , pady=(20,0))

Botao2 = tk.Button(Frame7, text=" LIMPAR ", bg="navy", fg="white", command=lambda: Botao2_limpar(Entry1))
Botao2.place(x=20,y=90, relwidth=0.4)

Botao3 = tk.Button(Frame7, text=" CONFIRMAR ", bg="navy", fg="white", command=Botao3_confirmar)
Botao3.place(x=150,y=90, relwidth=0.4)

Botao4 = tk.Button(Frame8, text=" IMPRIMIR ", bg="navy", fg="white", command=Botao4_imprimir)
Botao4.pack(fill="both", padx=20 , pady=(15,0))

Botao5 = tk.Button(Frame8, text=" CALIBRAR ", bg="navy", fg="white", command=Botao5_calibrar)
Botao5.pack(fill="both", padx=20 , pady=(10,0))

Botao6 = tk.Button(Frame8, text=" SAIR ", bg="navy", fg="white", command=lambda: Botao6_sair(Interface))
Botao6.pack(fill="both", padx=20 , pady=(10,10))

# LABELS
Label1 = tk.Label(Frame6, text=" ( Vazio ) ", bg="white")
Label1.pack(fill="both", expand=True, padx=20 , pady=(0,20))

Label2 = tk.Label(Frame7, text=" NÚMERO DE CÓPIAS ", bg="navy", fg="white")
Label2.pack(fill="both", expand=False, padx=20 , pady=(20,0))

Entry1 = tk.Entry(Frame7, bg="white", justify="center")
Entry1.pack(fill="both", expand=True, padx=20, pady=(0,50))
Enter_confirmar(Botao3_confirmar)

Interface.mainloop() # MANTÉM A JANELA ABERTA ATÉ QUE SEJA ENCERRADA

# FINAL DO CÓDIGO ------------------------------------------------------------------------------------------------------
