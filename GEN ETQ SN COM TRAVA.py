import os                       # BIBLIOTECA PARA USAR FUNÇÕES DO SISTEMA OPERACIONAL
import subprocess               # BIBLIOTECA PARA USAR FUNÇÕES DO SISTEMA OPERACIONAL
from time import sleep          # BIBLIOTECA PARA USAR A FUNÇÃO DE TEMPORIZADOR

# VARIÁVEL PARA OBTER O LIMITE DE CARACTERES DO NÚMERO DE SÉRIE
trava = int(10)

# VARIÁVEL PARA OBTER O PREFIXO QUE O NÚMERO DE SÉRIE DEVE TER
prefixo = str("WS125K")

# VARIÁVEL PARA OBTER O NÚMERO DE CARATERES DO PREFIXO
contador_prefixo = int(len(prefixo))

# FUNÇÃO PARA GERAR UMA LINHA NA TELA
def draw():
    desenho=("=-" * 50 + "=")
    print(desenho)

# FUNÇÃO PARA CRIAR UMA QUEBRA DE LINHA NA TELA
def space():
    print()

# ESTRUTURA DE REPETIÇÃO PARA VALIDAR O NÚMERO DE SÉRIE
while True:

    # ESTRUTURA PADRÃO PARA REQUISITAR O INPUT DO NÚMERO DE SÉRIE
    space(), draw()
    SN = input("INSIRA O NÚMERO DE SÉRIE : ")
    draw(), space()

    # CONDIÇÃO PARA VERIFICAR SE O NÚMERO DE SÉRIE INSERIDO POSSUI ESPAÇOS VAZIOS " "
    if " " in SN:

        print(f"ERRO: O NÚMERO DE SÉRIE : ->{SN}<- NÃO DEVE POSSUIR ESPAÇOS")

    # CONDIÇÃO PARA VERIFICAR SE O NÚMERO DE SÉRIE INSERIDO POSSUI APENAS CARACTERES ALFANUMÉRICOS
    elif not SN.isalnum():

        print(f"ERRO: O NÚMERO DE SÉRIE : {SN} DEVE POSSUIR APENAS CARACTERES ALFANUMÉRICOS")

    # CONDIÇÃO PARA VERIFICAR SE O NÚMERO DE SÉRIE INSERIDO POSSUI LETRAS MINÚSCULAS
    elif not SN.isupper():

        print(f"ERRO: O NÚMERO DE SÉRIE : {SN} NÃO PODE CONTER LETRAS MINÚSCULAS")

    # CONTRA CONDIÇÃO PARA VALIDAR OUTROS REQUISITOS ( CADASTROS ) APÓS A PRIMEIRA VALIDAÇÃO DO NÚMERO DE SÉRIE INSERIDO
    else:

        # VARIÁVEL PARA OBTER A QUANTIDADE DE CARACTERES DO NÚMERO DE SÉRIE INSERIDO
        validar_trava = len(SN)

        # VARIÁVEL PARA OBTER O PREFIXO DO NÚMERO DE SÉRIE INSERIDO
        validar_prefixo = SN[:contador_prefixo]

        # CONDIÇÃO PARA VALIDAR O PREFIXO E O LIMITE DE CARCATERES DO NÚMERO DE SÉRIE INSERIDO
        if (validar_trava != trava) or (validar_prefixo != prefixo):

            # CONDIÇÃO PARA VALIDAR O LIMITE DE CARCATERES DO NÚMERO DE SÉRIE INSERIDO
            if (validar_trava != trava):
                print(f"ERRO: O NÚMERO DE SÉRIE DEVE POSSUIR EXATAMENTE {trava} DÍGITOS")

            # CONDIÇÃO PARA VALIDAR O PREFIXO DO NÚMERO DE SÉRIE INSERIDO
            if (validar_prefixo != prefixo):
                print(f"ERRO: O PREFIXO DO NÚMERO DE SÉRIE ESTÁ INCORRETO, PREFIXO CORRETO: {prefixo}")

        # CONTRA CONDIÇÃO PARA CASO O NÚMERO DE SÉRIE INSERIDO ESTEJA TOTALMENTE VALIDADO
        else:

            # OUTPUT PARA AVISAR QUE O NÚMERO DE SÉRIE INSERIDO FOI VALIDADO E QUE A ETIQUETA SERÁ GERADA
            print(f"NÚMERO DE SÉRIE ({SN}) VALIDADO COM SUCESSO!!! GERANDO ETIQUETA...")
            space(), draw(), space()

            # COMANDO PARA ENCERRAR A ESTRUTURA DE REPETIÇÃO
            break

# VARIÁVEL COM O ZPL DO LAYOUT ( FEITO A PARTE ANTES DE INCLUIR E AJUSTAR )
layout ='''CT~~CD,~CC^~CT~
^XA
~TA000
~JSN
^LT0
^MNW
^MTT
^PON
^PMN
^LH0,0
^JMA
^PR4,4
~SD20
^JUS
^LRN
^CI27
^PA0,1,1,0
^XZ
^XA
^MMT
^PW563
^LL484
^LS0
^FO2,2^GB559,480,16,,0^FS
^BY2,3,68^FT82,276^BCN,,N,N
^FH\^FD>:''' + SN + '''^FS
^FT185,190^A0N,28,28^FH\^CI28^FD''' + SN + '''^FS^CI27
^PQ1,0,1,Y
^XZ
'''

# CAMINHO ONDE SERÁ CRIADO O ARQUIVO .PRN ( %temp%\layout.prn )
caminho = os.path.join(os.environ['TEMP'], 'layout.prn')
with open(caminho, 'wb') as f:
    f.write(layout.encode('utf-8-sig'))

# COMANDO PARA OBTER O NOME DA MÁQUINA DO SISTEMA OPERACIONAL
maquina = subprocess.check_output('hostname', shell=True).decode('utf-8').strip()

# COMANDO PARA OBTER O NOME DA IMPRESSORA PADRÃO DO SISTEMA OPERACIONAL
impressora = subprocess.check_output('wmic printer get name,default | findstr TRUE', shell=True).decode('utf-8').strip()[9:]

# PARAMETRIZAÇÃO DO COMANDO COPY ( PARA IMPRIMIR O ARQUIVO .PRN NA IMPRESSORA PADRÃO )
comando = 'COPY "{0}" "\\\\{1}\\{2}"'.format(caminho, maquina, impressora)

# EXECUÇÃO DO COMANDO COPY
process = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

# COMANDO PARA EXCLUIR O ARQUIVO TEMPORÁRIO .PRN GERADO ANTERIORMENTE, USANDO UM DELAY DE 1s PARA QUE O COMANDO COPY
# CONSIGA SER EXECUTADO PARA GERAR UM SPOOL COM O ARQUIVO .PRN ANTES QUE O MESMO SEJA EXCLUIDO.
sleep(1)
os.remove(caminho)
