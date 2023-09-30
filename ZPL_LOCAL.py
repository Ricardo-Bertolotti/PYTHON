import subprocess
import os


# VARIÁVEL COM O ZPL DO LAYOUT
layout ='''
CT~~CD,~CC^~CT~
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
^FT0,140^A0N,68,68^FB562,1,17,C^FH\^CI28^FDUUUUUUUUUHHH^FS^CI27
^FT0,225^A0N,68,68^FB562,1,17,C^FH\^CI28^FDZÉ^FS^CI27
^FT0,310^A0N,68,68^FB562,1,17,C^FH\^CI28^FDDA^FS^CI27
^FT0,395^A0N,68,68^FB562,1,17,C^FH\^CI28^FDMANGAAAAAA^FS^CI27
^PQ1,0,1,Y
^XZ
'''

# CAMINHO ONDE SERÁ CRIADO O ARQUIVO .PRN
caminho_arquivo = os.path.join(os.environ['TEMP'], 'layout.prn')
with open(caminho_arquivo, 'wb') as f:
    f.write(layout.encode('utf-8-sig'))


# VARIÁVEIS DO COMANDO COPY
maquina = 'RB7-ENG-SW'
impressora = 'R - SALA SW ZD220'
caminho = '%temp%\layout.prn'

# COMANDO COPY
comando = 'COPY "{0}" "\\\\{1}\\{2}"'.format(caminho, maquina, impressora)

# EXECUÇÃO DO COMANDO COPY
process = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
