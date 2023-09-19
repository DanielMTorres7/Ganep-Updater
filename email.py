##################### Importando as bibliotecas do projeto #####################
import pyautogui
import time
import pyperclip
from datetime import datetime, time as dtTime
import subprocess

pyautogui.PAUSE = 1
hora_updt = 0
########################################## Definição da FUNÇÕES ##########################################
def click(x,y,duplo=False):
    old = pyautogui.PAUSE
    pyautogui.PAUSE = .05
    pyautogui.click(x, y)
    if duplo: pyautogui.click(x, y)
    pyautogui.PAUSE = old

def press(a,b=None,c=None,tempo=1):
    if c: pyautogui.hotkey(a,b,c)
    elif b: pyautogui.hotkey(a,b)
    else: pyautogui.hotkey(a)
    delay(tempo)
    print(pyautogui.PAUSE)
    return True
                
# Função para dar ALT + F4
def fechar(pasta = ""):
    if not pasta: 
        press('alt','f4')
    elif "exe" in pasta:
        subprocess.call(["taskkill","/F","/IM", pasta])
    elif pasta == "explorer":
        subprocess.call(["taskkill","/F","/IM", "explorer.exe","/FI", "WINDOWTITLE eq Planilhas"])
        subprocess.call(["taskkill","/F","/IM", "explorer.exe","/FI", "WINDOWTITLE eq updater"])
        subprocess.call(["taskkill","/F","/IM", "explorer.exe","/FI", "WINDOWTITLE eq Mapa Atual"])
        subprocess.call(["taskkill","/F","/IM", "explorer.exe","/FI", "WINDOWTITLE eq Explorador de Arquivos"])
    else:
        subprocess.call(["taskkill","/F","/IM", "explorer.exe","/FI", "WINDOWTITLE eq "+pasta+".csv - Resultados da Pesquisa em updater"])
    


def delay(tempo):
    time.sleep(tempo)

def esperarTela(path, tempo=0, conf=.95, vezes=0, clicar=False, duplo=False, confirmacao='', restart=True, esperaMax=25, pos=[0,0,1600,900], posconf=[0,0,1600,900]):
    tentativas = 0
    confiabilidade = conf

    while tentativas <= vezes or vezes == 0:
        tentativas += 1
        print(f"Procurando {path} pela {tentativas}° vez")

        if tentativas > 20 and confiabilidade > .60:
            confiabilidade -= .005

        try:
            x, y = pyautogui.locateCenterOnScreen(path, confidence=confiabilidade, region=pos)
            print(f"{path} encontrado na posição: {x}, {y}")

            if clicar or duplo:
                click(x, y, duplo)

            if confirmacao and not confirmarClick(confirmacao, esperaMax, posconf, restart):
                continue

            time.sleep(tempo)
            return True

        except Exception as e:
            print(f"não achou {path} confiabilidade = {confiabilidade}")

    print("tentativas encerradas")
    return False

def confirmarClick(confirmacao, esperaMax, posconf, restart):
    if esperarTela(confirmacao, 1, .85, vezes=esperaMax*2, pos=posconf, restart=False):
        print(f"Foi confirmado que o click em {confirmacao} foi bem sucedido")
        return True
    elif restart:
        print("Confirmação de click mal sucedida reiniciando processo")
        return False
    else:
        return False

# Função para clicar em uma imagem quando ela aparecer
def clicar(path, tempo=0, conf=.95, duplo = False, vezes=0,confirmacao='',restart=True,esperaMax=0, pos=[0,0,1600,900], posconf=[0,0,1600,900]):
    try:
        return esperarTela(path, tempo, conf, vezes=vezes, clicar=True, duplo=duplo,confirmacao=confirmacao,restart=restart,esperaMax=esperaMax, pos=pos, posconf=posconf)
    except:
        return False        


def abrirGooglePage(link, confirmacao=None):
    press('win','d')
    clicar('win_safeclick.png')
    clicar('appchrome.png',1 , 0.85, duplo = True,pos=[0, 259, 400, 659])
    delay(3)
    clicar('restaurar.png',vezes=20) 
    delay(3)
    clicar('restaurar.png',vezes=20) 
    press('esc')
    condicao = False
    if esperarTela('pesquisagoogle.png', 1, .8, vezes=20,pos=[0, 0, 400, 254]) or (condicao := esperarTela('search_google.png', 1, vezes=20)):
        if condicao: clicar('search_google.png', 1, 0.85) 
        write(link,3,True)
        while confirmacao and not esperarTela(confirmacao, 1,.85, esperaMax=3, vezes=20):
            press('f6')
            write(link,3,True)
    elif not esperarTela('opened_google.png', 1, .8, vezes=20):
        press('win', 'd')
        abrirGooglePage(link)
        

def excluirMapaAnterior(path):
    press('win','e', tempo=1)
    press('win','up')
    clicar(path, 3)
    press('ctrl','a', tempo=1)
    press('del', tempo=3)
    press('enter')
    #explorer()
    fechar("explorer")

def baixarPlanilhaGoogle(url):
    global hora_updt
    hora_updt = datetime.now().strftime("%d/%m/%Y as %H:%M")
    abrirGooglePage(url)
    delay(20)
    sequenciadownload()
    esperarTela('confirmacao.png',5,vezes=150)
    # google()
    delay(1)
    fechar("chrome.exe")

def sequenciadownload():
    condicoes = [
    clicar("sheets_arquivo.png",1,.8,pos=[0, 0, 386, 323],confirmacao="sheets_download.png",vezes=60,esperaMax=30,restart=False),
    clicar("sheets_download.png",1,.8,confirmacao="sheets_xlsx.png",vezes=80,esperaMax=30,restart=False),
    clicar("sheets_xlsx.png",4,.8,vezes=80,esperaMax=30,restart=False)]
    if all(condicoes):
        return True
    else:
        press('esc')
        sequenciadownload()

def atualizarMapa():
    excluirMapaAnterior('fav_planilhas.png')
    baixarPlanilhaGoogle("https://docs.google.com/spreadsheets/d/1b9dy9IRLpEhExVFKtTN2tcIqMni8gKwHzeWWnzYGklU/edit#gid=0")
    baixarPlanilhaGoogle("https://docs.google.com/spreadsheets/d/1PdmNUOkLCW4Wp-hN-wsLe2ZcJZUMyVJKYXcQMTyz7oY/edit?usp=drivesdk")
    sequenciaMapa()

def sequenciaMapa():
    press('win','e',tempo=5)
    clicar('fav_planilhas.png', 3, .9,pos=[ 0, 152, 400, 552])
    press('ctrl','a',tempo=1)
    press('ctrl','c',tempo=5)
    clicar('fav_mapa.png', 5, .9,pos=[0, 129, 400, 529])
    press('ctrl','v')
    if esperarTela("opened_substituir.png",2,vezes=120) and clicar("substituir.png",vezes=60):
        esperarTela('substituicaocompleta.png')
        #explorer()
        delay(80)
        fechar("explorer")
    else:
        #explorer()
        delay(80)
        fechar("explorer")
        sequenciaMapa()

def google():
    pyautogui.moveTo(1,899)
    clicar('icon_google.png',3,.9,confirmacao='opened_sheets.png',esperaMax=60,restart=False) if not esperarTela('opened_sheets.png',vezes=20) else clicar('google_safeclick.png',pos=[991, 0, 1600, 352])

def write(texto, tempo=1, enter=False):
    pyperclip.copy(texto)
    press('ctrl','v')
    if enter: 
        delay(1)
        press('enter')
    delay(tempo)
    return True
    

def anexar(arquivo):
    condicoes =[
    clicar('locaweb_anexar.png',5,.9,confirmacao='locaweb_escolher_arquivo.png',vezes=40,esperaMax=20,restart=False),
    clicar('locaweb_escolher_arquivo.png',13,.9,confirmacao='mapa_atual.png',vezes=40,esperaMax=20,restart=False),
    write(arquivo),
    press('enter',tempo=10),
    clicar('locaweb_enviar.png',10,.85,vezes=40)]
    return True if all(condicoes) else anexar()

def enviaremail():
    global hora_updt
    abrirGooglePage("https://webmail-seguro.com.br/dstorres.com.br/",'opened_locaweb.png')
    if esperarTela('locaweb_entrar.png',1,.9,10): clicar('locaweb_entrar.png',1,.9)
    
    clicar('locaweb_criar_email.png',5 ,.9,confirmacao='locaweb_anexar.png',esperaMax=10)
    write("<administracao@ganeplar.com.br>, <qualidade1@ganeplar.com.br>, <comercial@ganeplar.com.br>, <implantacao1@ganeplar.com.br>")
    press('tab',tempo=1)
    write("Atualização do Mapa")
    press('tab')
    write("Bom dia."+
            "\nSegue link com planilha atualizada em tempo real, e uma cópia atualizada em " + str(hora_updt) + 
            "\nMapa - https://docs.google.com/spreadsheets/d/1b9dy9IRLpEhExVFKtTN2tcIqMni8gKwHzeWWnzYGklU/edit#gid=0"+
            "\nAtt Daniel Torres.")
    anexar("\"Mapa.xlsx\" ")
    esperarTela('locaweb_envio_final.png',5,.9)
    clicar('locaweb_envio_final.png',40,.9)

    clicar('locaweb_criar_email.png',5 ,.9,confirmacao='locaweb_anexar.png',esperaMax=10)
    write("<administracao@ganeplar.com.br>, <qualidade1@ganeplar.com.br>, <comercial@ganeplar.com.br>, <implantacao1@ganeplar.com.br>")
    press('tab',tempo=1)
    write("Atualização da Margem dos Implantados")
    press('tab')
    write("Bom dia."+
            "\nSegue link com planilha atualizada em tempo real, e uma cópia atualizada em " + str(hora_updt) + 
            "\nMargem dos Implantados - https://docs.google.com/spreadsheets/d/1PdmNUOkLCW4Wp-hN-wsLe2ZcJZUMyVJKYXcQMTyz7oY/edit?usp=drivesdk"+
            "\nAtt Daniel Torres.")
    anexar("\"Margem dos Implantados.xlsx\"")
    esperarTela('locaweb_envio_final.png',5,.9)
    clicar('locaweb_envio_final.png',10,.9)


    fechar("chrome.exe")

###################################################################################################################################################



########################################## Abrindo o IW-Ganeplar ##########################################
#minimizando a aplicação
press('win','down')

atualizarMapa()
enviaremail()

###################################################################################################################################################
