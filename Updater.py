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
    confiabilidade = conf
    tentativas = 0
    # Criando um loop de espera pela imagem
    while tentativas <= vezes or vezes == 0:
        # Colocando dentro de um try para caso a imagem ainda não esteja na tela
        try:
            tentativas+=1
            # Procurando pela imagem na tela
            print("Procurando " + str(path) + " pela " + str(tentativas) + "° vez")
            if tentativas > 20 and confiabilidade > .60: confiabilidade -=.005
            x, y = pyautogui.locateCenterOnScreen(path, confidence=confiabilidade, region=pos)
            print(str(path) + " encontrado na posição: " + str(x) + ", " + str(y))
            delay(1.5)
            if clicar or duplo: click(x,y,duplo)
            # Saindo do loop
            if confirmacao:
                if esperarTela(confirmacao, 1, .85, vezes = esperaMax*2, pos= posconf,restart=False):
                    print("Foi confirmado que o click em " + str(path) + " foi bem sucedido")
                    delay(tempo)
                    return True
                elif restart:
                    print("Confirmação de click mal sucedida reiniciando processo")
                    tentativas=0
                else:
                    return False
            else:
                delay(tempo)
                return True
        except:
            delay(.5)
            print("não achou " + str(path) + " confiabilidade = " + str(confiabilidade))
            if tentativas >= vezes > 0:
                print("tentativas encerradas")
                return False


# Função para clicar em uma imagem quando ela aparecer
def clicar(path, tempo=0, conf=.95, duplo = False, vezes=0,confirmacao='',restart=True,esperaMax=0, pos=[0,0,1600,900], posconf=[0,0,1600,900]):
    try:
        return esperarTela(path, tempo, conf, vezes=vezes, clicar=True, duplo=duplo,confirmacao=confirmacao,restart=restart,esperaMax=esperaMax, pos=pos, posconf=posconf)
    except:
        return False        
    
def baixarPlanilha(path,tempo):
    # Caminho para baixar a planilha
    delay(2)
    clicar('relatorios.png', 1,.85,confirmacao='opened_relatorios.png',pos=[135, 0, 800, 233],posconf=[98, 0, 898, 350])
    if not clicar('planilhas.png', 1,.85,confirmacao='opened_planilhas.png',restart=False,pos=[98, 45, 898, 445],posconf=[177, 111, 977, 511]):
        baixarPlanilha(path,tempo)
    else:
        # selecionando a planilha
        clicar(path, tempo, .8)

# Função para baixar uma planilha do IW-Ganeplar
def salvar(path,nome,tempo = 1):

    baixarPlanilha(path,tempo)
    # Preparando a planilha para exportação
    clicar('exportar.png',5,.7,pos=[ 299, 0, 899, 413])
    clicar('desktop.png', conf=0.85,duplo = True, pos=[293, 81, 893, 681])
    clicar('iwganeplar.png',duplo = True, pos=[284, 40, 884, 640])
    clicar('pasta.png',duplo = True, pos=[285, 39, 885, 639])
    clicar('pesquisa.png', pos=[499, 208, 1099, 808])

    # Nomeando a planilha
    write(nome)

    # Salvando a planilha em Desktop    
    clicar('salvar.png', 3, .7, pos=[614, 277, 1214, 877])

    # Fechando a planilha
    fechar()
    while not esperarTela('rel_fechado.png',vezes=80): fechar()

# Função para atualizar uma planilha do google
def atualizar(tabela,tabelaGoogle):
    achou = False
    while not achou:
        # Abrindo o explorador de arquivos no desktop
        press('win','e', tempo=2)
        press('win','up')
        if clicar('fav_updater.png',conf=.8, vezes=80, pos=[ 0, 77, 600, 677]):
            # Foi verificado travamentos nesta parte então a função vai fazer a verificação da existencia da aba pesquisa na tela
            if clicar('pesquisarupdater.png',5,vezes=50, pos=[1118, 0, 1600, 359]):
                # Como o computador costuma travar será apertado f3 duas vezes para garantir a pesquisa
                press('f3', tempo=1)

                # Pesquisando pela planilha na pasta
                write(tabela+".csv",5)
                press('enter')
                

                if clicar('icon_planilha.png',duplo=True,confirmacao='opened_excel.png',vezes=40,esperaMax=20,restart=False):
                    achou=True
                else:
                    print ("Não foi encontrado o item desejado, processo será reiniciado 2")
            else:
                print ("Não foi encontrado o item desejado, processo será reiniciado 2")
        else:
            print ("Não foi encontrado o item desejado, processo será reiniciado 3")
        #explorer()
        fechar(tabela)

    

    # Selecionando e copiando todos os dados da planilha
    excel()
    press('ctrl','t',tempo=8)
    press('ctrl','c')
    delay(12) if not tabela == "orcamentos" else delay(30)
    planilhaCopiada = pyperclip.paste()
    press('esc')
    delay(2)
    fechar("excel.exe")

    # Voltando pra planilha do google
    google()

    # selecionando a planilha
    if tabelaGoogle != "none" : clicar(tabelaGoogle, 10, .85)
    
    # Inserindo os dados
    press('esc')
    pyperclip.copy(planilhaCopiada)
    delay(2)
    pyperclip.copy(planilhaCopiada)
    delay(1)
    press('ctrl','shift','v')
    esperarTela('sheets_pasted.png')
    delay(2) if not tabela  == "orcamentos" else delay(7)
    


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

def excel():
    pyautogui.moveTo(1,899)
    clicar('icon_excel.png',3,.9,confirmacao='opened_excel.png',esperaMax=60) if not esperarTela('opened_excel.png',vezes=20) else clicar('excel_safeclick.png',1,.6,pos=[ 100, 35, 530, 260])
 
# def explorer():
#     pyautogui.moveTo(1,899)
#     clicar('icon_explorer.png', 1, .9,confirmacao='opened_explorer.png',esperaMax=60,restart=False) if not esperarTela('opened_explorer.png',vezes=10) else clicar('explorer_safeclick.png')

def write(texto, tempo=1, enter=False):
    pyperclip.copy(texto)
    press('ctrl','v')
    if enter: 
        delay(1)
        press('enter')
    delay(tempo)
    return True
    
def in_between(now, start, end):
    return start <= now < end

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
    if dtTime(9) <= datetime.now().time() < dtTime(10):
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

def baixartudo():
    # Abrindo o "IW-Ganep"
    clicar('IW-Ganeplar.png', 1, 0.85, duplo = True, confirmacao='login.png', esperaMax=60,pos=[0, 0, 340, 339])

    # Realizando Login
    clicar('login.png', 2, 0.9,pos=[497, 313, 1097, 713])
    write("Ganeplar@7823")
    delay(1)
    press('enter',tempo=60)

    ###################################################################################################################################################

    ########################################## Baixando as planilhas do IW-Ganeplar ##########################################

    # Atendimentos
    salvar('dados_atendimento.png',"atendimentos_normal")

    # Intercorrencias
    salvar('dados_intercorrencias.png',"Intercorrencias")

    # CCID
    salvar('dados_ccid.png',"CCID ")

    # Atendimentos Completo
    salvar('dados_atendimento_completo.png',"atendimento_completo")

    # Equipe
    salvar('dados_equipe.png',"Equipes")

    # Orçamentos
    salvar('dados_orcamentos.png',"Orcamentos")

    # Fechando IW_Ganeplar
    fechar()
    delay(1)
    clicar('confirmar_fechar.png',pos=[ 375, 260, 1175, 660]) 



########################################## Abrindo o IW-Ganeplar ##########################################
#minimizando a aplicação
press('win','down')
fechar("explorer")


###################################################################################################################################################
baixartudo()
########################################## Atualizando a planilha do Google ##########################################

# Abrindo a planilha do google 
abrirGooglePage("https://docs.google.com/spreadsheets/d/19xmv5ijBpgfQsukWsTx9B-otjf2X0gVHZ65fuKaunZ0/edit",'opened_sheets.png')

# Planilha "Atendimento"  
atualizar("atendimentos_normal","none")

# Planilha "Intercorrencias"  
atualizar("Intercorrencias",'plan_intercorrencias.png')
        
# Planilha "CCID"  
atualizar("ccid",'plan_ccid.png')

# Planilha "Atendimentos Completo" 
atualizar("atendimento_completo",'plan_atendimento_completo.png')

# Planilha "Equipe" 
atualizar("Equipes",'plan_equipe.png')

# Planilha "Orçamentos" 
atualizar("orcamentos",'plan_orcamentos.png')

# google()
fechar("chrome.exe")
press('enter')

atualizarMapa()
enviaremail()

###################################################################################################################################################
