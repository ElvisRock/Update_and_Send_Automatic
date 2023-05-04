# Objetivo: Atualizar planilha e enviar por e-mail automática para ganhar tempo

# Neste jupyter notebook estou usando o Anaconda, embora possua diversas bibliotecas já incluso é necessário instalar algumas para esta demanda.
# instale !pip install pyautogui para controlar seu pc
# instale !pip install pywin32 Para controlar programas no pc, como o outlook, instale 

import pyautogui
import time
import pyperclip
import os
import webbrowser
import win32com.client as win32

#Atualizar planilha fora de perfil e enviar por e-mail 
time.sleep (5)
#Condição para a atualização da planilha, página do vetor aberta e logada.
webbrowser.open_new('https://vetorzkm.movida.com.br/login.php?logout=1')
time.sleep(3)
pyautogui.moveTo(210,475, duration = 0.25)
pyautogui.click(210,475, button = 'left', duration = 0.25)
#Esperar
time.sleep(4)
#abrir a pasta/planilha excel
os.startfile(r"C:\Users\elvis.monteiro\Desktop\Teste para envio fora de perfil\Fora do Perfil - Atualização Diária 2021.xlsb")
#Esperar
pyautogui.PAUSE=6
#Clicar no botão e rodar a macro do excel
pyautogui.click(62,296)
#esperar 5 minutos ou 300 segundos
time.sleep(300)
#Fechar planilha
pyautogui.hotkey('alt','f4')
#Esperar a planilha fechar e salvar automaticamente
time.sleep(5)



#Enviar e-mail
import win32com.client as win32

#Criar integração com o  outlook
outlook = win32.Dispatch('outlook.application')
# Criar um email
email = outlook.CreateItem(0)
#configurar informações de email - email exemplo
email.To = " emailexemplo1@.com;emailexemplo2@.com.br" 
email.Subject ="Planilha Fora de Perfil"
email.HTMLBody = """
<p>Olá, Bom dia ! Segue em anexo a planilha Fora de Perfil atualizada 2021.</p>

<p>Att 

Elvis M.</p>

"""
#Enviar Anexo
anexo = r"C:\Users\elvis.monteiro\Desktop\Teste para envio fora de perfil\Fora do Perfil - Atualização Diária 2021.xlsb"
email.Attachments.Add(anexo)

email.Send ()
print ("email enviado")

