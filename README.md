# **Target:** Task automation in Excel and sending via email

_12/04/2021_

![](Python-Excel.jpeg)

[Font Image](https://www.pyxll.com/)




<font color = 'grey'>Description of Process:</font>
 
- Every day, there is a significant delay in the execution of the creation of the Excel spreadsheet and in sending it by email. The spreadsheet contains various pieces of information that are consolidated into reports that are extremely important for the business area. They are generated through VBA Macros and delivered to the manager
---
 
- The first problem is that, even with many codes to generate this automatic spreadsheet and send it to the manager, there is a considerable amount of time spent processing everything, and it has to be done daily and delivered to the manager first thing in the morning. This results in someone having to arrive 15 to 30 minutes earlier every day to run the macros and send the spreadsheet to the manager. If there is an error in loading the data, the process needs to be redone in the macro, causing rework.
----
  
- In simple terms, to address the problem until all the data and business rules can be migrated to a more robust and ideal tool such as Tableau or Power BI, this code was developed to meet the manager's needs by generating the report earlier, automatically, and sending it by email. This reduces to zero the overtime required by someone to perform this task manually
---
  
 - There are VBA codes in the spreadsheet used as the main object, which perform various update processes, such as pulling data from other spreadsheets and a CRM portal via a web browser. Therefore, this script will not detail that part. It will only focus on executing the VBA macro 'click' using Python and sending it by email to the manager.

 - # Required Installations
 - #### In this script, I am using Anaconda and saving it as a .py file. Although Jupyter Notebook includes various packages and libraries, it is necessary to install some additional ones for this mini-task.
```
!pip install pyautogui    # Para controlar seu pc
!pip install pywin32      # Para controlar programas no pc como outlook
```
```
# Imports

import pyautogui                    # Para controlar o pc 
import time                         # Tempo de execução específico                    
import pyperclip
import os                           # Para navegar e interagir com diretórios 
import webbrowser                   # Para controlar o navegador
import win32com.client as win32     # Para controlar programs no pc

```
```
webbrowser.open_new('https://site')

# Esperar
time.sleep(3)

# Localização por coordenadas x,y
pyautogui.moveTo(210,475, duration = 0.25)
pyautogui.click(210,475, button = 'left', duration = 0.25)

time.sleep(4)

# Abrir a pasta/planilha excel
os.startfile(r"C:\Users\elvis.monteiro\Desktop\Teste para envio fora de perfil\Fora do Perfil - Atualização Diária 2021.xlsb")

pyautogui.PAUSE=6

#Executar macro do excel
pyautogui.click(62,296)

# Esperar 15 minutos (tempo em segundos)
time.sleep(900)

#Fechar planilha
pyautogui.hotkey('alt','f4')

#Esperar a planilha fechar e salvar automaticamente(Já com programação para salvar direto por vba)
time.sleep(5)
```


# Send by email
```
# Import

import win32com.client as win32
```
```
#Criar integração com o  outlook
outlook = win32.Dispatch('outlook.application')

# Criar um email
email = outlook.CreateItem(0)

#configurar informações de email - email exemplo
email.To = " emailexemplo1@.com;emailexemplo2@.com.br" 
email.Subject ="Assunto"

email.HTMLBody = """
<p>Olá, Bom dia ! Segue em anexo a planilha XYZ atualizada 2021.</p>

<p>Att 

Elvis M.</p>

"""
#Enviar Anexo
anexo = r"C:\Users\elvis.monteiro\Desktop\Teste para envio fora de perfil\Fora do Perfil - Atualização Diária 2021.xlsb"
email.Attachments.Add(anexo)

email.Send ()
print ("email enviado")
```
# Start of Execution

    1 - Este Script é feito no Jupiter Notebook e salvo em formato .py. Sistema Operacional usado Windows.
    
    2 - É usado o agendador de tarefas para iniciar o processo automático dentro do horário estipulado no computador.(Talvez seja necessário acesso root da máquina).

    3 - Basta criar uma basta no agendador de tarefas nomeando a tarefa e subindo o script .py.


## Conclusion and Considerations
In order to measure in monetary terms the gain from this small script made in 15 minutes, consider this simple calculation: A professional who arrives approximately ≈30 minutes earlier every day, earning 8k per month and working 22 days. They would earn approximately ≈36.36 reais per hour with 220 hours per month, having approximately ≈0.61 reais per minute. Therefore, the value for 30 minutes is approximately ≈18.18 reais, or about 400 reais per month. Over a year, disregarding all labor-related calculations, the expense would be approximately ≈4,800 reais per year, which would be reduced to zero

# End

### Links and References

* Docs: [Python-PyAutoGUI](https://pyautogui.readthedocs.io/en/latest/)
* Book: [Automate the Boring Stuff with Python](https://www.amazon.com.br/Automate-Boring-Stuff-Python-2nd/dp/1593279922/ref=asc_df_1593279922/?tag=googleshopp00-20&linkCode=df0&hvadid=379726160779&hvpos=&hvnetw=g&hvrand=17894222063597453754&hvpone=&hvptwo=&hvqmt=&hvdev=c&hvdvcmdl=&hvlocint=&hvlocphy=9074180&hvtargid=pla-842272648989&psc=1&mcid=21d65bd15b84302d865dbcc8252b84bc)

