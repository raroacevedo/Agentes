
from os import listdir
from selenium.common.exceptions import TimeoutException
from time import sleep
from webbrowser import Chrome
from numpy import select, true_divide
from pyshadow.main import Shadow
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import pyperclip
import pandas as pd

#crear instancia de web driver
def crear_WebDriver():

    WD = Service(r"..\Chrome\chromedriver.exe")

    chrome_options = Options()
    chrome_options.add_argument("--disable-extension")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

    # Desactivar el gestor de contraseñas de Chrome
    chrome_prefs = {
        "useAutomationExtension": False,
        "credentials_enable_service": False,                        # Desactiva el guardado automático de contraseñas
        "profile.password_manager_enabled": False,                  # Desactiva el gestor de contraseñas
        "profile.default_content_setting_values.local_network": 1,  # Evita el popup de red local permitiendo sin prompt
        "profile.default_content_setting_values.notifications": 2,  # Desactiva notificaciones
    }

    chrome_options.add_experimental_option("prefs", chrome_prefs)

    driver = webdriver.Chrome (service=WD, options=chrome_options)
    driver.implicitly_wait(40)

    return driver

#ingresar a la pagina de login
def login(ldriver, usr, pwd, dfa):
    try:
        # get the login page
        ldriver.get("https://virtual.upb.edu.co/d2l/login?noRedirect=1")
  
        #get input elemment - login
        username = WebDriverWait(ldriver, 3).until(
            EC.presence_of_element_located((By.ID, "userName")) 
        )
        password = WebDriverWait(ldriver, 3).until(
            EC.presence_of_element_located((By.ID, "password"))
        )

        # send the credentials
        username.send_keys(usr)
        #password.send_keys(pwd)
        password.send_keys("2/2*2Hijosmemsam")
        password.send_keys(Keys.RETURN)

        #2FA
        ldriver.get("https://virtual.upb.edu.co/d2l/lp/auth/twofactorauthentication/TwoFactorCodeEntry.d2l")
        
        l2fa = WebDriverWait(ldriver, 3).until(
            EC.presence_of_element_located((By.ID, "z_i"))
        )
        l2fa.send_keys(dfa)
        l2fa.send_keys(Keys.RETURN)

        sleep(3)

        print ("\n<<Si pudo ingresar con las credenciales>> ✔️")
        return True

    except:
        print ("\n**No pudo ingresar con las credenciales** ❌")
        return False

#cheCk la integralidad del archivo excel
def check_user_file(agente):
    # get the dataframe columns
    columns = list(agente.columns)
    if len(columns) != 5:
        return False
    return (
            "NRC" in columns
        and "IDCURSO" in columns
        and "IDAGENTE" in columns
        and "FECHAINICIO" in columns
        and "FECHAFINAL" in columns
    )

#************************************************************************
#Gestion de los agentes : habilitar, fecha inicio y fecha final
#************************************************************************
def setdataAgen(ldriver, dlrow):
    agente_url="https://virtual.upb.edu.co/d2l/le/intelligentagents/agent/"+str(dlrow["IDCURSO"])+"/edit?agentId="+str(dlrow["IDAGENTE"])
    print(agente_url)
   
    ldriver.get(agente_url)
    wait = WebDriverWait(ldriver, 15)

    # Identifica el idioma de la página
    language = ldriver.find_element(By.CSS_SELECTOR, "html").get_attribute("lang")

    try:
        #**********************
        #datos del agente
        #**********************
        #field 1 - activar agente
        checkBox = WebDriverWait(ldriver, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@id="detailsData$Status"]')))
        checkBoxestado = checkBox.get_attribute('checked')
  
        #solo se activan los agentes que esten es estado desactivados
        if str(checkBoxestado)=="None":
             
            ldriver.execute_script("arguments[0].click();", checkBox)

            #Encuentra el elemento que contiene el shadow root utilizando el selector CSS
            planea_title = 'Planificación' if language == "es-MX" else 'Scheduling'
            planeacion0 = ldriver.find_element(By.CSS_SELECTOR, f"d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='{planea_title}']")

            #Obtiene el shadow root a apartir de la seccion colapsada de planeacion
            planeacion1 = ldriver.execute_script('return arguments[0].shadowRoot', planeacion0)

            # Encuentra y realiza una acción con el elemento dentro del shadow root (clic para abrir)
            expanplan = planeacion1.find_element(By.CSS_SELECTOR, "d2l-icon-custom[class='d2l-skeletize']")
            ldriver.execute_script("arguments[0].click();", expanplan)

            #field 2
            #activar fecha de inicio
            checkBox = WebDriverWait(ldriver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#schedulingData\$HasStartDate')))
            ldriver.execute_script("arguments[0].click();", checkBox)
            sleep(1)
            
            #Obtener fecha de inicio
            fini = ldriver.execute_script('return document.querySelector("d2l-input-date").shadowRoot.querySelector("d2l-input-text").shadowRoot.querySelector("input")')
        
            #field 3
            #activar fecha de final
            checkBox = WebDriverWait(ldriver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#schedulingData\$HasEndDate')))
            ldriver.execute_script("arguments[0].click();", checkBox)
            sleep(2)

            #fecha de final
            wait=WebDriverWait(ldriver, 10)
            root1 = wait.until(EC.presence_of_element_located((By.XPATH , '(//d2l-input-date[@id="schedulingData$EndDate"])[1]')))
            root11 = ldriver.execute_script('return arguments[0].shadowRoot', root1)

            root2 = root11.find_element(By.CLASS_NAME, "d2l-dropdown-opener")
            root22 = ldriver.execute_script('return arguments[0].shadowRoot', root2)
        
            #Obtener fecha de inicio
            ffin=root22.find_element(By.CSS_SELECTOR,"input[id^='d2l-uid']")

            #SE LIMPIA EL CAMPO DE FECHA FINAL Y INICIAL
            fini.clear()
            ffin.clear()

            # Convertir los fechas Timestamp a cadenas en el formato YYYY-MM-DD
            fecha_inicio = dlrow["FECHAINICIO"].strftime("%Y-%m-%d")
            fecha_final = dlrow["FECHAFINAL"].strftime("%Y-%m-%d")
    
            #SET FECHA INICIAL Y FINAL
            fini.send_keys(fecha_inicio)
            ffin.send_keys(fecha_final)
    
        #guardar y cerrar
        boton_texto = 'Guardar y Cerrar' if language == 'es-MX' else 'Save and Close'
        guardar_cerrar_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//button[contains(@class, 'd2l-button') and contains(text(), '{boton_texto}')]")
            ))
        
        guardar_cerrar_btn.click()
        sleep(5)    

    except TimeoutException:
         print("Un elemento no se encontró dentro del tiempo estimado.")


#**********
#Parametrizar los datos de los agentes: Frecuencia, repite y hora de ejecución
#**********
def CambiotipoAgen(ldriver, dlrow):
    # access the agent page
    agenteid="https://virtual.upb.edu.co/d2l/le/intelligentagents/agent/"\
        +str (dlrow["IDCURSO"])\
        +"/edit?agentId="\
        +str (dlrow["IDAGENTE"])
   
    ldriver.get(agenteid)
    sleep(1)
    
    #datos del agente
    #*********************************
    #selector de frecuencia - diario
    cfre = WebDriverWait(ldriver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#schedulingData\$ScheduleFrequency")))
    selerolobj=Select(cfre)
    selerolobj.select_by_value('Daily')
    #print("frecuencia")
    #input()

    #*********************************
    #repite cada 3 días  
    shadow0 = ldriver.find_element(By.ID, "schedulingData$NumberOfDays")
    root1 = ldriver.execute_script('return arguments[0].shadowRoot', shadow0)
    #print(shadow0.get_attribute("outerHTML"))
    #input()

    shadow1 = root1.find_element(By.CSS_SELECTOR, "d2l-input-text[id^='d2l-uid']")
    root2 = ldriver.execute_script('return arguments[0].shadowRoot', shadow1)
    #print(shadow1.get_attribute("outerHTML"))
    #input()

    repi = root2.find_element(By.CSS_SELECTOR,"input[id^='d2l-uid']")
    #print(repi.get_attribute("outerHTML"))
    #input()
    repi.clear()
    repi.send_keys("3")

    #*********************************
    #Hora programada 3 pm
    #repi = ldriver.execute_script('''return document.querySelector("d2l-input-number").shadowRoot.querySelector("d2l-input-text").shadowRoot.querySelector("input.d2l-input")''')
    shadow0 = ldriver.find_element(By.ID, "schedulingData$Time")
    root1 = ldriver.execute_script('return arguments[0].shadowRoot', shadow0)
    #print(shadow0.get_attribute("outerHTML"))

    horpro = root1.find_element(By.CSS_SELECTOR,"input[id^='d2l-uid']")
    #print(horpro.get_attribute("outerHTML")) - "15:00:00" 03:00 p.m. get_attribute('value'))

    #horpro.click()
    horpro.send_keys([Keys.BACKSPACE] * 13)
    #horpro.send_keys([Keys.DELETE])
    #horpro.send_keys('03:00 p.m.')
    ldriver.execute_script("arguments[0].value='03:00 p.m.';", horpro)
      
    #guardar y cerrar
    #button = ldriver.find_element(By.XPATH, '/html/body/div[2]/d2l-floating-buttons/button[1]')
    button = ldriver.find_element(By.XPATH, '/html[1]/body[1]/div[2]/d2l-floating-buttons[1]/button[1]')
    button.click()
    sleep(2)

#*************
# Función para acceder al Shadow Root
#*************
def expand_shadow_element(element, ldriver):
  shadow_root = ldriver.execute_script('return arguments[0].shadowRoot', element)
  return shadow_root

#**********
#cambia el agente Sin ingreso al curso - Docentes: Acciones Correo
#**********
def CambioAgenteDocente(ldriver, dlrow, NuevoHTML):
    # access the agent page
    agenteid="https://virtual.upb.edu.co/d2l/le/intelligentagents/agent/"\
        +str (dlrow["IDCURSO"])\
        +"/edit?agentId="\
        +str (dlrow["IDAGENTE"])
    
    ldriver.get(agenteid)
    sleep(1)

    # Obtiene el idioma de la página
    language = ldriver.find_element(By.CSS_SELECTOR, "html").get_attribute("lang")
    #print (language)
     
    #****************
    #Datos del agente
    #*****************
    #
    #abrir planeacion colapsado
    #
  
    #"""
    #Encuentra el elemento que contiene el shadow root utilizando el selector CSS
    if language=="es-MX":
        planeacion0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Planificación']")
    else:
        planeacion0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Scheduling']")

    #Obtiene el shadow root del host
    planeacion1 = ldriver.execute_script('return arguments[0].shadowRoot', planeacion0)

    # Encuentra y realiza una acción dentro del shadow root (clic para abrir)
    expanplan = planeacion1.find_element(By.CSS_SELECTOR, "d2l-icon-custom[class='d2l-skeletize']")
    ldriver.execute_script("arguments[0].click();", expanplan)
  
    #Repetir cada - Tienes dos Shadow
    #*********************************
    shadow0 = ldriver.find_element(By.ID, "schedulingData$NumberOfDays")
    root1 = ldriver.execute_script('return arguments[0].shadowRoot', shadow0)

    shadow1 = root1.find_element(By.CSS_SELECTOR, "d2l-input-text[id^='d2l-uid']")
    root2 = ldriver.execute_script('return arguments[0].shadowRoot', shadow1)

    repi = root2.find_element(By.CSS_SELECTOR,"input[id^='d2l-uid']")
    repi.clear()
    repi.send_keys("3")
    
    #
    #abrir acciones colapsado (Criterios)
    #
    #Encuentra el elemento que contiene el shadow root utilizando el selector CSS
    if language=="es-MX":
        criterios = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Criterios']")
    else:
        criterios = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Criteria']")

    #Obtiene el shadow root del host
    criterios1 = ldriver.execute_script('return arguments[0].shadowRoot', criterios)

    # Encuentra y realiza la accion dentro del shadow root (clic para abrir)
    expanacci = criterios1.find_element(By.CSS_SELECTOR, "d2l-icon-custom[class='d2l-skeletize']")
    ldriver.execute_script("arguments[0].click();", expanacci)

    #El usuario no accedió al curso por al menos - Tienes dos Shadow
    #*********************************
    shadow0 = ldriver.find_element(By.ID, "criteriaData$NoCourseActivityNumDays")
    root1 = ldriver.execute_script('return arguments[0].shadowRoot', shadow0)

    shadow1 = root1.find_element(By.CSS_SELECTOR, "d2l-input-text[id^='d2l-uid']")
    root2 = ldriver.execute_script('return arguments[0].shadowRoot', shadow1)

    NoAcc = root2.find_element(By.CSS_SELECTOR,"input[id^='d2l-uid']")
    
    NoAcc.clear()
    NoAcc.send_keys("3")
    #"""
 
    #""""
    #abrir acciones colapsado (Correo)
    #Encuentra el elemento que contiene el shadow root utilizando el selector CSS
    if language=="es-MX":
        acciones0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Acciones']")
    else:
        acciones0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Actions']")

    #Obtiene el shadow root del host
    acciones1 = ldriver.execute_script('return arguments[0].shadowRoot', acciones0)

    # Encuentra y realiza una acción con el elemento deseado dentro del shadow root (clic para abrir)
    expanacci = acciones1.find_element(By.CSS_SELECTOR, "d2l-icon-custom[class='d2l-skeletize']")
    ldriver.execute_script("arguments[0].click();", expanacci)

    #"""
    #cambiar datos del correo
    #Para:
    Para=ldriver.find_element(By.CSS_SELECTOR,"#actionsData\$To")
    Para.clear()
    Para.send_keys("soporte.upbvirtual@upb.edu.co")

    #CC:
    CC=ldriver.find_element(By.CSS_SELECTOR,"#actionsData\$Cc")
    CC.clear()
    #"""
 
    #cambiar datos del correo 
    #Asunto:
    Para=ldriver.find_element(By.XPATH,"(//input[@id='actionsData$EmailSubject'])[1]")
    Para.clear()
    Para.send_keys("Moderador sin ingresar al curso: {OrgUnitName}")
 
    #
    #Abrir codigo fuente HTML del editor **TIENE 3 Shadow DOM
    #

    #"""
    # Navegar al primer nivel del Shadow DOM
    root1_selector = "body > div:nth-child(11) > div:nth-child(1) > form:nth-child(3) > d2l-collapsible-panel:nth-child(4) > div:nth-child(4) > div:nth-child(6) > div:nth-child(1) > div:nth-child(1) > d2l-htmleditor:nth-child(2)"
    root1 = ldriver.find_element(By.CSS_SELECTOR, root1_selector)
    shadow_root1 = expand_shadow_element(root1, ldriver)
 
    # Navegar al segundo nivel del Shadow DOM
    root2_selector = "d2l-htmleditor-toolbar-full"
    root2 = shadow_root1.find_element(By.CSS_SELECTOR, root2_selector)
    shadow_root2 = expand_shadow_element(root2, ldriver)
 
    # Navegar al tercer nivel del Shadow DOM
    root3_selector = "d2l-htmleditor-button[icon='sourcecode']"
    root3 = shadow_root2.find_element(By.CSS_SELECTOR, root3_selector)
    shadow_root3 = expand_shadow_element(root3, ldriver)

    # Acceder al elemento abir codigo fuente y dar clic
    if language=="es-MX":
         CodigoFuente = shadow_root3.find_element(By.CSS_SELECTOR, "button[title='Código Fuente']")
    else:
         CodigoFuente = shadow_root3.find_element(By.CSS_SELECTOR, "button[title='Source Code']")
    
    ldriver.execute_script("arguments[0].click();", CodigoFuente)
    
    #
    #Gestionar codigo fuente HTML del editor **TIENE 2 Shadow DOM
    #
  
    # Navegar al primer nivel del Shadow DOM
    root1_selector = "body > div:nth-child(11) > div:nth-child(1) > form:nth-child(3) > d2l-collapsible-panel:nth-child(4) > div:nth-child(4) > div:nth-child(6) > div:nth-child(1) > div:nth-child(1) > d2l-htmleditor:nth-child(2)"
    root1 = ldriver.find_element(By.CSS_SELECTOR, root1_selector)
    shadow_root1 = expand_shadow_element(root1, ldriver)

    # Navegar al segundo nivel del Shadow DOM
    root2_selector = "d2l-htmleditor-sourcecode-dialog"
    root2 = shadow_root1.find_element(By.CSS_SELECTOR, root2_selector)
    shadow_root2 = expand_shadow_element(root2, ldriver)
  
    # Acceder al editor HTML
    EditorHtml = shadow_root2.find_element(By.CSS_SELECTOR, ".cm-content.cm-lineWrapping")

    # Copiar NuevoHTML al portapapeles sobre el editor HTML
    # Intentar copiar NuevoHTML al portapapeles
    for i in range(3):
        try:
            pyperclip.copy(NuevoHTML)
            break      # Romper el bucle si tiene éxito
        except pyperclip.PyperclipWindowsException:
            print("Error accediendo al portapapeles. Reintentando...")
            sleep(1)  # Esperar un segundo antes de volver a intentarlo

    # Emular las acciones de teclado
    actions = ActionChains(ldriver)

    # Seleccionar todo el contenido existente y borrarlo
    actions.move_to_element(EditorHtml)
    actions.click()
    actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)
    actions.send_keys(Keys.DELETE)
    #EditorHtml.send_keys([Keys.BACKSPACE] * 1700)
  
    # Pegar el nuevo HTML del portapapeles
    actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL)
    actions.perform()

    #
    #Guardar y cerrar el nuevo HTML **TIENE 2 Shadow DOM
    #
    # Acceder al primer Shadow DOM
    root1_selector = "body > div:nth-child(11) > div:nth-child(1) > form:nth-child(3) > d2l-collapsible-panel:nth-child(4) > div:nth-child(4) > div:nth-child(6) > div:nth-child(1) > div:nth-child(1) > d2l-htmleditor:nth-child(2)"
    root1 = ldriver.find_element(By.CSS_SELECTOR, root1_selector)
    shadow_root1 = expand_shadow_element(root1, ldriver)
   
    # Acceder al segundo Shadow DOM desde el primero
    root2_selector = "d2l-htmleditor-sourcecode-dialog"
    root2 = shadow_root1.find_element(By.CSS_SELECTOR, root2_selector)
    shadow_root2 = expand_shadow_element(root2, ldriver)
         
    # Encontrar el elemento botongrabarhtml y hacer clic en él
    cssbotongrabarhtml = "d2l-dialog:nth-child(1) > d2l-button:nth-child(2)"
    botongrabarhtml = shadow_root2.find_element(By.CSS_SELECTOR, cssbotongrabarhtml)
    botongrabarhtml.click()
    #"""

    #
    #Guardar y cerrar cambio del Agente docente
    #
    button = ldriver.find_element(By.XPATH, '/html[1]/body[1]/div[9]/d2l-floating-buttons[1]/button[1]')
    button.click()
    sleep(2)


#**********
#cambia el agente Sin ingreso ala plataforma - estudiantes
#**********
def CambioIngresoPlataforma(ldriver, dlrow, NuevoHTML):
    # access the agent page
    agenteid="https://virtual.upb.edu.co/d2l/le/intelligentagents/agent/"\
        +str (dlrow["IDCURSO"])\
        +"/edit?agentId="\
        +str (dlrow["IDAGENTE"])
    
    ldriver.get(agenteid)
    sleep(1)

    # Obtiene el idioma de la página
    language = ldriver.find_element(By.CSS_SELECTOR, "html").get_attribute("lang")
     
    #****************
    #Datos del agente
    #*****************
    #
    #abrir planeacion colapsado
    #

    #Encuentra el elemento que contiene el shadow root utilizando el selector CSS
    if language=="es-MX":
        planeacion0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Planificación']")
    else:
        planeacion0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Scheduling']")

    #Obtiene el shadow root del host
    planeacion1 = ldriver.execute_script('return arguments[0].shadowRoot', planeacion0)

    # Encuentra y realiza una acción dentro del shadow root (clic para abrir)
    expanplan = planeacion1.find_element(By.CSS_SELECTOR, "d2l-icon-custom[class='d2l-skeletize']")
    ldriver.execute_script("arguments[0].click();", expanplan)
    sleep(1)
  
    #Repetir cada - Tienes dos Shadow
    #*********************************
    shadow0 = ldriver.find_element(By.ID, "schedulingData$NumberOfDays")
    root1 = ldriver.execute_script('return arguments[0].shadowRoot', shadow0)

    shadow1 = root1.find_element(By.CSS_SELECTOR, "d2l-input-text[id^='d2l-uid']")
    root2 = ldriver.execute_script('return arguments[0].shadowRoot', shadow1)

    repi = root2.find_element(By.CSS_SELECTOR,"input[id^='d2l-uid']")
    
    repi.clear()
    repi.send_keys("3")
    
    #
    #abrir acciones colapsado (Criterios)
    #
    #Encuentra el elemento que contiene el shadow root utilizando el selector CSS
    if language=="es-MX":
        criterios = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Criterios']")
    else:
        criterios = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Criteria']")

    #Obtiene el shadow root del host
    criterios1 = ldriver.execute_script('return arguments[0].shadowRoot', criterios)

    # Encuentra y realiza la accion dentro del shadow root (clic para abrir)
    expanacci = criterios1.find_element(By.CSS_SELECTOR, "d2l-icon-custom[class='d2l-skeletize']")
    ldriver.execute_script("arguments[0].click();", expanacci)
    sleep(1)

    #El usuario no accedió al curso por al menos - Tienes dos Shadow
    #*********************************
    shadow0 = ldriver.find_element(By.ID, "criteriaData$NumberOfDaysNotLoggedIn") 
    root1 = ldriver.execute_script('return arguments[0].shadowRoot', shadow0)

    shadow1 = root1.find_element(By.CSS_SELECTOR, "d2l-input-text[id^='d2l-uid']")
    root2 = ldriver.execute_script('return arguments[0].shadowRoot', shadow1)

    NoAcc = root2.find_element(By.CSS_SELECTOR,"input[id^='d2l-uid']")
    NoAcc.clear()
    NoAcc.send_keys("3")
 
    #""""
    #abrir acciones colapsado (Correo)
    #Encuentra el elemento que contiene el shadow root utilizando el selector CSS
    if language=="es-MX":
        acciones0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Acciones']")
    else:
        acciones0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Actions']")

    #Obtiene el shadow root del host
    acciones1 = ldriver.execute_script('return arguments[0].shadowRoot', acciones0)

    # Encuentra y realiza una acción con el elemento deseado dentro del shadow root (clic para abrir)
    expanacci = acciones1.find_element(By.CSS_SELECTOR, "d2l-icon-custom[class='d2l-skeletize']")
    ldriver.execute_script("arguments[0].click();", expanacci)
    sleep(1)

    #Cambiar Repetición de acción: a Realizar una acción sólo la primera vez
    RepeAcc= ldriver.find_element(By.XPATH,"(//input[@id='actionsData$FirstTimeAction'])[1]")
    #RepeAcc.click()
    ldriver.execute_script("arguments[0].click();", RepeAcc)

    #cambiar datos del correo 
    #Asunto:
    Para=ldriver.find_element(By.XPATH,"(//input[@id='actionsData$EmailSubject'])[1]")
    Para.clear()
    Para.send_keys("¡Empieza tu aventura académica! Accede ahora al Campus Virtual de la UPB")

    #
    #Abrir codigo fuente HTML del editor **TIENE 3 Shadow DOM
    #

    # Navegar al primer nivel del Shadow DOM
    root1_selector = "body > div:nth-child(11) > div:nth-child(1) > form:nth-child(3) > d2l-collapsible-panel:nth-child(4) > div:nth-child(4) > div:nth-child(6) > div:nth-child(1) > div:nth-child(1) > d2l-htmleditor:nth-child(2)"
    root1 = ldriver.find_element(By.CSS_SELECTOR, root1_selector)
    shadow_root1 = expand_shadow_element(root1, ldriver)
 
    # Navegar al segundo nivel del Shadow DOM
    root2_selector = "d2l-htmleditor-toolbar-full"
    root2 = shadow_root1.find_element(By.CSS_SELECTOR, root2_selector)
    shadow_root2 = expand_shadow_element(root2, ldriver)
 
    # Navegar al tercer nivel del Shadow DOM
    root3_selector = "d2l-htmleditor-button[icon='sourcecode']"
    root3 = shadow_root2.find_element(By.CSS_SELECTOR, root3_selector)
    shadow_root3 = expand_shadow_element(root3, ldriver)

    # Acceder al elemento abir codigo fuente y dar clic
    if language=="es-MX":
         CodigoFuente = shadow_root3.find_element(By.CSS_SELECTOR, "button[title='Código Fuente']")
    else:
         CodigoFuente = shadow_root3.find_element(By.CSS_SELECTOR, "button[title='Source Code']")
    
    ldriver.execute_script("arguments[0].click();", CodigoFuente)
    sleep(1)
    
    #
    #Gestionar codigo fuente HTML del editor **TIENE 2 Shadow DOM
    #
  
    # Navegar al primer nivel del Shadow DOM
    root1_selector = "body > div:nth-child(11) > div:nth-child(1) > form:nth-child(3) > d2l-collapsible-panel:nth-child(4) > div:nth-child(4) > div:nth-child(6) > div:nth-child(1) > div:nth-child(1) > d2l-htmleditor:nth-child(2)"
    root1 = ldriver.find_element(By.CSS_SELECTOR, root1_selector)
    shadow_root1 = expand_shadow_element(root1, ldriver)

    # Navegar al segundo nivel del Shadow DOM
    root2_selector = "d2l-htmleditor-sourcecode-dialog"
    root2 = shadow_root1.find_element(By.CSS_SELECTOR, root2_selector)
    shadow_root2 = expand_shadow_element(root2, ldriver)
  
    # Acceder al editor HTML
    EditorHtml = shadow_root2.find_element(By.CSS_SELECTOR, ".cm-content.cm-lineWrapping")

    # Copiar NuevoHTML al portapapeles sobre el editor HTML
    # Intentar copiar NuevoHTML al portapapeles 3 veces **para evadir error de acceso al portapales
    for i in range(3):
        try:
            pyperclip.copy(NuevoHTML)
            break      # Romper el bucle si tiene éxito
        except pyperclip.PyperclipWindowsException:
            print("Error accediendo al portapapeles. Reintentando...")
            sleep(1)  # Esperar un segundo antes de volver a intentarlo

    # Emular las acciones de teclado
    actions = ActionChains(ldriver)

    # Seleccionar todo el contenido existente y borrarlo
    actions.move_to_element(EditorHtml)
    actions.click()
    actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)
    actions.send_keys(Keys.DELETE)
  
    # Pegar el nuevo HTML del portapapeles
    actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL)
    actions.perform()

    #
    #Guardar y cerrar el nuevo HTML **TIENE 2 Shadow DOM
    #
    # Acceder al primer Shadow DOM
    root1_selector = "body > div:nth-child(11) > div:nth-child(1) > form:nth-child(3) > d2l-collapsible-panel:nth-child(4) > div:nth-child(4) > div:nth-child(6) > div:nth-child(1) > div:nth-child(1) > d2l-htmleditor:nth-child(2)"
    root1 = ldriver.find_element(By.CSS_SELECTOR, root1_selector)
    shadow_root1 = expand_shadow_element(root1, ldriver)
   
    # Acceder al segundo Shadow DOM desde el primero
    root2_selector = "d2l-htmleditor-sourcecode-dialog"
    root2 = shadow_root1.find_element(By.CSS_SELECTOR, root2_selector)
    shadow_root2 = expand_shadow_element(root2, ldriver)
         
    # Encontrar el elemento botongrabarhtml y hacer clic en él
    cssbotongrabarhtml = "d2l-dialog:nth-child(1) > d2l-button:nth-child(2)"
    botongrabarhtml = shadow_root2.find_element(By.CSS_SELECTOR, cssbotongrabarhtml)
    botongrabarhtml.click()
    #"""

    #
    #Guardar y cerrar cambio del Agente docente
    #
    button = ldriver.find_element(By.XPATH, '/html[1]/body[1]/div[9]/d2l-floating-buttons[1]/button[1]')
    button.click()
    sleep(2)

#**********
#cambia el agente Sin ingreso al curso - estudiantes
#**********
def CambioIngresoAlCurso(ldriver, dlrow, NuevoHTML):
    # access the agent page
    agenteid="https://virtual.upb.edu.co/d2l/le/intelligentagents/agent/"\
        +str (dlrow["IDCURSO"])\
        +"/edit?agentId="\
        +str (dlrow["IDAGENTE"])
    
    ldriver.get(agenteid)
    sleep(1)

    # Obtiene el idioma de la página
    language = ldriver.find_element(By.CSS_SELECTOR, "html").get_attribute("lang")
     
    #****************
    #Datos del agente
    #*****************
    #
    #abrir planeacion colapsado
    #

    #Encuentra el elemento que contiene el shadow root utilizando el selector CSS
    if language=="es-MX":
        planeacion0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Planificación']")
    else:
        planeacion0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Scheduling']")

    #Obtiene el shadow root del host
    planeacion1 = ldriver.execute_script('return arguments[0].shadowRoot', planeacion0)
    sleep(1)

    # Encuentra y realiza una acción dentro del shadow root (clic para abrir)
    expanplan = planeacion1.find_element(By.CSS_SELECTOR, "d2l-icon-custom[class='d2l-skeletize']")
    ldriver.execute_script("arguments[0].click();", expanplan)
    sleep(1)
  
    #Repetir cada - Tienes dos Shadow
    #*********************************
    shadow0 = ldriver.find_element(By.ID, "schedulingData$NumberOfDays")
    root1 = ldriver.execute_script('return arguments[0].shadowRoot', shadow0)

    shadow1 = root1.find_element(By.CSS_SELECTOR, "d2l-input-text[id^='d2l-uid']")
    root2 = ldriver.execute_script('return arguments[0].shadowRoot', shadow1)

    repi = root2.find_element(By.CSS_SELECTOR,"input[id^='d2l-uid']")
    repi.clear()
    repi.send_keys("3")
    
    #
    #abrir acciones colapsado (Criterios)
    #
    #Encuentra el elemento que contiene el shadow root utilizando el selector CSS
    if language=="es-MX":
        criterios = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Criterios']")
    else:
        criterios = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Criteria']")

    #Obtiene el shadow root del host
    criterios1 = ldriver.execute_script('return arguments[0].shadowRoot', criterios)
    sleep(1)

    # Encuentra y realiza la accion dentro del shadow root (clic para abrir)
    expanacci = criterios1.find_element(By.CSS_SELECTOR, "d2l-icon-custom[class='d2l-skeletize']")
    ldriver.execute_script("arguments[0].click();", expanacci)
    sleep(1)

    #
    #Validar Tomar medidas sobre la actividad: Actividad del curso este activo
    #
    checkBox = WebDriverWait(ldriver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@id='criteriaData$CourseActivity']")))
    #print("**chek 1*")
    estado = checkBox.get_attribute('checked')
    #print(estado)

    #solo se activa para registrar los dias usaurio accedio al cursso
    if str(estado)=="None":
        ldriver.execute_script("arguments[0].click();", checkBox)
  
    #El usuario no accedió al curso por al menos - Tienes dos Shadow
    #*********************************
    shadow0 = ldriver.find_element(By.ID, "criteriaData$NoCourseActivityNumDays") 
    root1 = ldriver.execute_script('return arguments[0].shadowRoot', shadow0)

    shadow1 = root1.find_element(By.CSS_SELECTOR, "d2l-input-text[id^='d2l-uid']")
    root2 = ldriver.execute_script('return arguments[0].shadowRoot', shadow1)

    NoAcc = root2.find_element(By.CSS_SELECTOR,"input[id^='d2l-uid']")
    NoAcc.clear()
    NoAcc.send_keys("3")
 
    #abrir acciones colapsado (Correo)
    #Encuentra el elemento que contiene el shadow root utilizando el selector CSS
    if language=="es-MX":
        acciones0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Acciones']")
    else:
        acciones0 = ldriver.find_element(By.CSS_SELECTOR, "d2l-collapsible-panel[type='default'][class='d2l-collapsible-panel'][panel-title='Actions']")

    #Obtiene el shadow root del host
    acciones1 = ldriver.execute_script('return arguments[0].shadowRoot', acciones0)
    sleep(1)

    # Encuentra y realiza una acción con el elemento deseado dentro del shadow root (clic para abrir)
    expanacci = acciones1.find_element(By.CSS_SELECTOR, "d2l-icon-custom[class='d2l-skeletize']")
    ldriver.execute_script("arguments[0].click();", expanacci)
    sleep(1)

    #Cambiar Repetición de acción: a Realizar una acción cada vez que el agente sea evaluado y haya concordancia 
    RepeAcc= ldriver.find_element(By.XPATH,"(//input[@id='actionsData$EveryTimeAction'])[1]")
    ldriver.execute_script("arguments[0].click();", RepeAcc)

    #cambiar datos del correo 
    #Asunto:
    Para=ldriver.find_element(By.XPATH,"(//input[@id='actionsData$EmailSubject'])[1]")
    Para.clear()
    Para.send_keys("¡No pierdas el ritmo! Vuelve a tu curso: {OrgUnitName}")

    #
    #Abrir codigo fuente HTML del editor **TIENE 3 Shadow DOM
    #

    # Navegar al primer nivel del Shadow DOM
    root1_selector = "body > div:nth-child(11) > div:nth-child(1) > form:nth-child(3) > d2l-collapsible-panel:nth-child(4) > div:nth-child(4) > div:nth-child(6) > div:nth-child(1) > div:nth-child(1) > d2l-htmleditor:nth-child(2)"
    root1 = ldriver.find_element(By.CSS_SELECTOR, root1_selector)
    shadow_root1 = expand_shadow_element(root1, ldriver)
 
    # Navegar al segundo nivel del Shadow DOM
    root2_selector = "d2l-htmleditor-toolbar-full"
    root2 = shadow_root1.find_element(By.CSS_SELECTOR, root2_selector)
    shadow_root2 = expand_shadow_element(root2, ldriver)
 
    # Navegar al tercer nivel del Shadow DOM
    root3_selector = "d2l-htmleditor-button[icon='sourcecode']"
    root3 = shadow_root2.find_element(By.CSS_SELECTOR, root3_selector)
    shadow_root3 = expand_shadow_element(root3, ldriver)

    # Acceder al elemento abir codigo fuente y dar clic
    if language=="es-MX":
         CodigoFuente = shadow_root3.find_element(By.CSS_SELECTOR, "button[title='Código Fuente']")
    else:
         CodigoFuente = shadow_root3.find_element(By.CSS_SELECTOR, "button[title='Source Code']")
    
    ldriver.execute_script("arguments[0].click();", CodigoFuente)
    sleep(1)
    
    #
    #Gestionar codigo fuente HTML del editor **TIENE 2 Shadow DOM
    #
  
    # Navegar al primer nivel del Shadow DOM
    root1_selector = "body > div:nth-child(11) > div:nth-child(1) > form:nth-child(3) > d2l-collapsible-panel:nth-child(4) > div:nth-child(4) > div:nth-child(6) > div:nth-child(1) > div:nth-child(1) > d2l-htmleditor:nth-child(2)"
    root1 = ldriver.find_element(By.CSS_SELECTOR, root1_selector)
    shadow_root1 = expand_shadow_element(root1, ldriver)

    # Navegar al segundo nivel del Shadow DOM
    root2_selector = "d2l-htmleditor-sourcecode-dialog"
    root2 = shadow_root1.find_element(By.CSS_SELECTOR, root2_selector)
    shadow_root2 = expand_shadow_element(root2, ldriver)
  
    # Acceder al editor HTML
    EditorHtml = shadow_root2.find_element(By.CSS_SELECTOR, ".cm-content.cm-lineWrapping")

    # Copiar NuevoHTML al portapapeles sobre el editor HTML
    # Intentar copiar NuevoHTML al portapapeles 3 veces **para evadir error de acceso al portapales
    for i in range(3):
        try:
            pyperclip.copy(NuevoHTML)
            break      # Romper el bucle si tiene éxito
        except pyperclip.PyperclipWindowsException:
            print("Error accediendo al portapapeles. Reintentando...")
            sleep(1)  # Esperar un segundo antes de volver a intentarlo

    # Emular las acciones de teclado
    actions = ActionChains(ldriver)

    # Seleccionar todo el contenido existente y borrarlo
    actions.move_to_element(EditorHtml)
    actions.click()
    actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)
    actions.send_keys(Keys.DELETE)
  
    # Pegar el nuevo HTML del portapapeles
    actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL)
    actions.perform()

    #
    #Guardar y cerrar el nuevo HTML **TIENE 2 Shadow DOM
    #
    # Acceder al primer Shadow DOM
    root1_selector = "body > div:nth-child(11) > div:nth-child(1) > form:nth-child(3) > d2l-collapsible-panel:nth-child(4) > div:nth-child(4) > div:nth-child(6) > div:nth-child(1) > div:nth-child(1) > d2l-htmleditor:nth-child(2)"
    root1 = ldriver.find_element(By.CSS_SELECTOR, root1_selector)
    shadow_root1 = expand_shadow_element(root1, ldriver)
   
    # Acceder al segundo Shadow DOM desde el primero
    root2_selector = "d2l-htmleditor-sourcecode-dialog"
    root2 = shadow_root1.find_element(By.CSS_SELECTOR, root2_selector)
    shadow_root2 = expand_shadow_element(root2, ldriver)
         
    # Encontrar el elemento botongrabarhtml y hacer clic en él
    cssbotongrabarhtml = "d2l-dialog:nth-child(1) > d2l-button:nth-child(2)"
    botongrabarhtml = shadow_root2.find_element(By.CSS_SELECTOR, cssbotongrabarhtml)
    botongrabarhtml.click()
 
    #
    #Guardar y cerrar cambio del Agente docente
    #
    button = ldriver.find_element(By.XPATH, '/html[1]/body[1]/div[9]/d2l-floating-buttons[1]/button[1]')
    button.click()
    sleep(2)