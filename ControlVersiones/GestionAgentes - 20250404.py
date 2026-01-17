import py_compile
import pandas as pd
import classadd
from getpass import getpass

def main():
    # captura de credenciales
    usr = input ("Ingresar usuario: ")
    pwd = getpass ("Ingresar contraseña: ")
    dfa = input("Clave 2FA: ")

    #get the file path
    path="agentes.xlsx"

    # mostrar datos
    #print ("Usuario ->"+usr)
    #print ("contraseña ->"+pwd)

    # crear instancia chrome driver
    driver=classadd.crear_WebDriver()

    # login con las credenciales
    classadd.login(driver, usr, pwd, dfa)

    #Cargar los datos del archivo
    agentes=pd.read_excel(path)
    print (agentes)

    #validar los datos del excel
    agentesFlag = classadd.check_user_file(agentes)
    tam = len(agentes)
    #Gestionar fechas de los agentes
    for index, drow in agentes.iterrows():
        percentage = ((index + 1) / tam) * 100
        print(
            "Gestionando agente para el ID curso ", str(drow["IDCURSO"]), " -> ",
            index + 1,
            "de",
            len(agentes),
            "(" + str(round(percentage,2)) + "%)",
        )
        classadd.setdataAgen(driver, drow)
    
    print("Terminando la ejecución del programa...")
    driver.close()

if __name__ == "__main__": 
    main()