import os,gc
import pandas as pd
from getpass import getpass
import mod_classadd as classadd # Asegúrar que classadd.py esté en el mismo directorio o en PYTHONPATH

#obtener credenciales
def obtener_credenciales():
    """Captura de credenciales del usuario de forma segura."""
    usuario = input("Ingresar usuario: ")
    contrasena = getpass("Ingresar contraseña: ")
    clave_2fa = input("Clave 2FA: ")
    return usuario, contrasena, clave_2fa

#cargar datos desde un archivo Excel
def cargar_datos(ruta_archivo):
    """Carga y valida los datos desde un archivo Excel."""
    try:
        datos = pd.read_excel(ruta_archivo)
        if classadd.check_user_file(datos):
            return datos
        else:
            raise ValueError("El archivo proporcionado no cumple con los requisitos esperados.")
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        exit(1)

#gestionar agentes - activa los agentes, asigna fecha: inicio y fin
def gestionar_agentes(driver, agentes):
    """Itera sobre los agentes y gestiona los datos en la plataforma."""
    total_agentes = len(agentes)

    for index, agente in agentes.iterrows():
        porcentaje = ((index + 1) / total_agentes) * 100
        print(f"\nGestionando agente para el ID curso {agente['IDCURSO']} -> {index + 1} de {total_agentes} ({porcentaje:.2f}%)")

        classadd.setdataAgen(driver, agente)
   
# Main function to execute the script
def main():
    ruta_archivo = "agentes.xlsx"

    usuario, contrasena, clave_2fa = obtener_credenciales()

    driver = classadd.crear_WebDriver()

    classadd.login(driver, usuario, contrasena, clave_2fa)
    
    agentes = cargar_datos(ruta_archivo)
    print("\n-------------------------------")  
    print("Agentes a gestionar")
    print(agentes)

    gestionar_agentes(driver, agentes)

    print("\nTerminando la ejecución del Bot...")
    driver.quit()
 
    # Eliminar la variable de ruta del archivo para evitar confusiones
    del ruta_archivo
    gc.collect()  # Recolectar basura para liberar memoria


if __name__ == "__main__":
    main()