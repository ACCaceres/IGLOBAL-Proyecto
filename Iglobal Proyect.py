#Proyecto para Generar documento electronicos al SII
#Boletas-Facturas con IGlobal y pyautogui

# Librerias
import pyautogui
import time
import pandas as pd
import pygetwindow as gw
import os
import keyboard
from datetime import datetime

# Variable global para controlar la ejecución del script
detener = False

# Función para detener la ejecución al presionar Escape
def detener_ejecucion():
    global detener
    detener = True
    print("Ejecución detenida por el usuario.")

# Asigna la tecla Escape a la función detener_ejecucion
keyboard.add_hotkey('esc', detener_ejecucion)

# Carga de datos desde un archivo Excel o CSV
data_file = "C:/Users/areca/Desktop/IGLOBAL Proyecto/Pagos.xlsx"
sheet_name = "Hoja1"  # Cambia esto si tu archivo tiene múltiples hojas
datos = pd.read_excel(data_file, sheet_name=sheet_name)

def esperar_segundos(segundos):
    time.sleep(segundos)

def preparar_ventana_iglobal():
    ventanas = [v for v in gw.getAllTitles() if "iGlobal ERP" in v]
    if ventanas:
        ventana = gw.getWindowsWithTitle(ventanas[0])[0]
        ventana.activate()
        ventana.maximize()
    else:
        print("No se encontró la ventana de iGlobal. Asegúrate de que esté abierta y con el título correcto.")
        exit()

def procesar_documento(fila):
    global detener
    if detener: return

    pyautogui.click(x=189, y=69)
    if detener: return
    esperar_segundos(0.5)

    pyautogui.write(str(fila['Codigo']))
    pyautogui.press("enter")
    if detener: return
    esperar_segundos(0.5)

    pyautogui.click(x=1066, y=604)
    if detener: return
    esperar_segundos(0.5)

    pyautogui.click(x=1237, y=257)
    pyautogui.write(str(fila['Precio']))
    pyautogui.press("tab")
    if detener: return
    esperar_segundos(0.3)

    pyautogui.write(str(fila['Cantidad']))
    pyautogui.press("tab")
    if detener: return
    esperar_segundos(0.3)

    pyautogui.click(x=1222, y=328)
    pyautogui.write(fila['Descripcion'])
    pyautogui.press("tab")
    if detener: return
    esperar_segundos(0.3)

    pyautogui.click(x=1554, y=336)
    if detener: return
    esperar_segundos(0.5)

    tipo_documento = fila['TipoDocumento']
    pyautogui.doubleClick(x=1315, y=403)
    if tipo_documento == "Cotizacion":
        pyautogui.press("down", presses=2, interval=0.3)
    elif tipo_documento == "Factura":
        pyautogui.press("down", presses=3, interval=0.3)
    elif tipo_documento == "NCredito":
        pyautogui.press("down", presses=6, interval=0.3)
    if detener: return
    esperar_segundos(0.5)

    pyautogui.click(x=1356, y=516)
    if detener: return
    esperar_segundos(0.6)
    pyautogui.write(str(fila['ClienteNombre']))
    pyautogui.press("enter")
    if detener: return
    esperar_segundos(1)

    pyautogui.doubleClick(x=335, y=255)
    if detener: return
    esperar_segundos(0.5)

    pyautogui.doubleClick(x=1324, y=647)
    pyautogui.write(fila['Descripcion'])
    if detener: return
    esperar_segundos(0.5)

    tipo_pago = fila['MedioPago']
    pyautogui.doubleClick(x=1320, y=479)
    if tipo_pago == "Transferencia":
        pyautogui.press("down", presses=4, interval=0.3)
    elif tipo_pago == "Cheque":
        pyautogui.press("up", presses=2, interval=0.3)
    if detener: return
    esperar_segundos(0.5)

    pyautogui.press("f8")
    if detener: return
    esperar_segundos(1)

    pyautogui.click(x=1031, y=602)
    if detener: return
    esperar_segundos(1)

    pyautogui.click(x=544, y=596)
    cliente_nombre = str(fila['ClienteNombre'])[:240] if len(str(fila['ClienteNombre'])) > 240 else str(fila['ClienteNombre'])
    nombre_archivo = f"{cliente_nombre}.pdf"
    pyautogui.write(nombre_archivo)
    if detener: return
    pyautogui.press("enter")
    esperar_segundos(1)

def automatizar_proceso():
    global detener

    # Inicio de tiempo de ejecución
    inicio = datetime.now()

    preparar_ventana_iglobal()
    for index, fila in datos.iterrows():
        if detener:
            print("Ejecución detenida.")
            break
        procesar_documento(fila)
        print(f"Documento {index + 1} procesado correctamente.")

    # Fin de tiempo de ejecución
    fin = datetime.now()
    duracion = fin - inicio
    print(f"Ejecución completada en: {duracion}")

if __name__ == "__main__":
    automatizar_proceso()
