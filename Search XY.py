#Buscas las coordenas X, Y para una posicion en el mouse
import pyautogui
import time

print("Mueve el cursor a la posición deseada. Las coordenadas se mostrarán cada segundo.\nPresiona Ctrl + C para detener.")
try:
    while True:
        x, y = pyautogui.position()
        print(f"Posición actual: X={x}, Y={y}")
        time.sleep(1)
except KeyboardInterrupt:
    print("\nFinalizado.")
