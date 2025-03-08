# Este archivo se subirá a GitHub
import pyautogui
import time

def mover_mouse():
    """Función simple para mover el mouse en un patrón cuadrado"""
    print("Iniciando prueba de pyautogui...")
    
    # Guardar posición original
    pos_original = pyautogui.position()
    
    # Mover en forma de cuadrado
    for _ in range(2):  # Hacer el cuadrado dos veces
        pyautogui.moveTo(pos_original.x + 100, pos_original.y, duration=1)
        pyautogui.moveTo(pos_original.x + 100, pos_original.y + 100, duration=1)
        pyautogui.moveTo(pos_original.x, pos_original.y + 100, duration=1)
        pyautogui.moveTo(pos_original.x, pos_original.y, duration=1)
    
    # Mostrar un mensaje
    pyautogui.alert(text="Prueba completada con éxito!", title="Código remoto funcionando", button="OK")
    print("Prueba finalizada.")

def main():
    print("Código cargado exitosamente desde GitHub")
    time.sleep(1)
    print("Esperando 3 segundos antes de ejecutar la prueba...")
    time.sleep(3)
    mover_mouse()

# Este if no se ejecutará cuando se importe como módulo
if __name__ == "__main__":
    main()