import time
from tkinter import messagebox
import tkinter as tk
import os
import pyautogui
import sys
import pandas as pd
import uuid
from pathlib import Path
import json
import requests
import datetime


def cargar_configuracion():
    # Construir la ruta completa al archivo de configuración
    ruta_configuracion = os.path.join(os.path.expanduser(
        "~"), "Desktop", "Automatizar SUP", "configuracion.json")

    # Cargar la configuración desde el archivo JSON
    with open(ruta_configuracion, 'r') as archivo:
        configuracion = json.load(archivo)

    # Configuración del proxy
    configuracion_proxy = configuracion.get('proxy', {})
    os.environ['HTTP_PROXY'] = configuracion_proxy.get('HTTP_PROXY', '')
    os.environ['HTTPS_PROXY'] = configuracion_proxy.get('HTTPS_PROXY', '')

    # Configuración de la aplicación (agrega aquí todas las configuraciones que necesitas)
    configuracion_aplicacion = configuracion.get(
        'configuracion_aplicacion', {})

    return configuracion_aplicacion


# Cargar toda la configuración en una variable
config = cargar_configuracion()
# Acceso a las configuraciones específicas
PERIODO_OCASIONAL = config.get('periodo_ocasional')
HABER_OCASIONAL = config.get('haber_ocasional')
CODIGO_AIRHSP_OCASIONAL = config.get('codigo_airhsp_ocasional')
NOMBRE_COLUMNAS_REGISTRAR_OCASIONALES = ["orden", "modular", "haber", "monto"]
NOMBRE_COLUMNAS_REGISTRAR_REINTEGROS = [
    "orden", "modular", "monto92", "monto99", "leyenda_mensual", "leyenda_permanente", "periodo"]
NOMBRE_COLUMNAS_REGISTRAR_CONTRATOS = [
    "orden",
    "dni",
    "modular",
    "cod_plaza",
    "rdl",
    "cod_regimen",
    "tipo_afp",
    "cuspp",
    "fecha_afil",
    "fecha_devengue",
    "nivel_magisterial",
    "dias_laborados",
    "cuenta",
    "airhsp",
    "leyenda_mensual",
    "leyenda_permanente"
]
NOMBRE_COLUMNAS_REGISTRAR_NOMBRADOS = [
    "orden",
    "dni",
    "modular",
    "cod_plaza",
    "rdl",
    "fecha_inicio",
    "fecha_termino",
    "cod_regimen",
    "tipo_afp",
    "cuspp",
    "fecha_afil",
    "fecha_devengue", 
    "nec",
    "clave8",
    "cuenta",
    "airhsp",
    "leyenda_mensual",
    "leyenda_permanente"
]
TIEMPO_INICIAR_PROGRAMA = config.get('tiempo_iniciar_programa')
TIEMPO_ABRIR_NUEVO = config.get('tiempo_abrir_nuevo')
TIEMPO_ABRIR_HABERES = config.get('tiempo_abrir_haberes')
TIEMPO_PASAR_CAMPOS = 0.2  # config.get('tiempo_pasar_campos')
TIEMPO_MAXIMO_CARGA_NEXUS = config.get('tiempo_maximo_carga_nexus')
INTENTOS = config.get('intentos_buscar_imagen')
DURACION_CLIC = config.get('duracion_clic_imagen')
PREDETERMINADO = config.get('opcion_predeterminado')

columnas_ocasionales = '|'.join(NOMBRE_COLUMNAS_REGISTRAR_OCASIONALES)
columnas_reintegros = '|'.join(NOMBRE_COLUMNAS_REGISTRAR_REINTEGROS)
columnas_contratados = '|'.join(NOMBRE_COLUMNAS_REGISTRAR_CONTRATOS)
columnas_nombrados = '|'.join(NOMBRE_COLUMNAS_REGISTRAR_NOMBRADOS)


def obtener_mac():
    try:
        direccion_mac = ':'.join(['{:02x}'.format(
            (uuid.getnode() >> elements) & 0xff) for elements in range(5, -1, -1)])
        return direccion_mac
    except Exception as e:
        print(f"Error al obtener la dirección MAC: {e}")
        return None


def validar_cantidad_registros(log_file_path, cantidad_minima_registros):
    try:
        # Obtén la cantidad de líneas en el archivo de registro
        with open(log_file_path, "r") as log_file:
            line_count = sum(1 for line in log_file)

        return line_count < cantidad_minima_registros
    except FileNotFoundError:
        # Si el archivo no existe, también consideramos que no se ha superado la cantidad
        return True


def imprimir_encabezado(columnas_unidas, anchura_maxima=100):
    # Divide la cadena de columnas en palabras (separadas por el delimitador)
    columnas = columnas_unidas.split('|')

    # Prepara variables para acumular la línea actual y las líneas finales
    linea_actual = ""
    lineas_finales = []

    for columna in columnas:
        # Verificar si agregar la próxima columna excede la anchura máxima
        if len(linea_actual) + len(columna) + 1 > anchura_maxima:  # +1 por el delimitador
            # Agrega la línea actual a las líneas finales
            lineas_finales.append(linea_actual)
            # Comienza una nueva línea con la columna actual
            linea_actual = columna
        else:
            # Agrega la columna a la línea actual, con manejo para el inicio sin delimitador
            linea_actual = '|'.join(
                [linea_actual, columna]) if linea_actual else columna

    # Asegúrate de agregar la última línea acumulada
    lineas_finales.append(linea_actual)

    # Imprime cada línea con sus guiones
    for linea in lineas_finales:
        linea_guiones = '-' * len(linea)
        print("\nAntes de Iniciar verifique que todas las columnas necesarias se encuentren en el archivo CSV.")
        print("Columnas necesarias:\n")
        print(linea_guiones)
        print(linea)
        print(linea_guiones)


# Obtén la dirección MAC del equipo actual
mac_actual = obtener_mac()


# URL del archivo JSON en GitHub
github_json_url = (
    "https://raw.githubusercontent.com/edwinwmendez/Automatizar-SUP/main/Usuarios.json"
)

# Realiza la solicitud HTTP para obtener el contenido del JSON
response = requests.get(github_json_url)


TIEMPO_INICIO = 0.2
TIEMPO_PRESIONAR_BOTON = 0

# Obtener el directorio del escritorio y construir la ruta de la carpeta "Automatizar SUP"
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
automatizar_sup_path = os.path.join(desktop_path, "Automatizar SUP")

# Asegurar que la carpeta "Automatizar SUP" exista, si no, crearla
if not os.path.exists(automatizar_sup_path):
    os.makedirs(automatizar_sup_path)

# Construir la ruta de la carpeta de imágenes
IMG_PATH = os.path.join(automatizar_sup_path, "imagenes")

# Ruta del archivo CSV
CSV_FILE_REGISTRAR_OCASIONALES = os.path.join(
    automatizar_sup_path, "bd_ocasionales.csv")
CSV_FILE_REGISTRAR_CONTRATOS = os.path.join(
    automatizar_sup_path, "bd_contratos.csv")
CSV_FILE_REGISTRAR_NOMBRADOS = os.path.join(
    automatizar_sup_path, "bd_nombrados.csv")
CSV_FILE_REGISTRAR_REINTEGROS = os.path.join(
    automatizar_sup_path, "bd_reintegros.csv")

# Rutas de las imágenes con nombres diferentes
SELECCIONAR_REGISTRO_OCASIONAL = [
    os.path.join(IMG_PATH, "seleccionar_registro_ocasional.png"),
    os.path.join(IMG_PATH, "seleccionar_registro_ocasional2.png"),
    os.path.join(IMG_PATH, "seleccionar_registro_ocasional3.png")
]

SELECCIONAR_REGISTRO_HABILITADO = [
    os.path.join(IMG_PATH, "seleccionar_registro_habilitado.png"),
    os.path.join(IMG_PATH, "seleccionar_registro_habilitado2.png"),
    os.path.join(IMG_PATH, "seleccionar_registro_habilitado3.png")
]

CODIGO_MODULAR = [
    os.path.join(IMG_PATH, "codigo_modular.png"),
    os.path.join(IMG_PATH, "codigo_modular2.png"),
    os.path.join(IMG_PATH, "codigo_modular3.png")
]

CODIGO_MODULAR_NOMBRADO = [
    os.path.join(IMG_PATH, "codigo_modular_nombrado.png"),
    os.path.join(IMG_PATH, "codigo_modular_nombrado2.png"),
    os.path.join(IMG_PATH, "codigo_modular_nombrado3.png")
]


DNI = [
    os.path.join(IMG_PATH, "dni.png"),
    os.path.join(IMG_PATH, "dni2.png"),
    os.path.join(IMG_PATH, "dni3.png")
]

LEYENDA_MENSUAL = [
    os.path.join(IMG_PATH, "leyenda_mensual.png"),
    os.path.join(IMG_PATH, "leyenda_mensual2.png"),
    os.path.join(IMG_PATH, "leyenda_mensual3.png")
]

LEYENDA_PERMANENTE = [
    os.path.join(IMG_PATH, "leyenda_permanente.png"),
    os.path.join(IMG_PATH, "leyenda_permanente2.png"),
    os.path.join(IMG_PATH, "leyenda_permanente3.png")
]

DOCUMENTO_REFERENCIA = [
    os.path.join(IMG_PATH, "documento_referencia.png"),
    os.path.join(IMG_PATH, "documento_referencia2.png"),
    os.path.join(IMG_PATH, "documento_referencia3.png"),
]

ERROR_PADRON_NEXUS = os.path.join(
    IMG_PATH, "error_padron_nexus_contratado.png")

ERROR_PADRON_NEXUS_NOMBRADO = os.path.join(
    IMG_PATH, "error_padron_nexus_nombrado.png")

ERROR_PADRON_NEXUS_NOMBRADO_MODULAR = os.path.join(
    IMG_PATH, "error_padron_nexus_nombrado_modular.png")
CARGA_NEXUS_EXITOSA = os.path.join(IMG_PATH, "carga_nexus_exitosa.png")

CODIGO_FISCAL = [
    os.path.join(IMG_PATH, "codigo_fiscal.png"),
    os.path.join(IMG_PATH, "codigo_fiscal2.png"),
    os.path.join(IMG_PATH, "codigo_fiscal3.png"),
]

NIVEL_MAGISTERIAL = [
    os.path.join(IMG_PATH, "nivel_magisterial.png"),
    os.path.join(IMG_PATH, "nivel_magisterial2.png"),
    os.path.join(IMG_PATH, "nivel_magisterial3.png"),
]

DIAS_LABORADOS = [
    os.path.join(IMG_PATH, "dias_laborados.png"),
    os.path.join(IMG_PATH, "dias_laborados2.png"),
    os.path.join(IMG_PATH, "dias_laborados3.png"),
]
MODO_PAGO = [
    os.path.join(IMG_PATH, "modo_pago.png"),
    os.path.join(IMG_PATH, "modo_pago2.png"),
    os.path.join(IMG_PATH, "modo_pago3.png"),
]

CODIGO_NEXUS = [
    os.path.join(IMG_PATH, "codigo_nexus.png"),
    os.path.join(IMG_PATH, "codigo_nexus2.png"),
    os.path.join(IMG_PATH, "codigo_nexus3.png")
]

CODIGO_AIRHSP = [
    os.path.join(IMG_PATH, "airhsp.png"),
    os.path.join(IMG_PATH, "airhsp2.png"),
    os.path.join(IMG_PATH, "airhsp3.png")
]

NEC = [
    os.path.join(IMG_PATH, "nec.png"),
    os.path.join(IMG_PATH, "nec2.png"),
    os.path.join(IMG_PATH, "nec3.png")
]

ESTABLECIMIENTO = [
    os.path.join(IMG_PATH, "establecimiento.png"),
    os.path.join(IMG_PATH, "establecimiento2.png"),
    os.path.join(IMG_PATH, "establecimiento3.png")
]


IMG_BTN_BUSCAR = [
    os.path.join(IMG_PATH, "btn_buscar.png"),
    os.path.join(IMG_PATH, "btn_buscar2.png"),
    os.path.join(IMG_PATH, "btn_buscar3.png")
]

IMG_BTN_EMPLEADOS_NUEVO = [
    os.path.join(IMG_PATH, "btn_empleados_nuevo.png"),
    os.path.join(IMG_PATH, "btn_empleados_nuevo2.png"),
    os.path.join(IMG_PATH, "btn_empleados_nuevo3.png")
]

IMG_BTN_PAGO_OCASIONAL = [
    os.path.join(IMG_PATH, "btn_pago_ocasional.png"),
    os.path.join(IMG_PATH, "btn_pago_ocasional2.png"),
    os.path.join(IMG_PATH, "btn_pago_ocasional3.png")
]

IMG_BTN_REGISTRO_ACTUALIZAR = [
    os.path.join(IMG_PATH, "btn_registro_actualizar.png"),
    os.path.join(IMG_PATH, "btn_registro_actualizar2.png"),
    os.path.join(IMG_PATH, "btn_registro_actualizar3.png")
]

IMG_BTN_REGISTRO_CERRAR = [
    os.path.join(IMG_PATH, "btn_registro_cerrar.png"),
    os.path.join(IMG_PATH, "btn_registro_cerrar2.png"),
    os.path.join(IMG_PATH, "btn_registro_cerrar3.png")
]

IMG_BTN_REGISTRO_HABERES = [
    os.path.join(IMG_PATH, "btn_registro_haberes.png"),
    os.path.join(IMG_PATH, "btn_registro_haberes2.png"),
    os.path.join(IMG_PATH, "btn_registro_haberes3.png")
]

IMG_BTN_REGISTRO_INGRESAR = [
    os.path.join(IMG_PATH, "btn_registro_ingresar.png"),
    os.path.join(IMG_PATH, "btn_registro_ingresar2.png"),
    os.path.join(IMG_PATH, "btn_registro_ingresar3.png")
]

IMG_BTN_HABERES_NUEVO = [
    os.path.join(IMG_PATH, "btn_haberes_nuevo.png"),
    os.path.join(IMG_PATH, "btn_haberes_nuevo2.png"),
    os.path.join(IMG_PATH, "btn_haberes_nuevo3.png")
]


IMG_BTN_NEXUS = [
    os.path.join(IMG_PATH, "btn_nexus.png"),
    os.path.join(IMG_PATH, "btn_nexus2.png"),
    os.path.join(IMG_PATH, "btn_nexus3.png"),
]


def validar_equipo_y_obtener_datos(response, mac_actual):
    if response.status_code == 200:
        data = json.loads(response.text)
        for entry in data:
            if entry["MAC"] == mac_actual:
                usuario = entry["Usuario"]
                registros = entry["Registros"]
                ugel = entry["UGEL"]
                mac = entry["MAC"]
                periodo = entry["Periodo"]
                iniciar = entry["Iniciar"]

                documents_path = Path(os.path.expanduser("~/Documents"))
                logs_directory = documents_path / "logs"
                if not logs_directory.exists():
                    logs_directory.mkdir()
                log_file_path = logs_directory / "log_addOcasionales.txt"

                cantidad_minima_registros = registros

                return {
                    "usuario": usuario,
                    "registros": registros,
                    "ugel": ugel,
                    "mac": mac,
                    "periodo": periodo,
                    "iniciar": iniciar,
                    "log_file_path": log_file_path,
                    "cantidad_minima_registros": cantidad_minima_registros,
                    "autorizado": True
                }
        return {"autorizado": False}
    else:
        raise ValueError(
            f"Error al obtener el archivo JSON. Código de estado: {response.status_code}")


def clic_imagen(imagenes):
    try:
        for imagen in imagenes:
            for i in range(INTENTOS):
                location = pyautogui.locateOnScreen(imagen)
                if location:
                    x, y = pyautogui.center(location)
                    pyautogui.click(x, y, duration=DURACION_CLIC)
                    return  # Salir de la función si la imagen se encuentra
                else:
                    # print(f"Intento {i+1} de {INTENTOS}")
                    pyautogui.sleep(0.2)

        # Si llegamos aquí es porque la imagen no se encontró en los intentos
        continuar_bucle_interno = True
        while continuar_bucle_interno:
            root = tk.Tk()
            root.withdraw()  # Oculta la ventana principal de tkinter

            respuesta = messagebox.askyesno(
                "Error al encontrar imagen",
                f"No se pudo encontrar la imagen {imagen}. ¿Quieres continuar?"
            )

            root.destroy()  # Cierra la ventana

            if respuesta:
                for i in range(INTENTOS):  # Reiniciar el contador para el bucle interno
                    location = pyautogui.locateOnScreen(imagen)
                    if location:
                        x, y = pyautogui.center(location)
                        pyautogui.click(x, y, duration=DURACION_CLIC)
                        return  # Salir de la función si la imagen se encuentra
                    else:
                        # print(f"Intento {i+1} de {INTENTOS}")
                        pyautogui.sleep(0.2)
            else:
                continuar_bucle_interno = False
                # Termina el programa
                sys.exit("No se pudo encontrar la imagen")
    except Exception as e:
        print(f"Error al hacer clic en imagen: {e.__class__.__name__}, {e}")


def clic_imagen_derecha(imagenes):
    try:
        for imagen in imagenes:
            for i in range(INTENTOS):
                location = pyautogui.locateOnScreen(imagen)
                if location:
                    x, y, width, height = location
                    x_derecha = x + width + 10  # Calcular la posición derecha de la imagen
                    # Hacer clic en el lado derecho
                    pyautogui.click(x_derecha, y, duration=DURACION_CLIC)
                    return  # Salir de la función si la imagen se encuentra
                else:
                    # print(f"Intento {i+1} de {INTENTOS}")
                    pyautogui.sleep(0.2)

            # Si llegamos aquí es porque la imagen no se encontró en los intentos
            continuar_bucle_interno = True
            while continuar_bucle_interno:
                root = tk.Tk()
                root.withdraw()  # Oculta la ventana principal de tkinter

                respuesta = messagebox.askyesno(
                    "Error al encontrar imagen",
                    f"No se pudo encontrar la imagen {imagen}. ¿Quieres continuar?"
                )

                root.destroy()  # Cierra la ventana

                if respuesta:
                    for i in range(INTENTOS):  # Reiniciar el contador para el bucle interno
                        location = pyautogui.locateOnScreen(imagen)
                        if location:
                            x, y, width, height = location
                            x_derecha = x + width + 10  # Calcular la posición derecha de la imagen
                            # Hacer clic en el lado derecho
                            pyautogui.click(
                                x_derecha, y, duration=DURACION_CLIC)
                            return  # Salir de la función si la imagen se encuentra
                        else:
                            # print(f"Intento {i+1} de {INTENTOS}")
                            pyautogui.sleep(0.2)
                else:
                    continuar_bucle_interno = False
                    # Termina el programa
                    sys.exit("No se pudo encontrar la imagen")
    except Exception as e:
        print(f"Error al hacer clic en imagen: {e}")


def doble_clic_imagen(imagenes):
    try:
        for imagen in imagenes:
            for i in range(INTENTOS):
                location = pyautogui.locateOnScreen(imagen)
                if location:
                    x, y = pyautogui.center(location)
                    pyautogui.doubleClick(x, y, duration=DURACION_CLIC)
                    return  # Salir de la función si la imagen se encuentra
                else:
                    # print(f"Intento {i+1} de {INTENTOS}")
                    pyautogui.sleep(0.2)

        # Si llegamos aquí es porque la imagen no se encontró en los intentos
        continuar_bucle_interno = True
        while continuar_bucle_interno:
            root = tk.Tk()
            root.withdraw()  # Oculta la ventana principal de tkinter

            respuesta = messagebox.askyesno(
                "Error al encontrar imagen",
                f"No se pudo encontrar la imagen {imagen}. ¿Quieres continuar?"
            )

            root.destroy()  # Cierra la ventana

            if respuesta:
                for i in range(INTENTOS):  # Reiniciar el contador para el bucle interno
                    location = pyautogui.locateOnScreen(imagen)
                    if location:
                        x, y = pyautogui.center(location)
                        pyautogui.doubleClick(x, y, duration=DURACION_CLIC)
                        return  # Salir de la función si la imagen se encuentra
                    else:
                        # print(f"Intento {i+1} de {INTENTOS}")
                        pyautogui.sleep(0.2)
            else:
                continuar_bucle_interno = False
                # Termina el programa
                sys.exit("No se pudo encontrar la imagen")
    except Exception as e:
        print(f"Error al hacer clic en imagen: {e.__class__.__name__}, {e}")


def ingresar_y_buscar(modular):
    """Esta función permitirá ingresar el código modular y dar clic en Buscar"""
    try:
        pyautogui.sleep(TIEMPO_INICIO)
        clic_imagen(CODIGO_MODULAR)
        pyautogui.press('delete', presses=15)
        pyautogui.press('backspace', presses=15)
        pyautogui.write(modular)
        pyautogui.press("tab", presses=1)
        pyautogui.press("space")
    except Exception as e:
        print(
            f"Error al ingresar y buscar el código modular {modular}: {e}")


def crear_registro(modular):
    """ Creará un nuevo registro de pago ocasional, si ya está creado, solo se habilitará y actualizará """
    try:
        # Ingresamos el código modular y damos clic en Buscar
        ingresar_y_buscar(modular)
        # Hacemos clic en Pagos Ocacionales
        clic_imagen(IMG_BTN_PAGO_OCASIONAL)
        pyautogui.sleep(1.6)
        # Hacemos clic en Ingresar
        clic_imagen(IMG_BTN_REGISTRO_INGRESAR)
        pyautogui.press("space")
        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
        pyautogui.press("space")
        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
        # Hacemos clic en actualizar
        clic_imagen(IMG_BTN_REGISTRO_ACTUALIZAR)
        pyautogui.press("space")
        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
        # Cerramos la ventana
        cerrar_ventana_registro_first()
    except Exception as e:
        print(
            f"Error en la función crear_registro para el modular {modular}: {e}")
        raise


def ingresar_haber(haber, monto, airhsp, periodo):
    """ Esta función sirve para crear un nuevo haber con código de pago Ocasional"""
    try:
        # nos ubicamos en el botón haber es y hacemos clic
        clic_imagen(IMG_BTN_REGISTRO_HABERES)
        pyautogui.sleep(TIEMPO_ABRIR_HABERES)
        # Agregamos un nuevo haber
        clic_imagen(IMG_BTN_HABERES_NUEVO)

        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
        pyautogui.press('delete', presses=3)
        pyautogui.press('backspace', presses=3)
        pyautogui.write(haber)
        pyautogui.press('tab', presses=2)
        # Escribimos el monto
        pyautogui.write(monto)
        pyautogui.press('tab')
        # Borramos el contenido del periodo
        pyautogui.press('delete', presses=10)
        pyautogui.press('backspace', presses=10)
        pyautogui.write(periodo)
        pyautogui.press('tab')
        # Borramos el contenido del periodo
        pyautogui.press('delete', presses=10)
        pyautogui.press('backspace', presses=10)
        pyautogui.write(periodo)
        # Nos ubicamos en el botón Ingresar
        pyautogui.press('tab', presses=2)
        # Presionamos el botón Ingresar
        pyautogui.press('space')
        pyautogui.sleep(0.3)
        # Nos ubicamos en el botón Cerrar y lo presionamos
        pyautogui.press('space')
        pyautogui.sleep(0.3)
        # cerrar la ventana de haberes
        pyautogui.press('tab', presses=2)
        pyautogui.press('space')
        pyautogui.sleep(0.3)
        # Agregamos la leyenda mensual
        ingresar_leyenda_mensual(periodo)
        pyautogui.sleep(0.5)
        ingresamos_airhsp(airhsp)
        pyautogui.sleep(0.2)
        # Actualizamos el registro
        actualizar_registro()
    except Exception as e:
        print(f"Error en la función ingresar_haber: {e}")


def ingresar_leyenda_mensual(leyenda_mensual):
    try:
        # Nos ubicamos en el campo de leyenda mensual
        clic_imagen_derecha(LEYENDA_MENSUAL)
        borrar_campo(15)
        pyautogui.write(leyenda_mensual)
        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
    except Exception as e:
        print(f"Error al ingresar la leyenda mensual: {e}")


def ingresar_leyenda_permanente(leyenda_permanente):
    try:
        # Nos ubicamos en el campo de leyenda mensual
        clic_imagen_derecha(LEYENDA_PERMANENTE)
        borrar_campo(50)
        pyautogui.write(leyenda_permanente)
        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
    except Exception as e:
        print(f"Error al ingresar la leyenda mensual: {e}")


def ingresamos_airhsp(airhsp):
    try:
        # Nos ubicamos en el campo de leyenda mensual
        clic_imagen_derecha(CODIGO_AIRHSP)
        pyautogui.press('delete', presses=7)
        pyautogui.press('backspace', presses=7)
        pyautogui.write(airhsp)
        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
    except Exception as e:
        print(f"Error al ingresar Codigo AIRHSP: {e}")


def cerrar_ventana_registro_first():
    """Esta función cierra la ventana de registro first"""
    try:
        clic_imagen(IMG_BTN_REGISTRO_CERRAR)
        pyautogui.sleep(0.7)
    except Exception as e:
        print(
            f"Ha ocurrido un error al cerrar la ventana de registro first: {e}")


def cerrar_ventana_registro():
    """Esta función cierra la ventana de registro first"""
    try:
        clic_imagen(IMG_BTN_REGISTRO_CERRAR)
        pyautogui.sleep(0.9)
        pyautogui.press('tab', presses=17)
        pyautogui.press('delete', presses=15)
        pyautogui.press('backspace', presses=15)
        pyautogui.click(30, 30)
    except Exception as e:
        print(f"Error al cerrar la ventana de registro: {e}")


def log_event(event, log_file_name):
    timestamp = str(datetime.datetime.now())
    log_line = f"{timestamp} - {event}\n"
    # Ruta de la carpeta "Mis documentos"
    documents_path = Path(os.path.expanduser("~/Documents"))
    # Crear la carpeta 'logs' dentro de "Mis documentos" si no existe
    logs_directory = documents_path / "logs"
    logs_directory.mkdir(parents=True, exist_ok=True)
    # Seleccionar el archivo de log basado en el parámetro log_file_name
    log_file_path = logs_directory / log_file_name

    with open(log_file_path, "a") as log_file:
        log_file.write(log_line)


def error_registro_contrato(position, cod_modular):
    try:
        """ Esta función sirve para ingresar al registro del docente seleccionado"""
        # Ingresamos el código modular y damos clic en Buscar
        # ingresar_y_buscar(modular)
        log_event(
            f"{position}: Registro no procesado: {cod_modular}", "log_addContratosError.txt")
        pyautogui.sleep(0.6)
        pyautogui.press('space')
        pyautogui.sleep(0.6)

        clic_imagen(IMG_BTN_REGISTRO_CERRAR)
        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
        pyautogui.press('tab', presses=9)
        pyautogui.press('delete', presses=15)
        pyautogui.press('backspace', presses=15)
        pyautogui.click(30, 30)
    except Exception as e:
        print(f"Error al cerrar la ventana de registro: {e}")


def ingresar_al_registro(position, modular, haber, monto, airhsp, periodo):
    try:
        """ Esta función sirve para ingresar al registro del docente seleccionado"""
        # Ingresamos el código modular y damos clic en Buscar
        # ingresar_y_buscar(modular)
        pyautogui.sleep(0.6)

        doble_clic_imagen(SELECCIONAR_REGISTRO_OCASIONAL)
        print(position, ": Ingresando Ocasional de: ", modular)
        ingresar_haber(haber, monto, airhsp, periodo)
        cerrar_ventana_registro()
        # print("Registrado con éxito")

        log_event(f"{position}: Registrado con éxito: {modular}",
                  "log_addOcasionales.txt")

    except Exception as e:
        print(f"Error al ingresar al registro: {e}")


def ingresar_al_registro_habilitado(position, modular):
    try:
        """ Esta función sirve para ingresar al registro del docente seleccionado"""
        # Ingresamos el código modular y damos clic en Buscar
        # ingresar_y_buscar(modular)
        pyautogui.sleep(0.6)

        doble_clic_imagen(SELECCIONAR_REGISTRO_HABILITADO)
        # print(position, ": Ingresando reintegro de: ", modular)
        # print("Registrado con éxito")

    except Exception as e:
        print(f"Error al ingresar al registro: {e}")


def ingresar_haber_reintegros(monto92, monto99, periodo):
    """ Esta función sirve para crear un nuevo haber de reintegros"""
    try:
        # nos ubicamos en el botón haber es y hacemos clic
        clic_imagen(IMG_BTN_REGISTRO_HABERES)
        pyautogui.sleep(TIEMPO_ABRIR_HABERES)
        # Ingresar Haber 92
        clic_imagen(IMG_BTN_HABERES_NUEVO)

        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
        pyautogui.press('delete', presses=3)
        pyautogui.press('backspace', presses=3)
        pyautogui.write("92")
        pyautogui.press('tab', presses=2)
        # Escribimos el monto
        pyautogui.write(monto92)
        pyautogui.press('tab')
        # Borramos el contenido del periodo
        pyautogui.press('delete', presses=10)
        pyautogui.press('backspace', presses=10)
        pyautogui.write(periodo)
        pyautogui.press('tab')
        # Borramos el contenido del periodo
        pyautogui.press('delete', presses=10)
        pyautogui.press('backspace', presses=10)
        pyautogui.write(periodo)
        # Nos ubicamos en el botón Ingresar
        pyautogui.press('tab', presses=2)
        # Presionamos el botón Ingresar
        pyautogui.press('space')
        pyautogui.sleep(0.3)
        # Nos ubicamos en el botón Cerrar y lo presionamos
        pyautogui.press('space')
        pyautogui.sleep(0.3)

        # Ingresar Haber 99
        clic_imagen(IMG_BTN_HABERES_NUEVO)

        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
        pyautogui.press('delete', presses=3)
        pyautogui.press('backspace', presses=3)
        pyautogui.write("99")
        pyautogui.press('tab', presses=2)
        # Escribimos el monto
        pyautogui.write(monto99)
        pyautogui.press('tab')
        # Borramos el contenido del periodo
        pyautogui.press('delete', presses=10)
        pyautogui.press('backspace', presses=10)
        pyautogui.write(periodo)
        pyautogui.press('tab')
        # Borramos el contenido del periodo
        pyautogui.press('delete', presses=10)
        pyautogui.press('backspace', presses=10)
        pyautogui.write(periodo)
        # Nos ubicamos en el botón Ingresar
        pyautogui.press('tab', presses=2)
        # Presionamos el botón Ingresar
        pyautogui.press('space')
        pyautogui.sleep(0.3)
        # Nos ubicamos en el botón Cerrar y lo presionamos
        pyautogui.press('space')
        pyautogui.sleep(0.3)

        # cerrar la ventana de haberes
        pyautogui.press('tab', presses=2)
        pyautogui.press('space')
        pyautogui.sleep(0.3)

    except Exception as e:
        print(f"Error en la función ingresar_haber: {e}")


def actualizar_registro():
    try:
        """ Esta función sirve para actualizar el registro luego de haber realizado cualquier modificación a la misma"""
        # Nos ubicamos en el botón Actualizar
        clic_imagen(IMG_BTN_REGISTRO_ACTUALIZAR)
        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
        pyautogui.press("space")
    except Exception as e:
        print("Error al actualizar registro: {e}")


def buscar_modular(cod_modular):
    """Esta función permitirá ingresar el código modular y dar clic en Buscar"""
    try:
        pyautogui.sleep(TIEMPO_INICIO)
        clic_imagen(CODIGO_MODULAR)
        pyautogui.write(cod_modular)
        pyautogui.press("tab", presses=1)
        pyautogui.press("space")
    except Exception as e:
        print(
            f"Error al ingresar y buscar el código modular {cod_modular}: {e}")


def ingresar_modular_nombrado(cod_modular):
    """Esta función permitirá ingresar el código modular para registrar a nombrado y dar clic en registrar"""
    try:
        clic_imagen_derecha(CODIGO_MODULAR_NOMBRADO)
        borrar_campo(5)
        pyautogui.write(cod_modular)
        pyautogui.press("tab", presses=1)
        pyautogui.press("space")
    except Exception as e:
        print(
            f"Error al ingresar y buscar el código modular {cod_modular} del registro de nombrado: {e}")


def ingresar_dni_nombrado(dni):
    """Esta función permitirá ingresar el dni y dar clic en Registrar para habilitar el código modular"""
    try:
        pyautogui.sleep(TIEMPO_INICIO)
        clic_imagen_derecha(DNI)
        borrar_campo(10)
        pyautogui.write(dni)
        pyautogui.press("tab", presses=1)
        pyautogui.press("space")
        pausar(TIEMPO_PRESIONAR_BOTON)
    except Exception as e:
        print(
            f"Error al ingresar y buscar el DNI {dni}: {e}")


def pausar(tiempo):
    pyautogui.sleep(tiempo)


def borrar_campo(cantidad):
    pyautogui.press('delete', presses=cantidad)
    pyautogui.press('backspace', presses=cantidad)


def ingresar_codigo_plaza(cod_plaza):
    clic_imagen_derecha(CODIGO_NEXUS)
    borrar_campo(12)
    pyautogui.write(cod_plaza)
    clic_imagen(IMG_BTN_NEXUS)
    pausar(TIEMPO_PRESIONAR_BOTON)


def ingresar_codigo_plaza_nombrado(cod_plaza):
    clic_imagen_derecha(CODIGO_NEXUS)
    borrar_campo(12)
    pyautogui.write(cod_plaza)
    pausar(TIEMPO_PRESIONAR_BOTON)


def ingresar_codigo_airhsp(airhsp):
    clic_imagen_derecha(CODIGO_AIRHSP)
    borrar_campo(12)
    pyautogui.write(airhsp)
    pausar(TIEMPO_PRESIONAR_BOTON)


def ingresar_documento_referencia(doc_referencia):
    clic_imagen_derecha(DOCUMENTO_REFERENCIA)
    borrar_campo(20)
    pyautogui.write(doc_referencia)
    pausar(TIEMPO_PRESIONAR_BOTON)


def ingresar_vinculo_laboral(fecha_inicio, fecha_termino):
    pausar(TIEMPO_ABRIR_NUEVO)
    borrar_campo(10)
    pyautogui.write(fecha_inicio)
    pyautogui.press('tab')
    borrar_campo(10)
    pyautogui.write(fecha_termino)
    pausar(TIEMPO_PRESIONAR_BOTON)


def ingresar_nec(nec):
    clic_imagen_derecha(NEC)
    borrar_campo(2)
    pyautogui.write(nec)
    pausar(TIEMPO_PRESIONAR_BOTON)


def ingresar_clave8(clave8):
    clic_imagen_derecha(ESTABLECIMIENTO)
    borrar_campo(8)
    pyautogui.write(clave8)
    pausar(TIEMPO_PRESIONAR_BOTON)


def ingresar_regimen_pensionario(cod_regimen, tipo_afp, cuspp, fecha_afil, fecha_devengue):
    clic_imagen(CODIGO_FISCAL)
    pyautogui.write(cod_regimen)
    if cod_regimen == "2":
        pausar(TIEMPO_PRESIONAR_BOTON)
    else:
        pyautogui.write(tipo_afp)
        pyautogui.press("tab", presses=2)
        pyautogui.write(cuspp)
        pyautogui.press("tab", presses=1)
        pyautogui.write(fecha_afil)
        pyautogui.press("tab", presses=1)
        pyautogui.write(fecha_devengue)
        pyautogui.press("tab", presses=1)
        pausar(TIEMPO_PRESIONAR_BOTON)


def ingresar_nivel_magisterial(nivel_magisterial):
    clic_imagen(NIVEL_MAGISTERIAL)
    borrar_campo(1)
    pyautogui.write(nivel_magisterial)
    pausar(TIEMPO_PRESIONAR_BOTON)


def ingresar_dias_laborados(dias_laborados):
    clic_imagen(DIAS_LABORADOS)
    borrar_campo(2)
    pyautogui.write(dias_laborados)
    pausar(TIEMPO_PRESIONAR_BOTON)


def ingresar_cuenta(cuenta):
    if cuenta == "0":
        pausar(TIEMPO_PRESIONAR_BOTON)
    else:
        clic_imagen(MODO_PAGO)
        borrar_campo(3)
        pyautogui.write("2")
        pyautogui.press("tab", presses=2)
        pyautogui.write(cuenta)
        pausar(TIEMPO_PRESIONAR_BOTON)


def clic_ingresar():
    # Hacemos clic en Ingresar
    clic_imagen(IMG_BTN_REGISTRO_INGRESAR)
    pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
    pyautogui.press("space")
    pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
    pyautogui.press("space")
    pyautogui.sleep(TIEMPO_PASAR_CAMPOS)


def clic_cerrar_error():
    try:
        """ Esta función sirve para ingresar al registro del docente seleccionado"""
        # Ingresamos el código modular y damos clic en Buscar
        # ingresar_y_buscar(modular)
        clic_imagen(IMG_BTN_REGISTRO_CERRAR)
        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
        pyautogui.press('tab', presses=9)
        pyautogui.press('delete', presses=15)
        pyautogui.press('backspace', presses=15)
        pyautogui.click(30, 30)
    except Exception as e:
        print(f"Error al ingresar al registro: {e}")


def clic_cerrar(position, cod_modular):
    try:
        """ Esta función sirve para ingresar al registro del docente seleccionado"""
        # Ingresamos el código modular y damos clic en Buscar
        # ingresar_y_buscar(modular)
        clic_imagen(IMG_BTN_REGISTRO_CERRAR)
        pyautogui.sleep(TIEMPO_PASAR_CAMPOS)
        pyautogui.press('tab', presses=9)
        pyautogui.press('delete', presses=15)
        pyautogui.press('backspace', presses=15)
        pyautogui.click(30, 30)

    except Exception as e:
        print(f"Error al ingresar al registro: {e}")


def esperar_por_imagen_nexus(imagen, tiempo_maximo, orden, modular, intervalo=0.5):
    """
    Espera dinámicamente por una imagen en la pantalla hasta que aparezca o hasta que se alcance el tiempo máximo.

    :param imagen: Ruta a la imagen que se busca.
    :param tiempo_maximo: Tiempo máximo en segundos para esperar por la imagen.
    :param intervalo: Intervalo en segundos para revisar la presencia de la imagen.
    """
    tiempo_inicio = time.time()
    while True:
        if pyautogui.locateOnScreen(imagen) is not None:
            # print("Datos cargados de NEXUS, Continuando...")
            return True  # La imagen fue encontrada
        elif time.time() - tiempo_inicio > tiempo_maximo:
            mensaje_error = f"{orden}: Registro no procesado: {modular}"
            print(mensaje_error)
            # Actualizado para llamar a la nueva función
            error_registro_contrato(orden, modular)
            return False  # La imagen no fue encontrada y se alcanzó el tiempo máximo
        else:
            # print("Imagen no encontrada, esperando...")
            time.sleep(intervalo)


def registrar_ocasionales(df, log_file_path, cantidad_minima_registros, rows_to_drop):

    print("\n==============================================")
    print("Iniciando el registro de Planilla Ocasional...")
    print("==============================================\n")
    try:
        for i, row in df.iterrows():
            if i != 0:
                if validar_cantidad_registros(log_file_path, cantidad_minima_registros):
                    orden = df["orden"][i]
                    modular = df["modular"][i]
                    haber = HABER_OCASIONAL
                    monto = str(df["monto"][i])
                    airhsp = CODIGO_AIRHSP_OCASIONAL
                    periodo = PERIODO_OCASIONAL

                    # print(monto, " ", type(monto), modular,)
                    crear_registro(modular)
                    ingresar_al_registro(
                        orden, modular, haber, monto, airhsp, periodo)

                    rows_to_drop.append(i)
                    # Eliminamos la fila actual
                else:
                    # Mensaje de error por exceso de registros
                    mensaje_error = (
                        "Error: Registros superados... La cantidad de registros procesados autorizados ha sido superada.\n")
                    print(mensaje_error)
                    sys.exit()
        print("Registro Ocasional concluido.  :) \n")
        df = df.drop(rows_to_drop)
        df.to_csv("data.csv", index=False)

    except Exception as e:
        print(e)

    finally:
        # Esperar entrada del usuario antes de salir

        input("Presiona Enter para salir...")
        sys.exit()


def registrar_reintegros(df, log_file_path, cantidad_minima_registros, rows_to_drop):

    print("\n==============================================")
    print("Iniciando el registro de Planilla Reintegros...")
    print("==============================================\n")
    try:
        for i, row in df.iterrows():
            if i != 0:
                if validar_cantidad_registros(log_file_path, cantidad_minima_registros):
                    orden = df["orden"][i]
                    modular = df["modular"][i]
                    monto92 = str(df["monto92"][i])
                    monto99 = str(df["monto99"][i])
                    leyenda_mensual = df["leyenda_mensual"][i]
                    leyenda_permanente = df["leyenda_permanente"][i]
                    periodo = df["periodo"][i]

                    # print(monto, " ", type(monto), modular,)
                    # crear_registro(modular)
                    buscar_modular(modular)
                    ingresar_al_registro_habilitado(orden, modular)
                    pyautogui.sleep(0.5)
                    ingresar_haber_reintegros(monto92, monto99, periodo)
                    if leyenda_mensual != "0":
                        ingresar_leyenda_mensual(leyenda_mensual)
                    else:
                        pass

                    if leyenda_permanente != "0":
                        ingresar_leyenda_permanente(leyenda_permanente)
                    else:
                        pass

                    actualizar_registro()
                    print(f"{orden}: Registro procesado: {modular}")

                    log_event(f"{orden}: Registrado con éxito: {modular}",
                              "log_addReintegro.txt")
                    rows_to_drop.append(i)
                    cerrar_ventana_registro()
                    # Eliminamos la fila actual
                else:
                    # Mensaje de error por exceso de registros
                    mensaje_error = (
                        "Error: Registros superados... La cantidad de registros procesados autorizados ha sido superada.\n")
                    print(mensaje_error)
                    sys.exit()
        print("Registro de reintegros concluido.  :) \n")
        df = df.drop(rows_to_drop)
        df.to_csv("data.csv", index=False)

    except Exception as e:
        print(e)

    finally:
        # Esperar entrada del usuario antes de salir

        input("Presiona Enter para salir...")
        sys.exit()


def registrar_contratos(df, log_file_path, cantidad_minima_registros, rows_to_drop):
    import datetime

    print("\n========================================")
    print("Iniciando el registro de contratos...")
    print("========================================\n")
    try:
        for i, row in df.iterrows():
            if i != 0:
                if validar_cantidad_registros(log_file_path, cantidad_minima_registros):
                    orden = df["orden"][i]
                    dni = df["dni"][i]
                    modular = str(df["modular"][i])
                    cod_plaza = df["cod_plaza"][i]
                    rdl = df["rdl"][i]
                    cod_regimen = str(df["cod_regimen"][i])
                    tipo_afp = str(df["tipo_afp"][i])
                    cuspp = str(df["cuspp"][i])
                    fecha_afil = str(df["fecha_afil"][i])
                    fecha_devengue = str(df["fecha_devengue"][i])
                    nivel_magisterial = str(df["nivel_magisterial"][i])
                    dias_laborados = str(df["dias_laborados"][i])
                    cuenta = df["cuenta"][i]
                    codigo_airhsp = df["airhsp"][i]
                    leyenda_mensual = df["leyenda_mensual"][i]
                    leyenda_permanente = df["leyenda_permanente"][i]

                    # Comienza la ejecución del programa
                    buscar_modular(modular)
                    clic_imagen(IMG_BTN_EMPLEADOS_NUEVO)
                    pausar(TIEMPO_ABRIR_NUEVO)
                    ingresar_codigo_plaza(cod_plaza)

                    # Verificar si aparece el mensaje de error donde el docente no está registrado en Nexus

                    if not esperar_por_imagen_nexus(CARGA_NEXUS_EXITOSA, TIEMPO_MAXIMO_CARGA_NEXUS, orden, modular):
                        continue
                    ingresar_codigo_airhsp(codigo_airhsp)
                    ingresar_documento_referencia(rdl)
                    ingresar_regimen_pensionario(
                        cod_regimen, tipo_afp, cuspp, fecha_afil, fecha_devengue)
                    if nivel_magisterial != 0:
                        ingresar_nivel_magisterial(
                            nivel_magisterial)
                    else:
                        pass

                    ingresar_dias_laborados(dias_laborados)
                    if cuenta != 0:
                        ingresar_cuenta(cuenta)
                    else:
                        pass

                    if leyenda_mensual != "0":
                        ingresar_leyenda_mensual(leyenda_mensual)
                    else:
                        pass

                    if leyenda_permanente != "0":
                        ingresar_leyenda_permanente(leyenda_permanente)
                    else:
                        pass

                    clic_ingresar()
                    clic_cerrar(orden, modular)
                    print(f"{orden}: Registro procesado: {modular}")

                    log_event(f"{orden}: Registrado con éxito: {modular}",
                              "log_addContratos.txt")
                    # Eliminamos la fila actual
                else:
                    # Mensaje de error por exceso de registros
                    mensaje_error = (
                        "Error: Registros superados... La cantidad de registros procesados autorizados ha sido superada.\n")
                    print(mensaje_error)
                    sys.exit()
        print("Registro de Contratos concluido.  :) \n")
        df = df.drop(rows_to_drop)
        df.to_csv("data.csv", index=False)

    except Exception as e:
        print(e)

    finally:
        # Esperar entrada del usuario antes de salir

        input("Presiona Enter para salir...")
        sys.exit()


def registrar_nombrados(df, log_file_path, cantidad_minima_registros, rows_to_drop):
    import datetime

    print("\n========================================")
    print("Iniciando el registro de nombrados...")
    print("========================================\n")
    try:
        for i, row in df.iterrows():
            if i != 0:
                if validar_cantidad_registros(log_file_path, cantidad_minima_registros):
                    orden = df["orden"][i]
                    dni = df["dni"][i]
                    modular = str(df["modular"][i])
                    cod_plaza = df["cod_plaza"][i]
                    rdl = df["rdl"][i]
                    fecha_inicio = df["fecha_inicio"][i]
                    fecha_termino = df["fecha_termino"][i]
                    cod_regimen = str(df["cod_regimen"][i])
                    tipo_afp = str(df["tipo_afp"][i])
                    cuspp = str(df["cuspp"][i])
                    fecha_afil = str(df["fecha_afil"][i])
                    fecha_devengue = str(df["fecha_devengue"][i])
                    nec = df["nec"][i]
                    clave8 = df["clave8"][i]
                    cuenta = df["cuenta"][i]
                    codigo_airhsp = df["airhsp"][i]
                    leyenda_mensual = df["leyenda_mensual"][i]
                    leyenda_permanente = df["leyenda_permanente"][i]

                    # Comienza la ejecución del programa
                    ingresar_dni_nombrado(dni)

                    # Verificar si aparece el mensaje de error donde el docente no está registrado en Nexus
                    pausar(0.2)
                    if pyautogui.locateOnScreen(ERROR_PADRON_NEXUS_NOMBRADO) is not None:
                        # Registra el error en el log
                        log_event(
                            f"{orden}: Registro nombrado no procesado: {modular}", "log_addNombradosError.txt")
                        # Presiona la tecla espacio y hace clic en el botón de cerrar
                        pyautogui.press('space')
                        print(f"{orden}: Registro no procesado: {modular}")
                        # Pasa al siguiente registro
                        continue
                    else:
                        pass

                    ingresar_modular_nombrado(modular)
                    pausar(0.2)
                    if pyautogui.locateOnScreen(ERROR_PADRON_NEXUS_NOMBRADO_MODULAR) is not None:
                        # Registra  Nombrado modular.", "log_addNombradosError.txt")
                        # Presiona la tecla espacio y hace clic en el botón de cerrar
                        pyautogui.press('space')
                        print(f"{orden}: Registro no procesado: {modular}")
                        # Pasa al siguiente registro
                        continue
                    else:
                        pass

                    ingresar_vinculo_laboral(fecha_inicio, fecha_termino)

                    ingresar_documento_referencia(rdl)

                    ingresar_regimen_pensionario(
                        cod_regimen, tipo_afp, cuspp, fecha_afil, fecha_devengue)
                    ingresar_nec(nec)
                    ingresar_clave8(clave8)
                    ingresar_codigo_airhsp(codigo_airhsp)
                    ingresar_codigo_plaza_nombrado(cod_plaza)

                    if cuenta != "0":
                        ingresar_cuenta(cuenta)
                    else:
                        pass

                    if leyenda_mensual != "0":
                        ingresar_leyenda_mensual(leyenda_mensual)
                    else:
                        pass

                    if leyenda_permanente != "0":
                        ingresar_leyenda_permanente(leyenda_permanente)
                    else:
                        pass
                    clic_ingresar()
                    clic_cerrar(orden, modular)
                    print(f"{orden}: Registro procesado: {modular}")

                    log_event(f"{orden}: Registrado con éxito: {modular}",
                              "log_addNombrados.txt")
                    # Eliminamos la fila actual
                else:
                    # Mensaje de error por exceso de registros
                    mensaje_error = (
                        "Error: Registros superados... La cantidad de registros procesados autorizados ha sido superada.\n")
                    print(mensaje_error)
                    sys.exit()
        print("Registro de Nombrados concluido.  :) \n")
        df = df.drop(rows_to_drop)
        df.to_csv("data.csv", index=False)

    except Exception as e:
        print(e)

    finally:
        # Esperar entrada del usuario antes de salir

        input("Presiona Enter para salir...")
        sys.exit()


def main():
    datos = validar_equipo_y_obtener_datos(response, mac_actual)
    if datos["autorizado"]:
        # Encontramos la MAC, ahora puedes acceder a los datos
        usuario = datos["usuario"]
        registros = datos["registros"]
        ugel = datos["ugel"]
        # mac = datos["mac"]
        # periodo = datos["periodo"]
        # iniciar = datos["iniciar"]

        # Obtener la ruta de la carpeta Mis Documentos y la carpeta logs
        documents_path = Path(os.path.expanduser("~/Documents"))
        logs_directory = documents_path / "logs"

        # Verificar si la carpeta logs existe, si no, crearla
        if not logs_directory.exists():
            logs_directory.mkdir()

        # Construir la ruta del archivo de log Registro Ocasionales
        log_file_path_ocasionales = logs_directory / "log_addOcasionales.txt"

        # Construir la ruta del archivo de los Registros de Reintegros
        log_file_path_reintegros = logs_directory / "log_addReintegros.txt"

        # Construir la ruta del archivo de log Registro Contratos
        log_file_path_contratos = logs_directory / "log_addContratos.txt"

        # Construir la ruta del archivo de log Registro Contratos
        log_file_path_nombrados = logs_directory / "log_addNombrados.txt"

        # Validar la cantidad de líneas en el archivo de log
        cantidad_minima_registros = registros
        print("\n\n========================================================")
        print(f"Bienvenid@ {usuario} - {ugel}")
        print("========================================================\n")
        # Damos la bienvenida al usuario

        rows_to_drop = []

        pyautogui.sleep(1.5)
        while True:
            print("\tSeleccione la opción que desea ejecutar:")
            print("\t1. Registrar Ocasionales")
            print("\t2. Registrar Contratos")
            print("\t3. Registrar Nombrados")
            print("\t4. Registrar Reintegros")
            print("\t5. Salir")
            opcion = input(
                f"\nIngrese el número de la opción deseada ó \nPresione Enter para la opción predeterminada ({PREDETERMINADO}): ")
            # Establecer el valor predeterminado si la entrada es vacía
            if opcion == '':
                opcion = PREDETERMINADO  # Valor predeterminado

            mensaje_inicio = f"\nEl programa iniciará en {TIEMPO_INICIAR_PROGRAMA} segundos!"

            if opcion == '1':
                df_registrar_ocasionales = pd.read_csv(
                    CSV_FILE_REGISTRAR_OCASIONALES, names=NOMBRE_COLUMNAS_REGISTRAR_OCASIONALES)
                imprimir_encabezado(columnas_ocasionales)
                print(mensaje_inicio)
                pyautogui.sleep(TIEMPO_INICIAR_PROGRAMA)
                registrar_ocasionales(
                    df_registrar_ocasionales, log_file_path_ocasionales, cantidad_minima_registros, rows_to_drop)
            elif opcion == '2':
                df_registrar_contratos = pd.read_csv(
                    CSV_FILE_REGISTRAR_CONTRATOS, names=NOMBRE_COLUMNAS_REGISTRAR_CONTRATOS, skiprows=0, dtype={'cuenta': str, 'airhsp': str})
                imprimir_encabezado(columnas_contratados)
                print(mensaje_inicio)
                pyautogui.sleep(TIEMPO_INICIAR_PROGRAMA)
                registrar_contratos(
                    df_registrar_contratos, log_file_path_contratos, cantidad_minima_registros, rows_to_drop)
            if opcion == '3':
                df_registrar_nombrados = pd.read_csv(
                    CSV_FILE_REGISTRAR_NOMBRADOS, names=NOMBRE_COLUMNAS_REGISTRAR_NOMBRADOS)
                imprimir_encabezado(columnas_nombrados)
                print(mensaje_inicio)
                pyautogui.sleep(TIEMPO_INICIAR_PROGRAMA)
                registrar_nombrados(
                    df_registrar_nombrados, log_file_path_nombrados, cantidad_minima_registros, rows_to_drop)
            if opcion == '4':
                df_registrar_reintegros = pd.read_csv(
                    CSV_FILE_REGISTRAR_REINTEGROS, names=NOMBRE_COLUMNAS_REGISTRAR_REINTEGROS)
                imprimir_encabezado(columnas_reintegros)
                print(mensaje_inicio)
                pyautogui.sleep(TIEMPO_INICIAR_PROGRAMA)
                registrar_reintegros(
                    df_registrar_reintegros, log_file_path_reintegros, cantidad_minima_registros, rows_to_drop)
            elif opcion == '5':
                print("\nSaliendo del programa...")
                break
            else:
                print("Opción no válida. Por favor, intente de nuevo.")
    else:
        # Mensaje de error y contacto
        mensaje_error = (
            "\nERROR: Este equipo no está autorizado para usar el programa.\n\n"
            "Si desea usar el programa debe contactarse con el desarrollador.\n"
            "Número de contacto: 948-924-822\n"
            f"Código de su equipo es: {mac_actual}\n"
        )
        print(mensaje_error)

        # Esperar entrada del usuario antes de salir
        input("Presiona Enter para salir...")


if __name__ == '__main__':
    main()
