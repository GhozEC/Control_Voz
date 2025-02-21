import speech_recognition as sr
import pyautogui
import os
import re
import tkinter as tk
import threading
from pycaw.pycaw import AudioUtilities
from pycaw.pycaw import IAudioEndpointVolume
import winsound
import subprocess
import time

# Estado del control por voz
activo = False
ultimo_comando_valido = time.time()

# Función para reproducir sonidos predeterminados de Windows
def reproducir_sonido(accion):
    """Reproduce un sonido dependiendo de la acción."""
    if accion == "activar":
        winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
    elif accion == "desactivar":
        winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
    elif accion == "comando_reconocido":
        winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
    elif accion == "comando_no_reconocido":
        winsound.PlaySound("SystemHand", winsound.SND_ALIAS)

def log(mensaje):
    """Actualiza el último mensaje mostrado en la interfaz."""
    estado_label.config(text=mensaje)
    estado_label.update_idletasks()  # Actualiza la interfaz después de cada cambio

def ajustar_volumen_porcentaje(porcentaje):
    """Ajusta el volumen según el porcentaje dado utilizando pycaw."""
    if 0 <= porcentaje <= 100:
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                IAudioEndpointVolume._iid_, 1, None)
            volume = interface.QueryInterface(IAudioEndpointVolume)
            
            # Ajustar el volumen en función del porcentaje
            volume.SetMasterVolumeLevelScalar(porcentaje / 100, None)
            log(f"🔊 Volumen ajustado al {porcentaje}%.")
        except Exception as e:
            log(f"❌ Error al ajustar el volumen: {e}")
    else:
        log("⚠️ El porcentaje debe ser entre 0 y 100.")

def ejecutar_comando(comando):
    """Ejecuta la acción según el comando reconocido."""
    global activo, ultimo_comando_valido

    if "activar control" in comando:
        if not activo:
            activo = True
            log("✅ Control por voz ACTIVADO.")
            reproducir_sonido("activar")
        else:
            log("⚠️ El control por voz ya está activado.")
        ultimo_comando_valido = time.time()
        return

    if "apagar control" in comando:
        if activo:
            activo = False
            log("❌ Control por voz DESACTIVADO.")
            reproducir_sonido("desactivar")
        else:
            log("⚠️ El control por voz ya está desactivado.")
        ultimo_comando_valido = time.time()
        return

    if activo:
        # Comandos de control
        if "pausa" in comando or "play" in comando:
            pyautogui.press('playpause')
            log("⏯️ Reproducción/Pausa.")
            reproducir_sonido("comando_reconocido")
        elif "siguiente" in comando:
            pyautogui.press('nexttrack')
            log("⏭️ Canción siguiente.")
            reproducir_sonido("comando_reconocido")
        elif "anterior" in comando:
            pyautogui.press('prevtrack')
            log("⏮️ Canción anterior.")
            reproducir_sonido("comando_reconocido")
        elif "subir volumen al" in comando:
            match = re.search(r'(\d+)', comando)
            if match:
                porcentaje = int(match.group(1))
                ajustar_volumen_porcentaje(porcentaje)
            else:
                pyautogui.press('volumeup')
                log("🔊 Subiendo volumen.")
            reproducir_sonido("comando_reconocido")
        elif "bajar volumen al" in comando:
            match = re.search(r'(\d+)', comando)
            if match:
                porcentaje = int(match.group(1))
                ajustar_volumen_porcentaje(porcentaje)
            else:
                pyautogui.press('volumedown')
                log("🔉 Bajando volumen.")
            reproducir_sonido("comando_reconocido")
        elif "abrir notepad" in comando or "bloc de notas" in comando:
            try:
                subprocess.Popen(["notepad.exe"])
                log("📝 Bloc de notas abierto.")
                reproducir_sonido("comando_reconocido")
            except FileNotFoundError:
                log("⚠️ No se encontró Notepad.exe.")
            except Exception as e:
                log(f"❌ Error al abrir Notepad: {e}")
        elif "abrir chrome" in comando:
            chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            if os.path.exists(chrome_path):
                subprocess.Popen([chrome_path])
                log("🌐 Google Chrome abierto.")
                reproducir_sonido("comando_reconocido")
            else:
                log("⚠️ No se encontró Google Chrome.")
        elif "suspender computadora" in comando:
            log("🛑 Suspendiendo la computadora...")
            subprocess.run("shutdown /h", check=True)
            reproducir_sonido("comando_reconocido")
        elif "reiniciar computadora" in comando:
            log("🔄 Reiniciando la computadora en 10 segundos...")
            subprocess.run("shutdown /r /t 10", check=True)
            reproducir_sonido("comando_reconocido")
        elif "apagar computadora" in comando:
            log("⚠️ Apagando la computadora en 10 segundos...")
            subprocess.run("shutdown /s /t 10", check=True)
            reproducir_sonido("comando_reconocido")
        else:
            log("❓ Comando no reconocido.")
            reproducir_sonido("comando_no_reconocido")
        ultimo_comando_valido = time.time()
    else:
        log("🔇 El control por voz está desactivado. Di 'activar control' para iniciar.")

def escuchar_comando():
    """Captura el audio y reconoce comandos continuamente."""
    global activo, ultimo_comando_valido

    recognizer = sr.Recognizer()
    recognizer.energy_threshold = 300
    recognizer.dynamic_energy_threshold = True

    with sr.Microphone() as source:
        log("🎙️ Esperando comandos (di 'activar control' para iniciar)...")

        while True:
            try:
                if activo and time.time() - ultimo_comando_valido > 60:
                    log("🔇 Control por voz inactivo. Desactivando automáticamente.")
                    activo = False
                    reproducir_sonido("desactivar")
                    continue
                audio = recognizer.listen(source, timeout=10)
                comando = recognizer.recognize_google(audio, language='es-ES').lower()
                log(f"🗣️ Comando reconocido: {comando}")
                ejecutar_comando(comando)
            except sr.WaitTimeoutError:
                continue
            except sr.UnknownValueError:
                if activo:
                    log("❌ No se entendió el comando.")
            except sr.RequestError as e:
                log(f"⚠️ Error de conexión: {e}")
            except Exception as e:
                log(f"❌ Error inesperado al procesar el comando: {e}")

def iniciar_escucha():
    """Inicia la escucha en un hilo separado para no bloquear la interfaz gráfica."""
    escucha_thread = threading.Thread(target=escuchar_comando)
    escucha_thread.daemon = True  # Permite que el hilo se termine cuando se cierre la aplicación
    escucha_thread.start()

# Interfaz gráfica con Tkinter
root = tk.Tk()
root.title("Control por Voz")
root.geometry("300x200")
root.configure(bg='#2C3E50')

# Estilo general
style_frame = tk.Frame(root, bg='#2C3E50')
style_frame.pack(expand=True, fill='both', padx=20, pady=20)

# Título
titulo = tk.Label(
    style_frame,
    text="🎙️ Control por Voz",
    font=("Helvetica", 14, "bold"),
    bg='#2C3E50',
    fg='#ECF0F1'
)
titulo.pack(pady=(0, 10))

# Estado actual
estado_label = tk.Label(
    style_frame,
    text="Estado: Inactivo",
    font=("Helvetica", 9),
    bg='#2C3E50',
    fg='#ECF0F1',
    wraplength=250, 
    justify="left" 
)
estado_label.pack(pady=10, fill="both", expand=True)

# Centrar la ventana en la pantalla
root.update_idletasks()
width = root.winfo_width()
height = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f'{width}x{height}+{x}+{y}')

# Ejecutar la escucha directamente al inicio
iniciar_escucha()

# Ejecutar la interfaz gráfica
root.mainloop()
