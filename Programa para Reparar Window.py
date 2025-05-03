import tkinter as tk
from tkinter import messagebox, ttk, simpledialog, filedialog
import os
import platform
import subprocess
import socket
import shutil
import psutil
import speedtest
import pyautogui
import hashlib
from PIL import Image
import win32com.client
import datetime
import docx
import openpyxl
import csv
import ctypes
import sys
from PIL import Image, ImageTk
from tkinter import ttk
from tkinter import simpledialog, messagebox, Toplevel, Listbox, Scrollbar, RIGHT, Y, END, LEFT, BOTH
from tkinter import scrolledtext
import threading
# === Funciones ===

def limpiar_temporales():
    os.system("del /f /s /q %temp%\\*")
    os.system("del /f /s /q C:\\Windows\\Temp\\*")
    messagebox.showinfo("Listo", "Archivos temporales eliminados.")

def vaciar_papelera():
    try:
        os.system("powershell -command \"Clear-RecycleBin -Force\"")
        messagebox.showinfo("Listo", "Papelera de reciclaje vaciada.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo vaciar la papelera: {e}")

def verificar_chkdsk():
    subprocess.run("chkdsk", shell=True)

def reparar_archivos_sfc():
    subprocess.run("sfc /scannow", shell=True)

def verificar_conexion():
    try:
        socket.create_connection(("8.8.8.8", 53))
        messagebox.showinfo("Conexión", "Tienes conexión a Internet.")
    except OSError:
        messagebox.showwarning("Sin conexión", "No tienes conexión a Internet.")

def limpiar_cache_navegadores():
    chrome_cache = os.path.expanduser("~\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache")
    if os.path.exists(chrome_cache):
        shutil.rmtree(chrome_cache)
    
    firefox_cache = os.path.expanduser("~\\AppData\\Local\\Mozilla\\Firefox\\Profiles")
    if os.path.exists(firefox_cache):
        for root, dirs, _ in os.walk(firefox_cache):
            for dir in dirs:
                if "cache2" in dir.lower():
                    shutil.rmtree(os.path.join(root, dir))

    messagebox.showinfo("Listo", "Caché de navegadores eliminada.")

def desfragmentar_disco():
    subprocess.run("defrag C: /O", shell=True)
    messagebox.showinfo("Listo", "Desfragmentación del disco completada.")

def monitorear_temperatura():
    temp = psutil.sensors_temperatures()
    cpu_temp = temp.get("coretemp", None)
    if cpu_temp:
        temperature = cpu_temp[0].current
        messagebox.showinfo("Temperatura", f"La temperatura actual de la CPU es: {temperature}°C")
    else:
        messagebox.showwarning("Error", "No se pudo obtener la temperatura de la CPU.")

def eliminar_programas_innecesarios():
    subprocess.run("appwiz.cpl", shell=True)

def reiniciar_sistema():
    if messagebox.askyesno("Reiniciar", "¿Quieres reiniciar el sistema?"):
        os.system("shutdown /r /t 0")

def actualizar_sistema():
    subprocess.run("powershell -command \"Get-WindowsUpdate -AcceptAll -Install\"", shell=True)

def eliminar_archivos_registro():
    for path in ["C:\\Windows\\Logs", os.path.expanduser("~\\AppData\\Local\\Temp")]:
        if os.path.exists(path):
            shutil.rmtree(path, ignore_errors=True)
    messagebox.showinfo("Listo", "Archivos de registro eliminados.")

def desactivar_programas_inicio():
    subprocess.run("taskmgr", shell=True)

def verificar_memoria_ram():
    subprocess.run("mdsched.exe", shell=True)

def reparar_archivos_boot():
    subprocess.run("bootrec /fixmbr", shell=True)
    subprocess.run("bootrec /fixboot", shell=True)
    subprocess.run("bootrec /rebuildbcd", shell=True)
    messagebox.showinfo("Listo", "Archivos de arranque reparados.")

def obtener_estado_smart():
    try:
        result = subprocess.run("wmic diskdrive get status", shell=True, capture_output=True, text=True)
        if "OK" in result.stdout:
            messagebox.showinfo("Estado SMART", "El disco duro está en buen estado.")
        else:
            messagebox.showwarning("Estado SMART", "El disco duro podría tener problemas.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo verificar el estado SMART: {e}")

def actualizar_drivers():
    subprocess.run("devmgmt.msc", shell=True)

def desactivar_actualizaciones():
    subprocess.run("sc stop wuauserv", shell=True)
    subprocess.run("sc config wuauserv start= disabled", shell=True)
    messagebox.showinfo("Listo", "Actualizaciones automáticas desactivadas.")

def crear_reporte_diagnostico():
    os.system("msinfo32")
    messagebox.showinfo("Reporte", "Informe de diagnóstico generado.")

def informacion_sistema():
    os.system("dxdiag")

def verificar_actualizaciones_windows():
    subprocess.run("powershell -command \"Get-WindowsUpdate\"", shell=True)

def copia_seguridad_registro():
    os.system("reg export HKLM\\Software\\ backup_registro.reg")
    messagebox.showinfo("Backup", "Copia del registro creada como backup_registro.reg")

def espacio_libre_disco():
    _, _, free = shutil.disk_usage("C:\\")
    messagebox.showinfo("Espacio en disco", f"Libre: {free // (2**30)} GB")

def restaurar_sistema():
    subprocess.run("rstrui.exe", shell=True)

def reparar_permisos():
    subprocess.run("icacls C:\\ /T /Q /C /RESET", shell=True)
    messagebox.showinfo("Permisos", "Permisos de archivos reparados.")

def optimizar_arranque():
    subprocess.run("msconfig", shell=True)

def verificar_puertos_abiertos():
    resultado = ""
    for conn in psutil.net_connections(kind='inet'):
        resultado += f"Puerto: {conn.laddr.port}, Estado: {conn.status}\n"
    messagebox.showinfo("Puertos Abiertos", resultado if resultado else "No se detectaron conexiones activas.")

def cambiar_resolucion():
    pyautogui.hotkey('win', 'p')
    messagebox.showinfo("Resolución", "Función de cambiar resolución en desarrollo.")

def verificar_velocidad_internet():
    st = speedtest.Speedtest()
    st.get_best_server()
    down = st.download() / 1_000_000
    up = st.upload() / 1_000_000
    messagebox.showinfo("Velocidad de Internet", f"Descarga: {down:.2f} Mbps\nSubida: {up:.2f} Mbps")

def eliminar_duplicados_en_directorio():
    directorio = filedialog.askdirectory(title="Selecciona una carpeta")
    if not directorio:
        return

    hash_archivos, duplicados = {}, []
    for carpeta, _, archivos in os.walk(directorio):
        for archivo in archivos:
            ruta = os.path.join(carpeta, archivo)
            try:
                with open(ruta, "rb") as f:
                    hash_val = hashlib.md5(f.read()).hexdigest()
                if hash_val in hash_archivos:
                    duplicados.append(ruta)
                else:
                    hash_archivos[hash_val] = ruta
            except: pass

    for dup in duplicados:
        try: os.remove(dup)
        except: pass

    messagebox.showinfo("Duplicados", f"{len(duplicados)} archivos duplicados eliminados.")

def eliminar_duplicados_imagenes():
    directorio = filedialog.askdirectory(title="Selecciona una carpeta")
    if not directorio:
        return

    hash_imagenes, duplicados = {}, []
    for carpeta, _, archivos in os.walk(directorio):
        for archivo in archivos:
            if archivo.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".gif")):
                ruta = os.path.join(carpeta, archivo)
                try:
                    with Image.open(ruta) as img:
                        hash_val = hashlib.md5(img.tobytes()).hexdigest()
                    if hash_val in hash_imagenes:
                        duplicados.append(ruta)
                    else:
                        hash_imagenes[hash_val] = ruta
                except: pass

    for dup in duplicados:
        try: os.remove(dup)
        except: pass

    messagebox.showinfo("Duplicados", f"{len(duplicados)} imágenes duplicadas eliminadas.")

def abrir_panel_control():
    subprocess.run("control", shell=True)

def crear_usuario_nuevo():
    nombre_usuario = simpledialog.askstring("Nuevo Usuario", "Ingresa el nombre del nuevo usuario:")
    if nombre_usuario:
        subprocess.run(f"net user {nombre_usuario} /add", shell=True)
        messagebox.showinfo("Usuario", f"Usuario {nombre_usuario} creado correctamente.")

def configurar_perifericos():
    try:
        os.system("start ms-settings:connecteddevices")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir la configuración de dispositivos: {e}")

def modo_seguro_con_red():
    try:
        subprocess.run('bcdedit /set {current} safeboot network', shell=True)
        subprocess.run('shutdown /r /t 0', shell=True)
        messagebox.showinfo("Reiniciando", "El sistema se reiniciará en modo seguro con red.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo activar el modo seguro: {e}")

def escanear_con_defender():
    try:
        subprocess.run('powershell -Command "Start-MpScan -ScanType QuickScan"', shell=True)
        messagebox.showinfo("Escaneo iniciado", "Se ha iniciado un escaneo rápido con Windows Defender.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo iniciar el escaneo: {e}")

def abrir_herramientas_administrativas():
    try:
        subprocess.run("control admintools", shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudieron abrir las herramientas administrativas: {e}")

def mostrar_procesos_activoss():
    try:
        subprocess.run("taskmgr", shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el Administrador de tareas: {e}")

def programar_tarea_automatica():
    nombre = simpledialog.askstring("Tarea", "Nombre de la tarea:")
    comando = simpledialog.askstring("Comando", "Comando o ruta del script a ejecutar:")
    hora = simpledialog.askstring("Hora", "Hora (HH:MM) en formato 24h:")

    if nombre and comando and hora:
        try:
            subprocess.run(
                f'schtasks /Create /SC DAILY /TN "{nombre}" /TR "{comando}" /ST {hora}',
                shell=True
            )
            messagebox.showinfo("Tarea Programada", f"Tarea '{nombre}' creada para las {hora}.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear la tarea: {e}")

def comprobar_rutas_sistema():
    try:
        rutas = os.environ.get("PATH", "")
        rutas_listadas = "\n".join(rutas.split(";"))
        messagebox.showinfo("Rutas del sistema (PATH)", rutas_listadas)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo obtener las rutas: {e}")

def limpiar_cache_dns():
    try:
        resultado = subprocess.run("ipconfig /flushdns", shell=True, capture_output=True, text=True)
        if "correctamente" in resultado.stdout.lower() or "successfully" in resultado.stdout.lower():
            messagebox.showinfo("DNS", "Caché DNS limpiada correctamente.")
        else:
            messagebox.showwarning("DNS", f"Resultado: {resultado.stdout.strip()}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo limpiar la caché DNS: {e}")

def optimizar_uso_ram():
    try:
        comando = 'powershell -command "[System.GC]::Collect()"'
        subprocess.run(comando, shell=True)
        messagebox.showinfo("RAM", "Intento de optimización de memoria RAM ejecutado.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo optimizar la RAM: {e}")

def exportar_lista_programas():
    try:
        archivo = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Archivo de texto", "*.txt")], title="Guardar lista de programas")
        if archivo:
            resultado = subprocess.run("wmic product get name,version", shell=True, capture_output=True, text=True)
            with open(archivo, "w", encoding="utf-8") as f:
                f.write(resultado.stdout)
            messagebox.showinfo("Exportación completada", f"Lista de programas guardada en:\n{archivo}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo exportar la lista: {e}")

def optimizar_bateria():
    try:
        # Establecer el plan de energía a "ahorro de energía"
        subprocess.run("powercfg /setactive a1841308-3541-4fab-bc81-f71556f20b4a", shell=True)

        # Opcional: deshabilitar dispositivos en segundo plano (modo ahorro)
        # Esto puede ampliarse si deseas usar `reg` para aplicar configuraciones más específicas

        messagebox.showinfo("Optimización de batería", "Se ha activado el modo de ahorro de energía.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo optimizar la batería: {e}")
    
def mejorar_rendimiento_sistema():
    try:
        # Abrir la ventana de propiedades del sistema (configuración de rendimiento)
        subprocess.run("SystemPropertiesPerformance", shell=True)
        messagebox.showinfo("Mejora de rendimiento", "Se ha abierto la configuración de rendimiento del sistema.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir la configuración de rendimiento: {e}")

def desinstalar_programas():
    try:
        # Abre la ventana "Programas y características" para desinstalar programas
        subprocess.run("appwiz.cpl", shell=True)
        messagebox.showinfo("Desinstalar programas", "Se ha abierto la ventana de desinstalación de programas.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir la ventana de desinstalación: {e}")

def realizar_copia_seguridad():
    try:
        # Abrir un cuadro de diálogo para seleccionar la carpeta de destino de la copia de seguridad
        destino = filedialog.askdirectory(title="Selecciona la carpeta de destino para la copia de seguridad")
        if not destino:
            return
        
        # Carpetas importantes a respaldar (puedes personalizarlas según sea necesario)
        carpetas_a_respaldar = [
            os.path.expanduser("~\\Documents"),
            os.path.expanduser("~\\Pictures"),
            os.path.expanduser("~\\Videos"),
            os.path.expanduser("~\\Desktop")
        ]
        
        # Crear un archivo comprimido con las carpetas seleccionadas
        respaldo_nombre = f"copia_seguridad_{platform.node()}.zip"
        shutil.make_archive(os.path.join(destino, respaldo_nombre), 'zip', root_dir=os.path.dirname(carpetas_a_respaldar[0]), base_dir=os.path.basename(carpetas_a_respaldar[0]))

        messagebox.showinfo("Copia de seguridad", f"Copia de seguridad realizada con éxito en: {destino}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo realizar la copia de seguridad: {e}")

def administrar_conexiones_red():
    try:
        conexiones = psutil.net_connections(kind='inet')
        if not conexiones:
            messagebox.showinfo("Conexiones", "No hay conexiones de red activas.")
            return

        # Crear una ventana para mostrar las conexiones activas
        ventana_conexiones = tk.Toplevel()
        ventana_conexiones.title("Conexiones de Red Activas")
        ventana_conexiones.geometry("500x300")

        # Mostrar información sobre las conexiones activas
        text_conexiones = tk.Text(ventana_conexiones, height=12, width=60)
        text_conexiones.pack(padx=10, pady=10)

        for conn in conexiones:
            laddr = conn.laddr
            raddr = conn.raddr
            status = conn.status
            text_conexiones.insert(tk.END, f"Local: {laddr.ip}:{laddr.port} <-> Remota: {raddr.ip if raddr else 'N/A'}:{raddr.port if raddr else 'N/A'} | Estado: {status}\n")

        # Función para cerrar la conexión seleccionada
        def cerrar_conexion():
            try:
                seleccion = text_conexiones.get(tk.SEL_FIRST, tk.SEL_LAST)
                if seleccion:
                    # Extraer la dirección IP y puerto para cerrar la conexión
                    partes = seleccion.split(" <-> ")
                    if len(partes) == 2:
                        ip_remota = partes[1].split(" |")[0].strip()
                        conn = next((c for c in conexiones if c.raddr.ip == ip_remota), None)
                        if conn:
                            conn.close()
                            messagebox.showinfo("Conexión cerrada", f"Conexión con {ip_remota} cerrada.")
                        else:
                            messagebox.showwarning("Error", "No se encontró la conexión para cerrar.")
            except Exception as e:
                messagebox.showerror("Error", f"Error al intentar cerrar la conexión: {e}")

        # Botón para cerrar la conexión seleccionada
        boton_cerrar = tk.Button(ventana_conexiones, text="Cerrar Conexión Seleccionada", command=cerrar_conexion)
        boton_cerrar.pack(pady=10)

        # Botón para cerrar la ventana
        boton_cerrar_ventana = tk.Button(ventana_conexiones, text="Cerrar", command=ventana_conexiones.destroy)
        boton_cerrar_ventana.pack(pady=10)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo gestionar las conexiones: {e}")

def recuperar_archivos_eliminados():
    try:
        # Crear un objeto COM para interactuar con la papelera de reciclaje de Windows
        shell = win32com.client.Dispatch("Shell.Application")
        reciclaje = shell.Namespace(10)  # '10' es el ID de la Papelera de reciclaje

        archivos = reciclaje.Items()  # Obtener todos los archivos en la papelera
        if archivos is None or len(archivos) == 0:
            messagebox.showinfo("Papelera vacía", "No hay archivos en la papelera de reciclaje.")
            return

        # Crear ventana para mostrar los archivos eliminados
        ventana_recuperacion = tk.Toplevel()
        ventana_recuperacion.title("Archivos en la Papelera de Reciclaje")
        ventana_recuperacion.geometry("500x300")

        # Mostrar los archivos en un Text widget
        text_archivos = tk.Text(ventana_recuperacion, height=12, width=60)
        text_archivos.pack(padx=10, pady=10)

        for archivo in archivos:
            text_archivos.insert(tk.END, f"{archivo.Name}\n")

        # Función para restaurar el archivo seleccionado
        def restaurar_archivo():
            try:
                seleccion = text_archivos.get(tk.SEL_FIRST, tk.SEL_LAST)
                if seleccion:
                    # Recuperar el archivo seleccionado de la papelera
                    for archivo in archivos:
                        if archivo.Name.strip() == seleccion.strip():
                            archivo.InvokeVerb("restaurar")  # Restaurar el archivo
                            messagebox.showinfo("Restauración exitosa", f"El archivo {seleccion} ha sido restaurado.")
                            return
                    messagebox.showwarning("Error", "No se encontró el archivo seleccionado.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo restaurar el archivo: {e}")

        # Botón para restaurar el archivo seleccionado
        boton_restaurar = tk.Button(ventana_recuperacion, text="Restaurar Archivo Seleccionado", command=restaurar_archivo)
        boton_restaurar.pack(pady=10)

        # Botón para cerrar la ventana
        boton_cerrar_ventana = tk.Button(ventana_recuperacion, text="Cerrar", command=ventana_recuperacion.destroy)
        boton_cerrar_ventana.pack(pady=10)

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo acceder a la papelera de reciclaje: {e}")

def eliminar_virus():
    try:
        subprocess.run(['powershell', 'Start-MpScan', '-ScanType', 'QuickScan'], check=True)
        print("Escaneo en curso...")
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar el escaneo: {e}")

def configurar_microfono():
    try:
        # Abrir la configuración de sonido directamente en la sección del micrófono (Windows 10/11)
        subprocess.run(["start", "ms-settings:sound"], shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir la configuración del micrófono: {e}")

def ejecutar_dism():
    try:
        log_path = os.path.join(os.getcwd(), "dism_log.txt")
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Construir comandos con redirección al archivo de log
        comando = (
            f'echo --- DISM Log - {timestamp} --- >> "{log_path}" && '
            f'DISM /Online /Cleanup-Image /CheckHealth >> "{log_path}" && '
            f'echo. >> "{log_path}" && '
            f'DISM /Online /Cleanup-Image /ScanHealth >> "{log_path}" && '
            f'echo. >> "{log_path}" && '
            f'DISM /Online /Cleanup-Image /RestoreHealth >> "{log_path}" && '
            f'echo. >> "{log_path}" && '
            f'echo --- FIN DEL LOG --- >> "{log_path}" && '
            f'notepad "{log_path}"'
        )

        # Ejecutar en CMD visible
        subprocess.Popen(f'start cmd /k {comando}', shell=True)
        messagebox.showinfo("DISM", f"DISM ejecutado y log guardado en:\n{log_path}")
        
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar DISM: {e}")

def ver_contraseña_wifi():
    try:
        # Ejecutar el comando para obtener los perfiles Wi-Fi guardados
        result = subprocess.run("netsh wlan show profiles", shell=True, capture_output=True, text=True)
        profiles = [line.split(":")[1][1:-1] for line in result.stdout.splitlines() if "All User Profile" in line]

        if not profiles:
            messagebox.showwarning("Sin perfiles", "No hay redes Wi-Fi guardadas.")
            return

        # Mostrar contraseñas para cada perfil Wi-Fi
        contraseñas = []
        for profile in profiles:
            command = f"netsh wlan show profile name=\"{profile}\" key=clear"
            result = subprocess.run(command, shell=True, capture_output=True, text=True)
            for line in result.stdout.splitlines():
                if "Key Content" in line:
                    contraseñas.append(f"{profile}: {line.split(':')[1][1:]}")

        if contraseñas:
            contraseñas_info = "\n".join(contraseñas)
            messagebox.showinfo("Contraseñas Wi-Fi", contraseñas_info)
        else:
            messagebox.showinfo("Sin contraseña", "No se pudo recuperar la contraseña de la red Wi-Fi.")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo obtener la contraseña Wi-Fi: {e}")

def mostrar_trafico_red():
    window = tk.Tk()
    window.title("Tráfico de Red en Tiempo Real")

    # Label para mostrar tráfico
    label = tk.Label(window, text="Cargando tráfico de red...", font=("Arial", 12))
    label.pack(padx=10, pady=10)

    def actualizar_trafico():
        try:
            net_io = psutil.net_io_counters()
            texto_trafico = f"Bytes enviados: {net_io.bytes_sent} bytes\n"
            texto_trafico += f"Bytes recibidos: {net_io.bytes_recv} bytes\n"
            texto_trafico += f"Paquetes enviados: {net_io.packets_sent}\n"
            texto_trafico += f"Paquetes recibidos: {net_io.packets_recv}"

            label.config(text=texto_trafico)  # Actualizar el texto del label

            window.after(1000, actualizar_trafico)  # Actualizar cada 1 segundo

        except Exception as e:
            label.config(text=f"Error: {e}")
    
    # Llamar a la función para iniciar la actualización en tiempo real
    actualizar_trafico()

def organizar_archivos():
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta a organizar")
    if not carpeta:
        return

    modo = simpledialog.askstring("Modo de organización", "¿Cómo deseas organizar los archivos?\n\nOpciones:\n1. tipo\n2. fecha\n3. tamaño\n\n(Escribe solo la palabra)")
    if not modo:
        return

    archivos_organizados = 0

    if modo.lower() == "tipo":
        tipos_archivos = {
            "Imágenes": [".png", ".jpg", ".jpeg", ".gif", ".bmp"],
            "Documentos": [".pdf", ".doc", ".docx", ".txt", ".xls", ".xlsx"],
            "Videos": [".mp4", ".avi", ".mov", ".mkv"],
            "Música": [".mp3", ".wav", ".aac", ".ogg"],
            "Comprimidos": [".zip", ".rar", ".7z", ".tar", ".gz"],
            "Ejecutables": [".exe", ".msi", ".bat"]
        }

        for archivo in os.listdir(carpeta):
            ruta = os.path.join(carpeta, archivo)
            if os.path.isfile(ruta):
                ext = os.path.splitext(archivo)[1].lower()
                for categoria, extensiones in tipos_archivos.items():
                    if ext in extensiones:
                        destino = os.path.join(carpeta, categoria)
                        os.makedirs(destino, exist_ok=True)
                        shutil.move(ruta, os.path.join(destino, archivo))
                        archivos_organizados += 1
                        break

    elif modo.lower() == "fecha":
        for archivo in os.listdir(carpeta):
            ruta = os.path.join(carpeta, archivo)
            if os.path.isfile(ruta):
                fecha = datetime.fromtimestamp(os.path.getmtime(ruta)).strftime("%Y-%m")
                destino = os.path.join(carpeta, fecha)
                os.makedirs(destino, exist_ok=True)
                shutil.move(ruta, os.path.join(destino, archivo))
                archivos_organizados += 1

    elif modo.lower() == "tamaño":
        for archivo in os.listdir(carpeta):
            ruta = os.path.join(carpeta, archivo)
            if os.path.isfile(ruta):
                size_mb = os.path.getsize(ruta) / (1024 * 1024)
                if size_mb < 1:
                    grupo = "Pequeños (<1MB)"
                elif size_mb < 100:
                    grupo = "Medianos (1MB-100MB)"
                else:
                    grupo = "Grandes (>100MB)"
                destino = os.path.join(carpeta, grupo)
                os.makedirs(destino, exist_ok=True)
                shutil.move(ruta, os.path.join(destino, archivo))
                archivos_organizados += 1

    else:
        messagebox.showwarning("Opción inválida", "Debes escribir: tipo, fecha o tamaño")
        return

    messagebox.showinfo("Organización completada", f"Se organizaron {archivos_organizados} archivos por {modo}.")

def desactivar_telemetria():
    try:
        # Comprobar la configuración actual de telemetría
        telemetria_activa = subprocess.check_output(
            'reg query "HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" /v AllowTelemetry',
            shell=True, stderr=subprocess.STDOUT).decode()

        # Si la telemetría ya está desactivada (valor 0), mostrar un mensaje
        if "0x0" in telemetria_activa:
            messagebox.showinfo("Telemetría", "La telemetría ya está desactivada.")
        else:
            # Desactivar la telemetría estableciendo el valor a 0
            subprocess.run(
                'reg add "HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" /v AllowTelemetry /t REG_DWORD /d 0 /f',
                shell=True)
            messagebox.showinfo("Telemetría", "Telemetría desactivada exitosamente.")

    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"No se pudo verificar o cambiar la telemetría: {e}")
    except Exception as e:
        messagebox.showerror("Error desconocido", f"Ocurrió un error: {e}")

def habilitar_telemetria():
    try:
        # Habilitar la telemetría estableciendo el valor a 1
        subprocess.run(
            'reg add "HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" /v AllowTelemetry /t REG_DWORD /d 1 /f',
            shell=True)
        messagebox.showinfo("Telemetría", "Telemetría habilitada exitosamente.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"No se pudo habilitar la telemetría: {e}")
    except Exception as e:
        messagebox.showerror("Error desconocido", f"Ocurrió un error: {e}")

def reparar_microsoft_store():
    try:
        subprocess.run('powershell -Command "Get-AppxPackage *WindowsStore* | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\"$($_.InstallLocation)\\AppXManifest.xml\\"}"', shell=True)
        messagebox.showinfo("Microsoft Store", "Microsoft Store ha sido reparada correctamente.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo reparar Microsoft Store:\n{str(e)}")
    
# Función para activar el modo juego
def activar_modo_juego():
    try:
        # Desactivar notificaciones
        subprocess.run('powershell.exe -Command "Add-Type -TypeDefinition \"using System; using System.Runtime.InteropServices; public class DisableNotification { [DllImport(\\"user32.dll\\", SetLastError = true)] public static extern int LockWorkStation(); }\"; [DisableNotification]::LockWorkStation()"', shell=True)
        
        # Desactivar OneDrive (evitar sincronización)
        subprocess.run('taskkill /F /IM OneDrive.exe', shell=True)
        
        # Detener la indexación de Windows (buscar archivos)
        subprocess.run('powershell.exe -Command "Stop-Service -Name WSearch"', shell=True)
        
        # Cambiar la prioridad de la CPU para procesos de juego (ajuste de rendimiento)
        subprocess.run('wmic process where "name=\'game.exe\'" call setpriority 128', shell=True)
        
        # Desactivar actualizaciones automáticas de Windows
        subprocess.run('powershell.exe -Command "Set-Service -Name wuauserv -StartupType Disabled"', shell=True)
        
        # Ajustar la configuración de la pantalla (ejemplo de bajar brillo)
        subprocess.run('powershell.exe -Command "Get-WmiObject -Namespace \\"root/wmi\\" -Class WmiMonitorBrightnessMethods | foreach { $_.WmiSetBrightness(1, 50) }"', shell=True)
        
        # Mostrar mensaje de éxito
        messagebox.showinfo("Modo Juego", "¡Modo juego activado! El rendimiento del sistema ha sido optimizado.")
    except Exception as e:
        messagebox.showerror("Error", f"Hubo un error al activar el modo juego: {e}")

def buscar_archivo_en_pc():
    termino = simpledialog.askstring("Buscar archivo", "Ingresa parte del nombre del archivo:")
    if not termino:
        return

    resultados = []

    def buscar():
        unidades = [f"{unidad}:\\" for unidad in "CDEFGHIJKLMNOPQRSTUVWXYZ" if os.path.exists(f"{unidad}:\\")]

        for unidad in unidades:
            for ruta, carpetas, archivos in os.walk(unidad):
                for archivo in archivos:
                    if termino.lower() in archivo.lower():
                        resultados.append(os.path.join(ruta, archivo))

        if resultados:
            mostrar_resultados(resultados)
        else:
            messagebox.showinfo("Sin resultados", "No se encontraron archivos que coincidan.")

    def mostrar_resultados(lista):
        ventana = Toplevel()
        ventana.title("Archivos encontrados")

        scrollbar = Scrollbar(ventana)
        scrollbar.pack(side=RIGHT, fill=Y)

        listbox = Listbox(ventana, yscrollcommand=scrollbar.set, width=100, height=30)
        for archivo in lista[:100]:  # mostrar máximo 100 archivos
            listbox.insert(END, archivo)
        listbox.pack(side=LEFT, fill=BOTH)

        scrollbar.config(command=listbox.yview)

        def abrir_archivo(event):
            seleccionado = listbox.get(listbox.curselection())
            try:
                os.startfile(seleccionado)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

        listbox.bind("<Double-Button-1>", abrir_archivo)

    threading.Thread(target=buscar).start()

# Función para proteger el USB con contraseña
def proteger_usb_contraseña():
    unidad = simpledialog.askstring("Letra de unidad", "Ingresa la letra de la unidad USB (ej. E):")
    contraseña = simpledialog.askstring("Contraseña", "Ingresa una contraseña:", show="*")

    if unidad and contraseña:
        try:
            comando = f'manage-bde -on {unidad}: -password -pw {contraseña} -usedspaceonly'
            resultado = subprocess.run(comando, shell=True, capture_output=True, text=True)
            
            if resultado.returncode == 0:
                messagebox.showinfo("Éxito", "BitLocker fue activado correctamente.")
            else:
                messagebox.showerror("Error", f"No se pudo activar BitLocker:\n{resultado.stderr}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    else:
        messagebox.showwarning("Datos incompletos", "Debes ingresar ambos campos.")

# Función para desactivar BitLocker
def desactivar_bitlocker():
    unidad = simpledialog.askstring("Letra de unidad", "Ingresa la letra de la unidad USB (ej. E):")

    if unidad:
        try:
            comando = f'manage-bde -off {unidad}:'
            resultado = subprocess.run(comando, shell=True, capture_output=True, text=True)

            if resultado.returncode == 0:
                messagebox.showinfo("Éxito", "BitLocker fue desactivado correctamente.")
            else:
                messagebox.showerror("Error", f"No se pudo desactivar BitLocker:\n{resultado.stderr}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    else:
        messagebox.showwarning("Datos incompletos", "Debes ingresar la letra de la unidad.")

def eliminar_malware_persistente():
    try:
        print("[+] Iniciando eliminación de malware persistente...")

        # 1. Detener procesos sospechosos (ejemplo: kill procesos conocidos)
        procesos_sospechosos = ["malware.exe", "virus.exe"]
        for proc in procesos_sospechosos:
            subprocess.run(f"taskkill /f /im {proc}", shell=True)

        # 2. Eliminar archivos y carpetas sospechosas
        rutas_sospechosas = [
            r"%APPDATA%\Malicious",
            r"C:\Users\%USERNAME%\AppData\Local\Temp\malware",
            r"C:\malware_temp"
        ]
        for ruta in rutas_sospechosas:
            ruta_exp = os.path.expandvars(ruta)
            subprocess.run(f"attrib -r -s -h \"{ruta_exp}\"", shell=True)
            subprocess.run(f"rmdir /s /q \"{ruta_exp}\"", shell=True)

        # 3. Escanear con Windows Defender
        subprocess.run("start cmd /k \"cd %ProgramFiles%\\Windows Defender & MpCmdRun.exe -Scan -ScanType 2\"", shell=True)

        print("[✔] Eliminación de malware iniciada.")
    
    except Exception as e:
        print(f"[!] Error: {e}")

def escaneo_avanzado_defender():
    try:
        print("[🔍] Iniciando escaneo completo con Windows Defender...")
        subprocess.run(
            r'start cmd /k "cd %ProgramFiles%\Windows Defender & MpCmdRun.exe -Scan -ScanType 2"',
            shell=True
        )
    except Exception as e:
        print(f"[!] Error al ejecutar el escaneo: {e}")

def mostrar_iconos_sistema_escritorio():
    try:
        print("[🖥️] Habilitando íconos del sistema en el escritorio...")

        # Activar íconos estándar: Mi PC (Este equipo), Red, Papelera, Panel de control
        iconos = {
            "Computer": "1",
            "Network": "1",
            "RecycleBin": "1",
            "ControlPanel": "1"
        }

        for icono, valor in iconos.items():
            subprocess.run([
                "reg", "add",
                r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel",
                "/v", icono,
                "/t", "REG_DWORD",
                "/d", "0",  # 0 = mostrar
                "/f"
            ], shell=True)

        # Refrescar el escritorio (opcional)
        subprocess.run("ie4uinit.exe -show", shell=True)
        subprocess.run("RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", shell=True)

        print("[✔] Iconos del sistema habilitados en el escritorio.")

    except Exception as e:
        print(f"[!] Error: {e}")

def mostrar_archivos_ocultos():
    try:
        subprocess.run('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v Hidden /t REG_DWORD /d 1 /f', shell=True)
        subprocess.run('taskkill /f /im explorer.exe', shell=True)
        subprocess.run('start explorer.exe', shell=True)
        messagebox.showinfo("Éxito", "Archivos ocultos ahora son visibles.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def crear_carpeta_personalizada():
    ruta = simpledialog.askstring("Crear Carpeta", "Ingrese la ruta donde desea crear la carpeta:")
    nombre = simpledialog.askstring("Crear Carpeta", "Nombre de la nueva carpeta:")
    if ruta and nombre:
        try:
            os.makedirs(os.path.join(ruta, nombre), exist_ok=True)
            messagebox.showinfo("Éxito", f"Carpeta '{nombre}' creada en {ruta}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

def buscar_archivos_grandes():
    from tkinter import filedialog
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta para buscar archivos grandes")
    if carpeta:
        grandes = []
        for root, _, files in os.walk(carpeta):
            for f in files:
                ruta = os.path.join(root, f)
                tamaño = os.path.getsize(ruta)
                if tamaño > 100 * 1024 * 1024:  # mayor a 100 MB
                    grandes.append((ruta, tamaño // (1024*1024)))
        if grandes:
            resultado = "\n".join([f"{ruta} - {tam} MB" for ruta, tam in grandes])
            messagebox.showinfo("Archivos grandes encontrados", resultado)
        else:
            messagebox.showinfo("Resultado", "No se encontraron archivos mayores a 100 MB.")
    
def buscar_sectores_dañados():
    unidad = "C:"  # Puedes modificar la letra de unidad si deseas permitir selección
    respuesta = messagebox.askyesno("Buscar sectores dañados", f"¿Deseas escanear la unidad {unidad} con CHKDSK?\nEsto podría requerir reiniciar el sistema.")
    if respuesta:
        try:
            os.system(f"chkdsk {unidad} /f /r")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo ejecutar CHKDSK:\n{str(e)}")

def comprobar_salud_ssd():
    try:
        # Ejecutar comando WMIC para obtener el estado SMART del disco
        resultado = subprocess.check_output("wmic diskdrive get model,status", shell=True, text=True)
        messagebox.showinfo("Salud del SSD", f"Resultado:\n{resultado}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo obtener la información SMART:\n{str(e)}")

def fusionar_espacio_libre():
    try:
        # Abrir herramienta de administración de discos de Windows
        os.system("diskmgmt.msc")
        messagebox.showinfo(
            "Fusión de espacio libre",
            "La herramienta de administración de discos se ha abierto.\n\n"
            "Para fusionar espacio libre:\n1. Elimina una partición contigua.\n2. Haz clic derecho en la principal.\n3. Elige 'Extender volumen...'."
        )
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir la herramienta de discos:\n{str(e)}")

def borrado_seguro_archivos():
    archivo = filedialog.askopenfilename(title="Selecciona un archivo para borrar de forma segura")
    if not archivo:
        return

    try:
        tamano = os.path.getsize(archivo)

        # Sobrescribir con ceros
        with open(archivo, "ba+", buffering=0) as f:
            f.seek(0)
            f.write(b"\x00" * tamano)

        # Eliminar archivo
        os.remove(archivo)
        messagebox.showinfo("Éxito", "Archivo borrado de forma segura.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo borrar el archivo:\n{str(e)}")

def calcular_uso_por_carpeta():
    carpeta = filedialog.askdirectory(title="Selecciona una carpeta")
    if not carpeta:
        return

    resultados = []
    try:
        for subdir in os.listdir(carpeta):
            ruta_subcarpeta = os.path.join(carpeta, subdir)
            if os.path.isdir(ruta_subcarpeta):
                tamano_total = 0
                for carpeta_raiz, _, archivos in os.walk(ruta_subcarpeta):
                    for archivo in archivos:
                        try:
                            ruta_archivo = os.path.join(carpeta_raiz, archivo)
                            tamano_total += os.path.getsize(ruta_archivo)
                        except:
                            continue
                resultados.append(f"{subdir}: {round(tamano_total / (1024**2), 2)} MB")
        
        if resultados:
            resultado_str = "\n".join(resultados)
            messagebox.showinfo("Tamaño por subcarpeta", resultado_str)
        else:
            messagebox.showinfo("Resultado", "No se encontraron subcarpetas.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def borrar_archivos_grandes():
    carpeta = filedialog.askdirectory(title="Selecciona una carpeta")
    if not carpeta:
        return

    tamaño_umbral_mb = 100  # Puedes cambiar este umbral
    archivos_grandes = []

    try:
        for carpeta_raiz, _, archivos in os.walk(carpeta):
            for archivo in archivos:
                ruta_archivo = os.path.join(carpeta_raiz, archivo)
                try:
                    tamaño_bytes = os.path.getsize(ruta_archivo)
                    tamaño_mb = tamaño_bytes / (1024 ** 2)
                    if tamaño_mb > tamaño_umbral_mb:
                        archivos_grandes.append((ruta_archivo, tamaño_mb))
                except:
                    continue

        if not archivos_grandes:
            messagebox.showinfo("Resultado", f"No se encontraron archivos mayores a {tamaño_umbral_mb} MB.")
            return

        confirmacion = messagebox.askyesno(
            "Confirmar borrado",
            f"Se encontraron {len(archivos_grandes)} archivos mayores a {tamaño_umbral_mb} MB.\n¿Deseas eliminarlos?"
        )

        if confirmacion:
            for ruta, _ in archivos_grandes:
                try:
                    os.remove(ruta)
                except:
                    continue
            messagebox.showinfo("Éxito", "Archivos grandes eliminados correctamente.")
        else:
            messagebox.showinfo("Cancelado", "No se eliminó ningún archivo.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def generar_reporte_disco():
    try:
        sistema = platform.system()
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        reporte = f"🧾 REPORTE DE USO DE DISCO - {fecha}\n\n"

        # En Windows, escanea desde la A hasta la Z
        if sistema == "Windows":
            for letra in range(65, 91):  # A-Z
                unidad = f"{chr(letra)}:\\"
                if os.path.exists(unidad):
                    total, usado, libre = shutil.disk_usage(unidad)
                    reporte += f"Unidad {unidad}\n"
                    reporte += f"  Total: {total // (2**30)} GB\n"
                    reporte += f"  Usado: {usado // (2**30)} GB\n"
                    reporte += f"  Libre: {libre // (2**30)} GB\n\n"
        else:  # Para Linux/Mac
            unidades = ["/"]
            for unidad in unidades:
                total, usado, libre = shutil.disk_usage(unidad)
                reporte += f"Unidad {unidad}\n"
                reporte += f"  Total: {total // (2**30)} GB\n"
                reporte += f"  Usado: {usado // (2**30)} GB\n"
                reporte += f"  Libre: {libre // (2**30)} GB\n\n"

        # Guardar el reporte
        with open("reporte_uso_disco.txt", "w", encoding="utf-8") as f:
            f.write(reporte)

        messagebox.showinfo("Reporte generado", "El reporte de uso de disco ha sido guardado como:\nreporte_uso_disco.txt")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al generar el reporte:\n{e}")

def buscar_archivos_comprimidos():
    try:
        # Selección de carpeta raíz para la búsqueda
        carpeta = filedialog.askdirectory(title="Selecciona la carpeta donde buscar archivos comprimidos")
        if not carpeta:
            return

        # Extensiones de archivos comprimidos comunes
        extensiones_comprimidos = ('.zip', '.rar', '.7z', '.tar', '.gz', '.bz2', '.xz')

        encontrados = []

        for dirpath, _, archivos in os.walk(carpeta):
            for archivo in archivos:
                if archivo.lower().endswith(extensiones_comprimidos):
                    ruta_completa = os.path.join(dirpath, archivo)
                    encontrados.append(ruta_completa)

        if encontrados:
            # Guardar resultados en un archivo
            with open("archivos_comprimidos_encontrados.txt", "w", encoding="utf-8") as f:
                f.write("📦 Archivos comprimidos encontrados:\n\n")
                for archivo in encontrados:
                    f.write(f"{archivo}\n")

            messagebox.showinfo("Completado", f"Se encontraron {len(encontrados)} archivos comprimidos.\nGuardado en 'archivos_comprimidos_encontrados.txt'")
        else:
            messagebox.showinfo("Sin resultados", "No se encontraron archivos comprimidos en la carpeta seleccionada.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error durante la búsqueda:\n{e}")

def detener_procesos_pesados():
    try:
        # Umbral en MB/s (puedes ajustar)
        umbral_disco = 1  # procesos con uso de disco > 1MB/s

        procesos_terminados = []
        procesos_examinados = []

        for proc in psutil.process_iter(['pid', 'name']):
            try:
                pid = proc.info['pid']
                nombre = proc.info['name']
                io = proc.io_counters()

                # Leer los bytes leídos y escritos
                uso_disco = (io.read_bytes + io.write_bytes) / (1024 * 1024)  # en MB

                if uso_disco > umbral_disco:
                    procesos_examinados.append((nombre, pid, round(uso_disco, 2)))
                    proc.terminate()
                    procesos_terminados.append(f"{nombre} (PID: {pid}) - {round(uso_disco, 2)} MB")
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        if procesos_terminados:
            resumen = "\n".join(procesos_terminados)
            messagebox.showinfo("Procesos detenidos", f"Se detuvieron los siguientes procesos:\n\n{resumen}")
        else:
            messagebox.showinfo("Sin procesos pesados", "No se detectaron procesos que consuman mucho disco.")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al detener procesos:\n{e}")

def mover_archivos_grandes(tamano_minimo_mb=100):
    try:
        # Selección de carpeta origen
        origen = filedialog.askdirectory(title="Seleccionar carpeta origen")
        if not origen:
            return

        # Selección de carpeta destino
        destino = filedialog.askdirectory(title="Seleccionar carpeta destino")
        if not destino:
            return

        archivos_movidos = []
        tamano_minimo_bytes = tamano_minimo_mb * 1024 * 1024

        for root, dirs, files in os.walk(origen):
            for archivo in files:
                ruta_archivo = os.path.join(root, archivo)
                try:
                    if os.path.getsize(ruta_archivo) >= tamano_minimo_bytes:
                        nuevo_path = os.path.join(destino, archivo)
                        shutil.move(ruta_archivo, nuevo_path)
                        archivos_movidos.append(archivo)
                except Exception as e:
                    print(f"Error con {archivo}: {e}")

        if archivos_movidos:
            messagebox.showinfo("Archivos movidos",
                                f"Se movieron {len(archivos_movidos)} archivos grandes a:\n{destino}")
        else:
            messagebox.showinfo("Sin archivos",
                                f"No se encontraron archivos mayores a {tamano_minimo_mb} MB en:\n{origen}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al mover archivos:\n{e}")

def optimizar_unidades():
    sistema = platform.system()
    
    if sistema != "Windows":
        messagebox.showwarning("No compatible", "Esta función solo está disponible en Windows.")
        return

    try:
        # Ejecuta el optimizador de unidades con privilegios normales
        resultado = subprocess.run(["dfrgui"], shell=True)
        if resultado.returncode == 0:
            messagebox.showinfo("Optimización", "La herramienta de optimización de unidades fue abierta.")
        else:
            messagebox.showwarning("Error", "No se pudo abrir la herramienta de optimización de discos.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al intentar abrir la herramienta:\n{e}")

def ver_firewall_reglas():
    try:
        resultado = subprocess.check_output(
            'netsh advfirewall firewall show rule name=all',
            shell=True, text=True, stderr=subprocess.STDOUT
        )

        # Mostrar en ventana con scroll
        ventana_resultado = tk.Toplevel()
        ventana_resultado.title("Reglas del Firewall")
        ventana_resultado.geometry("800x500")

        texto = scrolledtext.ScrolledText(ventana_resultado, wrap=tk.WORD, bg="#f8f9fa")
        texto.insert(tk.END, resultado)
        texto.pack(expand=True, fill="both")

    except subprocess.CalledProcessError as e:
        tk.messagebox.showerror("Error", f"No se pudo obtener reglas del firewall:\n{e.output}")

def renovar_ip_dns():
    try:
        # Liberar IP actual
        subprocess.run("ipconfig /release", shell=True, check=True)
        # Renovar IP
        subprocess.run("ipconfig /renew", shell=True, check=True)
        # Vaciar caché DNS
        subprocess.run("ipconfig /flushdns", shell=True, check=True)
        # Registrar de nuevo el DNS
        subprocess.run("ipconfig /registerdns", shell=True, check=True)

        messagebox.showinfo("Completado", "IP y DNS renovados correctamente.")

    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Ocurrió un error al renovar IP/DNS:\n{e}")

def ver_redes_disponibles():
    try:
        # Ejecutar comando para obtener las redes Wi-Fi disponibles
        result = subprocess.run("netsh wlan show networks", shell=True, check=True, text=True, capture_output=True)
        
        # Extraer la salida del comando
        redes = result.stdout

        # Mostrar las redes en una ventana emergente
        if redes:
            messagebox.showinfo("Redes Disponibles", redes)
        else:
            messagebox.showinfo("Sin Redes", "No se encontraron redes Wi-Fi disponibles.")

    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Ocurrió un error al obtener las redes:\n{e}")

def hacer_ping_manual():
    # Solicitar al usuario la dirección IP o dominio
    direccion = simpledialog.askstring("Ping Manual", "Ingrese la dirección IP o dominio a pingear:")
    
    if direccion:
        try:
            # Ejecutar el comando ping (4 intentos)
            resultado = subprocess.run(["ping", "-n", "4", direccion], capture_output=True, text=True, shell=True)
            salida = resultado.stdout

            # Mostrar resultado en ventana
            messagebox.showinfo("Resultado del Ping", salida)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo hacer ping:\n{e}")
    else:
        messagebox.showinfo("Ping cancelado", "No se ingresó ninguna dirección.")

def apagar_sistema():
    sistema = platform.system()
    if sistema == "Windows":
        os.system("shutdown /s /t 1")
    elif sistema == "Linux" or sistema == "Darwin":  # Darwin = macOS
        os.system("sudo shutdown now")
    else:
        print("Sistema operativo no compatible.")

def activar_modo_no_molestar():
    if platform.system() == "Windows":
        comando = (
            'powershell -Command "New-ItemProperty -Path '
            "'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Notifications\\Settings' "
            '-Name NOC_GLOBAL_SETTING_TOASTS_ENABLED -Value 0 -PropertyType DWORD -Force"'
        )
        subprocess.run(comando, shell=True)
        print("Modo 'No molestar' activado en Windows.")
    else:
        print("Función no disponible en este sistema.")

def forzar_reinicio():
    sistema = platform.system()
    try:
        if sistema == "Windows":
            # Reinicio inmediato en Windows (fuerza cierre de aplicaciones)
            os.system("shutdown /r /f /t 0")
        elif sistema == "Linux" or sistema == "Darwin":  # Darwin = macOS
            # Requiere privilegios de administrador
            subprocess.run(["sudo", "reboot"], check=True)
        else:
            print("Sistema no compatible para reinicio forzado.")
    except Exception as e:
        print(f"Error al intentar reiniciar: {e}")

def cerrar_sesion():
    sistema = platform.system()
    try:
        if sistema == "Windows":
            # Cierra sesión en Windows
            os.system("shutdown /l")
        elif sistema == "Linux":
            # En algunos entornos de escritorio como GNOME o KDE:
            subprocess.run(["gnome-session-quit", "--logout", "--no-prompt"])
        elif sistema == "Darwin":  # macOS
            # macOS requiere AppleScript o comandos específicos
            subprocess.run([
                "osascript", "-e",
                'tell application "System Events" to log out'
            ])
        else:
            print("Sistema no compatible para cerrar sesión.")
    except Exception as e:
        print(f"Error al cerrar sesión: {e}")

def cambiar_tipo_cuenta():
    usuario = simpledialog.askstring("Cambiar tipo de cuenta", "Ingresa el nombre del usuario:")
    if not usuario:
        return

    tipo = simpledialog.askstring("Tipo de cuenta", "Escribe 'admin' para Administrador o 'estandar' para Estándar:")
    if not tipo:
        return

    try:
        if tipo.lower() == "admin":
            comando = ["net", "localgroup", "Administradores", usuario, "/add"]
        elif tipo.lower() == "estandar":
            comando = ["net", "localgroup", "Administradores", usuario, "/delete"]
        else:
            messagebox.showerror("Error", "Tipo de cuenta no válido. Usa 'admin' o 'estandar'.")
            return

        resultado = subprocess.run(comando, capture_output=True, text=True, shell=True)
        if resultado.returncode == 0:
            messagebox.showinfo("Éxito", f"El tipo de cuenta del usuario '{usuario}' fue cambiado correctamente.")
        else:
            messagebox.showerror("Error", f"No se pudo cambiar el tipo de cuenta:\n{resultado.stderr}")

    except Exception as e:
        messagebox.showerror("Excepción", str(e))
# === Interfaz gráfica ===
ventana = tk.Tk()
ventana.title("Reparador del Sistema Operativo")
ventana.geometry("600x400")
ventana.configure(bg="#f0f4f7")

# Cargar y redimensionar la imagen
imagen = Image.open("repartidoe.jpg")
imagen = imagen.resize((1500, 400))  # Tamaño igual al de la ventana
imagen_tk = ImageTk.PhotoImage(imagen)

# Crear Label con imagen de fondo
fondo_label = tk.Label(ventana, image=imagen_tk)
fondo_label.place(x=0, y=0, relwidth=1, relheight=1)

titulo = tk.Label(ventana, text="🛠️ Reparador del Sistema Operativo", font=("Helvetica", 14, "bold"), bg="#f0f4f7")
titulo.pack(pady=10)

# Crear el Notebook con fondo transparente
notebook = ttk.Notebook(ventana)
notebook.place(x=0, y=70, relwidth=1, relheight=1)

notebook = ttk.Notebook(ventana)
frame_Gestion_de_Archivo_y_Carpetas = tk.Frame(notebook, bg="#f0f4f7")
frame_Disco_y_Almacenamiento = tk.Frame(notebook, bg="#f0f4f7")
frame_Red_e_Internet = tk.Frame(notebook, bg="#f0f4f7")
frame_Seguridad_Antivirus_y_Permisos = tk.Frame(notebook, bg="#f0f4f7")
frame_Sistema_y_Rendimiento=tk.Frame(notebook, bg="#f0f4f7")
frame_Memoria_Recurso_y_Registro = tk.Frame(notebook, bg="#f0f4f7")
frame_Usuarios_y_Configuraciones_Personal = tk.Frame(notebook, bg="#f0f4f7")
frame_Interfaz_Visual_y_Escritorio = tk.Frame(notebook, bg="#f0f4f7")
frame_Acciones_Rapidas_Reinicio = tk.Frame(notebook, bg="#f0f4f7")
notebook.add(frame_Gestion_de_Archivo_y_Carpetas, text="Gestion de Archivo y Carpetas")
notebook.add(frame_Disco_y_Almacenamiento, text="Disco y Almacenamiento")
notebook.add(frame_Red_e_Internet, text="Red e Internet")
notebook.add(frame_Seguridad_Antivirus_y_Permisos, text="Seguridad Antivirus y Permisos" )
notebook.add(frame_Sistema_y_Rendimiento,text="Sistema_y_Rendimiento")
notebook.add(frame_Memoria_Recurso_y_Registro, text="Memoria Recurso y Registro")
notebook.add(frame_Usuarios_y_Configuraciones_Personal , text="Usuarios y Configuraciones Personal")
notebook.add(frame_Interfaz_Visual_y_Escritorio, text="Interfaz Visual y Escritorio")
notebook.add(frame_Acciones_Rapidas_Reinicio, text="Acciones Rapidas Reinicio")
notebook.pack(expand=True, fill="both")

def agregar_botones(frame, botones):
    for texto, comando, color in botones:
        tk.Button(frame, text=texto, font=("Helvetica", 10), bg=color, fg="white",
                  command=comando, height=1, width=40, relief="flat").pack(pady=5)

botones_Gestion_de_Archivo_y_Carpetas = [
    ("🗑️ Vaciar papelera", vaciar_papelera, "#e67e22"),
    ("📂 Eliminar logs", eliminar_archivos_registro, "#9b59b6"),
    ("🗂️ Eliminar duplicados", eliminar_duplicados_en_directorio, "#c0392b"),
    ("📁 Mostrar archivos ocultos", mostrar_archivos_ocultos, "#2980b9"),
    ("📁 Crear carpeta personalizada", crear_carpeta_personalizada, "#27ae60"),
    ("🖼️ Eliminar imágenes duplicadas", eliminar_duplicados_imagenes, "#9b59b6"),
    ("🗂️ Organizar archivos", organizar_archivos, "#f39c12"),
    ("🗑️ Recuperar Archivos Eliminados", recuperar_archivos_eliminados, "#f39c12"),
    ("🧭 Buscar archivo en PC (abrir resultados)", buscar_archivo_en_pc, "#2c3e50"),
    ("🔍 Buscar archivos grandes", buscar_archivos_grandes, "#e67e22"),
    ("📦 Buscar archivos comprimidos", buscar_archivos_comprimidos, "#f39c12"),
    ("🧹 Borrar archivos basura grandes", borrar_archivos_grandes, "#d35400"),
    ("💣 Borrado seguro de archivos", borrado_seguro_archivos, "#c0392b"),
    ("🧮 Calcular uso por carpeta", calcular_uso_por_carpeta, "#34495e"),
    ("📤 Mover archivos grandes a otra unidad", mover_archivos_grandes, "#f1c40f"), 
]

botones_Disco_y_Almacenamiento = [
    ("🧹 Limpiar archivos temporales", limpiar_temporales, "#27ae60"),
    ("💾 Desfragmentar disco", desfragmentar_disco, "#34495e"),
    ("💽 Ejecutar CHKDSK", verificar_chkdsk, "#2980b9"),
    ("📊 Estado SMART", obtener_estado_smart, "#2ecc71"),
    ("🧮 Espacio libre en disco", espacio_libre_disco, "#27ae60"),
    ("🌡️ Monitorear temperatura CPU", monitorear_temperatura, "#9b59b6"),
    ("🔍 Buscar sectores dañados", buscar_sectores_dañados, "#e74c3c"),
    ("🧪 Comprobar salud SSD", comprobar_salud_ssd, "#8e44ad"),
    ("🔃 Fusionar espacio libre", fusionar_espacio_libre, "#27ae60"),
    ("🧾 Generar reporte de uso de disco", generar_reporte_disco, "#1abc9c"),
    ("🛑 Detener procesos que consumen mucho disco", detener_procesos_pesados, "#e74c3c"),
    ("🛠️ Optimizar unidades", optimizar_unidades, "#2980b9"), 

]

botones_Red_e_Internet = [
    ("📡 Ver redes disponibles", ver_redes_disponibles, "#9b59b6"),
    ("📡 Pingar dirección específica", hacer_ping_manual, "#2ecc71"), 
    ("🌐 Verificar conexión", verificar_conexion, "#16a085"),
    ("📶 Velocidad de Internet", verificar_velocidad_internet, "#16a085"),
    ("🌐 Ver tráfico de red", mostrar_trafico_red, "#3498db"),
    ("🌐 Verificar puertos abiertos", verificar_puertos_abiertos, "#2980b9"),
    ("🌐 Renovar IP y DNS", renovar_ip_dns, "#1abc9c"), 
    ("🔌 Administrar conexiones de red", administrar_conexiones_red, "#2980b9"),
    ("🔑 Ver contraseña Wi-Fi", ver_contraseña_wifi, "#8e44ad"),
    ("🧳 Limpiar caché navegadores", limpiar_cache_navegadores, "#f39c12"),
    ("🛡️ Ver firewall y reglas", ver_firewall_reglas, "#e74c3c"),
    
]

botones_Seguridad_Antivirus_y_Permisos = [
    ("🔍 Escanear con Windows Defender", escanear_con_defender, "#2ecc71"),
    ("🔍 Escaneo avanzado con Defender", escaneo_avanzado_defender, "#2ecc71"),
    ("🦠 Eliminar virus", eliminar_virus, "#e74c3c"),
    ("🧼 Eliminar malware persistente",eliminar_malware_persistente,"#e74c3c"),
    ("🔑 Reparar permisos", reparar_permisos, "#8e44ad"),
    ("❌ Desactivar BitLocker", desactivar_bitlocker, "#e74c3c"),
    ("🔒 Proteger USB con Contraseña", proteger_usb_contraseña, "#e74c3c"),  # Botón "Proteger USB"
    ("🧼 Limpiar caché DNS", limpiar_cache_dns, "#e67e22"),
    ("📦 Eliminar programas", eliminar_programas_innecesarios, "#c0392b"),
  
]

botones_Sistema_y_Rendimiento  =[
    ("🛠️ Reparar con SFC", reparar_archivos_sfc, "#8e44ad"),
    ("🧮 Ejecutar DISM", ejecutar_dism, "#1abc9c"),
    ("🛠️ Reparar BOOT", reparar_archivos_boot, "#e74c3c"),
    ("🔧 Desactivar actualizaciones", desactivar_actualizaciones, "#f39c12"),
    ("🔄 Verificar actualizaciones", verificar_actualizaciones_windows, "#16a085"),
    ("🕓 Restaurar sistema", restaurar_sistema, "#e67e22"),
    ("📥 Actualizar Drivers", actualizar_drivers, "#f39c12"),
    ("🪟 Actualizar sistema", actualizar_sistema, "#2980b9"),
    ("✅ Habilitar Telemetría", habilitar_telemetria, "#2ecc71"),
    ("❌ Desactivar Telemetría", desactivar_telemetria, "#e74c3c"),
    ("🎮 Activar modo juego", activar_modo_juego, "#9b59b6"),
    ("🚀 Mejorar rendimiento del sistema", mejorar_rendimiento_sistema, "#e67e22"),
    ("⚡ Optimizar arranque", optimizar_arranque, "#e67e22"),

]

botones_Memoria_Recurso_y_Registro =[
    ("🧠 Verificar memoria RAM", verificar_memoria_ram, "#34495e"),
    ("🧠 Optimizar uso de RAM", optimizar_uso_ram, "#2ecc71"),
    ("💾 Backup registro", copia_seguridad_registro, "#2980b9"),
   
]

botones_Usuarios_Config = [
    ("👤 Crear nuevo usuario", crear_usuario_nuevo, "#27ae60"),
    ("👥 Cambiar tipo de cuenta", cambiar_tipo_cuenta, "#2ecc71"),
    ("🖱️ Configurar periféricos", configurar_perifericos, "#2980b9"),
    ("🧾 Mostrar procesos activos", mostrar_procesos_activoss, "#27ae60"),
    ("🧰 Herramientas administrativas", abrir_herramientas_administrativas, "#34495e"),
    ("🗓 Programar tarea automática", programar_tarea_automatica, "#8e44ad"),
    ("🧭 Comprobar rutas del sistema (PATH)", comprobar_rutas_sistema, "#34495e"),
    ("📤 Exportar lista de programas", exportar_lista_programas, "#1abc9c"),
    ("📦 Desinstalar programas", desinstalar_programas, "#c0392b"),
    ("💾 Realizar copia de seguridad", realizar_copia_seguridad, "#27ae60"),
    ("🔋 Optimizar batería (laptops)", optimizar_bateria, "#f1c40f"),
    ("🧯 Modo seguro con red", modo_seguro_con_red, "#d35400"),
    ("🧰 Reparar Microsoft Store", reparar_microsoft_store, "#e67e22"),
   
]

botones_Interfaz_Visual_y_Escritorio = [
    ("🖥️ Mostrar iconos del sistema en escritorio", mostrar_iconos_sistema_escritorio, "#3498db"),
    ("🎙️ Configuración de micrófono", configurar_microfono, "#3498db"),
    ("📋 Crear reporte diagnóstico", crear_reporte_diagnostico, "#34495e"),
    ("🖥️ Información del sistema", informacion_sistema, "#9b59b6"),
    ("🖼️ Cambiar resolución", cambiar_resolucion, "#9b59b6"),
    ("⚙️ Panel de Control", abrir_panel_control, "#2980b9"),
    ("⚙️ Desactivar programas inicio", desactivar_programas_inicio, "#f1c40f"),

]

botones_Acciones_Rapidas_Reinicio =[
    ("🔄 Reiniciar sistema", reiniciar_sistema, "#16a085"),
    ("🔃 Forzar reinicio inmediato", forzar_reinicio, "#e74c3c"), 
    ("⏻ Apagar sistema", apagar_sistema, "#c0392b"), 
    ("🔕 Activar modo no molestar", activar_modo_no_molestar, "#34495e"), 
    ("🛑 Cerrar sesión", cerrar_sesion, "#e67e22"),
]

agregar_botones(frame_Gestion_de_Archivo_y_Carpetas , botones_Gestion_de_Archivo_y_Carpetas )
agregar_botones(frame_Disco_y_Almacenamiento, botones_Disco_y_Almacenamiento)
agregar_botones(frame_Red_e_Internet, botones_Red_e_Internet)
agregar_botones(frame_Seguridad_Antivirus_y_Permisos, botones_Seguridad_Antivirus_y_Permisos) 
agregar_botones(frame_Sistema_y_Rendimiento, botones_Sistema_y_Rendimiento)
agregar_botones(frame_Memoria_Recurso_y_Registro, botones_Memoria_Recurso_y_Registro)
agregar_botones(frame_Usuarios_y_Configuraciones_Personal, botones_Usuarios_Config)
agregar_botones(frame_Interfaz_Visual_y_Escritorio, botones_Interfaz_Visual_y_Escritorio)
agregar_botones(frame_Acciones_Rapidas_Reinicio, botones_Acciones_Rapidas_Reinicio)


ventana.mainloop()
