import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
import subprocess
import time
import pyautogui
from supabase import create_client, Client
from typing import Optional
import customtkinter as ctk
import threading
import math
from datetime import datetime

# === CONFIGURACIÓN ===
CARPETA_BOTONES = "Buttons"
URL_TECFOOD = "https://food.teknisa.com//df/#/df_entrada#dfe11000_lancamento_entrada"
URL_RETIRADA = "https://food.teknisa.com//est/#/est_relatorios#est31100_posicao_de_estoque"

SUPABASE_URL = "https://ulboklgzjriatmaxzpsi.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVsYm9rbGd6anJpYXRtYXh6cHNpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjI3ODQxNDAsImV4cCI6MjA3ODM2MDE0MH0.gY6_K4JQoJxPZmdXMIbFZfiJAOdavbg8jDJW1rOUSPk"
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# === VARIABLES GLOBALES ===
archivo_excel = None
df_facturas = pd.DataFrame()
productos_por_factura = {}
codigo_clinica = ""
origen_datos = ""
archivo_retirada = None
df_retirada = pd.DataFrame()
unidad_cargada = ""
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
# PALETA DE COLORES

PRIMARY = "#3A7BD5"
PRIMARY_HOVER = "#5EA0FF"
ACCENT = "#FFA726" 
ACCENT_HOVER = "#FFB74D"
CARD_BG = "#1C1E23"
TEXT_MAIN = "#FFFFFF"
TEXT_SECOND = "#B0B0B0"
BG = "#0E0F12"

# SISTEMA DE NOTIFICACIONES OPTIMIZADO

_toast_lock = threading.Lock()
_active_toasts = []

def _position_toast_window(win, width=380, height=96, margin_x=24, margin_y=24, offset_index=0):
    try:
        root_x = root.winfo_rootx()
        root_y = root.winfo_rooty()
        root_w = root.winfo_width()
        root_h = root.winfo_height()
    except Exception:
        root_x, root_y, root_w, root_h = 0, 0, win.winfo_screenwidth(), win.winfo_screenheight()

    x = root_x + max(root_w - width - margin_x, 0)
    y = root_y + max(root_h - height - margin_y - offset_index * (height + 12), 0)
    win.geometry(f"{width}x{height}+{x}+{y}")

def _reposition_toasts():
    with _toast_lock:
        for idx, t in enumerate(list(_active_toasts)):
            try:
                _position_toast_window(t, width=t.winfo_width() or 360, height=t.winfo_height() or 100, offset_index=idx)
            except Exception:
                pass

def _slide_and_fade_in(win, start_offset=40, steps=8, delay=15):
    try:
        geom = win.geometry()
        parts = geom.split('+')
        if len(parts) >= 3:
            base_x = int(parts[1])
            base_y = int(parts[2])
        else:
            base_x = win.winfo_x()
            base_y = win.winfo_y()
        
        for i in range(steps):
            frac = (i + 1) / steps
            y = int(base_y + start_offset * (1 - frac))
            alpha = frac
            try:
                win.geometry(f"{win.winfo_width()}x{win.winfo_height()}+{base_x}+{y}")
                win.attributes("-alpha", alpha)
            except Exception:
                pass
            win.update_idletasks()
            time.sleep(delay / 1000.0)
    except Exception:
        pass

def mostrar_toast(mensaje, tipo="info", duracion=3200, titulo=None):
    estilos = {
        "info":    {"icon": "ℹ️", "bg": "#0b1220", "dot": "#3A7BD5", "border": "#1E3A8A", "shadow": "#1E40AF"},
        "success": {"icon": "✅", "bg": "#07170b", "dot": "#16A34A", "border": "#166534", "shadow": "#15803D"},
        "warning": {"icon": "⚠️", "bg": "#1f1200", "dot": "#F59E0B", "border": "#92400E", "shadow": "#B45309"},
        "error":   {"icon": "❌", "bg": "#2b0e0e", "dot": "#EF4444", "border": "#991B1B", "shadow": "#DC2626"},
    }
    st = estilos.get(tipo, estilos["info"])

    TOAST_W = 360
    TOAST_H = 100

    with _toast_lock:
        offset_index = len(_active_toasts)

    toast = ctk.CTkToplevel(root)
    toast.overrideredirect(True)
    toast.attributes("-topmost", True)
    toast.configure(fg_color=st["bg"])  # Fondo principal sin borde
    
    # Frame principal con sombra sutil en la parte inferior
    main_frame = ctk.CTkFrame(toast, fg_color=st["bg"], corner_radius=16)
    main_frame.pack(fill="both", expand=True, padx=1, pady=1)
    
    # Frame de sombra inferior (simula sombra sutil)
    shadow_height = 4
    shadow_frame = ctk.CTkFrame(main_frame, fg_color=st["shadow"], height=shadow_height, 
                               corner_radius=8)
    shadow_frame.pack(side="bottom", fill="x", padx=8, pady=(0, 2))
    shadow_frame.pack_propagate(False)
    
    # Frame interno para el contenido
    wrapper = ctk.CTkFrame(main_frame, fg_color=st["bg"], corner_radius=14, 
                          width=TOAST_W-12, height=TOAST_H-12)
    wrapper.pack_propagate(False)
    wrapper.pack(fill="both", expand=False, padx=6, pady=(6, 4))  # Más padding abajo para la sombra

    # Frame del icono con círculo perfecto
    icon_frame = ctk.CTkFrame(wrapper, fg_color=st["bg"], width=56, height=56)
    icon_frame.pack_propagate(False)
    icon_frame.pack(side="left", padx=(12, 16), pady=12)

    # Canvas para el círculo perfectamente redondo
    dot_size = 48  # Tamaño del círculo
    dot = tk.Canvas(icon_frame, width=56, height=56, highlightthickness=0, bg=st["bg"])
    
    # Calcular posición para centrar perfectamente
    circle_x1 = (56 - dot_size) / 2
    circle_y1 = (56 - dot_size) / 2
    circle_x2 = circle_x1 + dot_size
    circle_y2 = circle_y1 + dot_size
    
    # Dibujar círculo perfectamente redondo
    dot.create_oval(circle_x1, circle_y1, circle_x2, circle_y2, 
                   fill=st["dot"], outline=st["dot"], width=0)
    
    # Calcular posición para centrar el emoji
    text_x = 56 / 2
    text_y = 56 / 2
    
    dot_text = dot.create_text(text_x, text_y, text=st["icon"], 
                              font=("Segoe UI Emoji", 16),
                              fill="#FFFFFF")
    dot.pack(fill="both", expand=True)

    # Contenedor de texto
    text_container = ctk.CTkFrame(wrapper, fg_color=st["bg"], corner_radius=0)
    text_container.pack(side="left", fill="both", expand=True, pady=12, padx=(0, 12))

    if titulo:
        lbl_title = ctk.CTkLabel(text_container, text=titulo, font=("Segoe UI", 12, "bold"), 
                                text_color="#FFFFFF", anchor="w")
        lbl_title.pack(fill="x", pady=(0, 2))
        lbl_msg = ctk.CTkLabel(text_container, text=mensaje, font=("Segoe UI", 11), 
                              text_color="#C7CCD1", wraplength=TOAST_W-140, 
                              anchor="w", justify="left")
        lbl_msg.pack(fill="both")
    else:
        lbl_msg = ctk.CTkLabel(text_container, text=mensaje, font=("Segoe UI", 12), 
                              text_color="#E6E9EE", wraplength=TOAST_W-140, 
                              anchor="w", justify="left")
        lbl_msg.pack(fill="both")

    # Función para cerrar la notificación al hacer click
    def close_toast(e=None):
        if toast.winfo_exists():
            with _toast_lock:
                if toast in _active_toasts:
                    _active_toasts.remove(toast)
            _animate_out_and_destroy(immediate=True)
    
    # Bind click para cerrar a todos los elementos
    wrapper.bind("<Button-1>", close_toast)
    text_container.bind("<Button-1>", close_toast)
    lbl_msg.bind("<Button-1>", close_toast)
    if titulo:
        lbl_title.bind("<Button-1>", close_toast)
    icon_frame.bind("<Button-1>", close_toast)
    dot.bind("<Button-1>", close_toast)
    main_frame.bind("<Button-1>", close_toast)
    shadow_frame.bind("<Button-1>", close_toast)

    toast.update_idletasks()

    _position_toast_window(toast, width=TOAST_W, height=TOAST_H, offset_index=offset_index)
    try:
        toast.attributes("-alpha", 0.0)
    except Exception:
        pass

    with _toast_lock:
        _active_toasts.append(toast)

    def _animate_in():
        try:
            steps = 6
            geom = toast.geometry()
            parts = geom.split('+')
            base_x = int(parts[1]) if len(parts) > 1 else toast.winfo_x()
            base_y = int(parts[2]) if len(parts) > 2 else toast.winfo_y()
            
            for i in range(steps):
                frac = (i + 1) / steps
                y_offset = int((1 - frac) * 20)
                y = base_y + y_offset
                alpha = frac
                
                try:
                    toast.geometry(f"{TOAST_W}x{TOAST_H}+{base_x}+{y}")
                    toast.attributes("-alpha", alpha)
                except Exception:
                    pass
                toast.update_idletasks()
                time.sleep(0.02)
        except Exception:
            pass

    def _animate_out_and_destroy(immediate=False):
        if not immediate:
            time.sleep(duracion / 1000.0)
        try:
            for i in range(5):
                alpha = max(0.0, 1 - (i + 1) / 5)
                try:
                    toast.attributes("-alpha", alpha)
                except Exception:
                    pass
                time.sleep(0.03)
        except Exception:
            pass
        try:
            toast.destroy()
        except Exception:
            pass
        with _toast_lock:
            if toast in _active_toasts:
                _active_toasts.remove(toast)
        try:
            _reposition_toasts()
        except Exception:
            pass

    threading.Thread(target=_animate_in, daemon=True).start()
    
    # Solo aplicar cierre automático si no se hizo click
    if not hasattr(toast, '_manually_closed'):
        threading.Thread(target=_animate_out_and_destroy, daemon=True).start()
# ANIMACIONES OPTIMIZADAS
def simple_button_hover(button, is_enter=True):
    if is_enter:
        if button == btn_excel:
            button.configure(fg_color=ACCENT_HOVER)
        elif button == btn_supabase:
            button.configure(fg_color=PRIMARY_HOVER)
        elif button == btn_iniciar and button.cget("state") == "normal":
            button.configure(fg_color=PRIMARY_HOVER)
        elif button == btn_volver:
            button.configure(fg_color="#7AB8FF")
        elif button == btn_close or button == btn_exit_menu:
            button.configure(fg_color="#FF6B6B")
        elif button == btn_descargar_informes:
            button.configure(fg_color=ACCENT_HOVER)
        elif button == btn_procesar_informes:
            button.configure(fg_color=PRIMARY_HOVER)
    else:
        if button == btn_excel:
            button.configure(fg_color=ACCENT)
        elif button == btn_supabase:
            button.configure(fg_color=PRIMARY)
        elif button == btn_iniciar and button.cget("state") == "normal":
            button.configure(fg_color=PRIMARY)
        elif button == btn_volver:
            button.configure(fg_color="#5EA0FF")
        elif button == btn_close or button == btn_exit_menu:
            button.configure(fg_color="#FF5252")
        elif button == btn_descargar_informes:
            button.configure(fg_color=ACCENT)
        elif button == btn_procesar_informes:
            button.configure(fg_color=PRIMARY)

def simple_item_hover(item, is_enter=True):
    if is_enter:
        item.configure(fg_color="#1E2025")
    else:
        item.configure(fg_color="#141518")

def quick_pulse_animation(widget, pulse_color, duration=0.3):
    def _pulse():
        try:
            original_color = widget.cget("fg_color")
            widget.configure(fg_color=pulse_color)
            time.sleep(duration)
            widget.configure(fg_color=original_color)
        except Exception:
            pass
    threading.Thread(target=_pulse, daemon=True).start()

def fade_transition(widget, target_alpha, duration=0.2):
    def _fade():
        try:
            steps = 4
            current_alpha = widget.attributes("-alpha") if widget.attributes("-alpha") else 1.0
            
            for i in range(steps):
                progress = (i + 1) / steps
                alpha = current_alpha + (target_alpha - current_alpha) * progress
                widget.attributes("-alpha", alpha)
                time.sleep(duration / steps)
        except Exception:
            widget.attributes("-alpha", target_alpha)
    
    threading.Thread(target=_fade, daemon=True).start()

# Confirmación cerrar app OPTIMIZADA
def confirmar_salida(titulo="Confirmar salida", mensaje="¿Deseas salir de la aplicación?"):
    result = {"value": False}
    width, height = 400, 140

    confirm = ctk.CTkToplevel(root)
    confirm.overrideredirect(True)
    confirm.attributes("-topmost", True)
    confirm.configure(fg_color="#0f1724")
    
    confirm.update_idletasks()
    screen_w = confirm.winfo_screenwidth()
    screen_h = confirm.winfo_screenheight()
    x = (screen_w // 2) - (width // 2)
    y = (screen_h // 2) - (height // 2)
    confirm.geometry(f"{width}x{height}+{x}+{y}")
    
    wrapper = ctk.CTkFrame(confirm, fg_color="#0f1724", corner_radius=14)
    wrapper.pack(fill="both", expand=True, padx=8, pady=8)

    barra = ctk.CTkFrame(wrapper, fg_color="#334155", width=8, corner_radius=14)
    barra.pack(side="left", fill="y", padx=(2, 8), pady=6)

    content = ctk.CTkFrame(wrapper, fg_color="#0f1724", corner_radius=8)
    content.pack(side="left", fill="both", expand=True, padx=(0, 8), pady=8)

    lbl_t = ctk.CTkLabel(content, text=titulo, font=("Segoe UI", 13, "bold"), text_color="#FFFFFF")
    lbl_t.pack(anchor="w", padx=8, pady=(6, 0))

    lbl_m = ctk.CTkLabel(
        content,
        text=mensaje,
        font=("Segoe UI", 11),
        text_color="#9CA3AF",
        wraplength=width - 80,
        anchor="w",
        justify="left"
    )
    lbl_m.pack(anchor="w", padx=8, pady=(4, 8))

    btn_frame = ctk.CTkFrame(content, fg_color="#0f1724")
    btn_frame.pack(anchor="e", padx=8, pady=(0, 8))

    def on_confirm():
        result["value"] = True
        fade_transition(confirm, 0.0, 0.15)
        confirm.after(150, confirm.destroy)

    def on_cancel():
        result["value"] = False
        fade_transition(confirm, 0.0, 0.15)
        confirm.after(150, confirm.destroy)

    btn_cancel = ctk.CTkButton(
        btn_frame,
        text="Cancelar",
        width=110,
        command=on_cancel,
        fg_color="#475569",
        hover_color="#64748B"
    )
    btn_cancel.pack(side="right", padx=(10, 0))

    btn_ok = ctk.CTkButton(
        btn_frame,
        text="Salir",
        width=110,
        command=on_confirm,
        fg_color="#DC2626",
        hover_color="#EF4444"
    )
    btn_ok.pack(side="right")
    
    try:
        confirm.attributes("-alpha", 0.0)
    except Exception:
        pass

    # Fade in rápido
    fade_transition(confirm, 1.0, 0.15)

    confirm.wait_window()
    return result["value"]

# ========= NUEVAS FUNCIONES PARA INVENTARIO =========

# Lista de códigos de clínica estáticos
CODIGOS_CLINICAS = ["0001", "0011", "0024", "0002", "0031", "0014", "0017", "0018", "0003"]

def descargar_informes_inventario():
    print("🏁 Iniciando descarga de informes de inventario...")
    mostrar_toast("Iniciando descarga de informes...", tipo="info", titulo="Descarga")
    
    def _proceso_descarga():
        try:
            print("🔗 Abriendo sistema de inventario en Edge...")
            subprocess.Popen(["C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe", URL_RETIRADA])
            time.sleep(10)
            
            # Proceso para descargar 9 informes (uno por cada clínica)
            for i, codigo_clinica in enumerate(CODIGOS_CLINICAS):
                print(f"📥 Descargando informe {i+1}/9 para clínica {codigo_clinica}...")
                
                # Buscar y hacer clic en el campo de unidad
                if not buscar_y_click("unidad_select.png", "unidad_select", confianza=0.6):
                    mostrar_toast("No se encontró el campo 'unidad'.", tipo="warning", titulo="Campo no encontrado")
                    continue
                
                try:
                    time.sleep(2)
                    pyautogui.typewrite(codigo_clinica)
                    time.sleep(3)
                    x, y = pyautogui.position()
                    pyautogui.moveTo(x, y + 50, duration=0.5)
                    pyautogui.click()
                    print(f"✅ Código de clínica {codigo_clinica} ingresado correctamente.")
                except Exception as e:
                    print(f"Error al ingresar código de clínica {codigo_clinica}: {e}")
                    continue
                time.sleep(3)
                # Llenar los campos necesarios para el informe
                if not buscar_y_click("tipo_costo.png", "tipo_costo"):
                    mostrar_toast("No se encontró 'tipo_costo.png'.", tipo="warning", titulo="Botón no encontrado")
                    continue
                
                time.sleep(5)
                pyautogui.typewrite("01")
                time.sleep(5)
                x, y = pyautogui.position()
                pyautogui.moveTo(x, y + 40, duration=0.5)
                pyautogui.click()
                time.sleep(5)
                
                if not buscar_y_click("nivel_totalizacion.png", "nivel_totalizacion", confianza=0.6):
                    mostrar_toast("No se encontró 'nivel_totalizacion.png'.", tipo="warning", titulo="Botón no encontrado")
                    continue
                
                time.sleep(5)
                x, y = pyautogui.position()
                pyautogui.moveTo(x, y + 40, duration=0.5)
                pyautogui.click()
                time.sleep(5)
                
                if not buscar_y_click("producto_inicial.png", "producto_inicial"):
                    mostrar_toast("No se encontró 'producto_inicial.png'.", tipo="warning", titulo="Botón no encontrado")
                    continue
                
                time.sleep(5)
                pyautogui.typewrite("1")
                time.sleep(5)
                x, y = pyautogui.position()
                pyautogui.moveTo(x, y + 40, duration=0.5)
                pyautogui.click()
                time.sleep(5)
                
                if not buscar_y_click("producto_final.png", "producto_final"):
                    mostrar_toast("No se encontró 'producto_final.png'.", tipo="warning", titulo="Botón no encontrado")
                    continue
                
                time.sleep(5)
                pyautogui.typewrite("5")
                time.sleep(5)
                x, y = pyautogui.position()
                pyautogui.moveTo(x, y + 40, duration=0.5)
                pyautogui.click()
                time.sleep(5)
                
                # Buscar y hacer clic en el botón de descarga XLSX
                if not buscar_y_click("generar_xlsx.png", "generar_xlsx", confianza=0.7):
                    mostrar_toast("No se encontró 'generar_xlsx.png'.", tipo="warning", titulo="Botón no encontrado")
                    continue
                
                print(f"✅ Descarga iniciada para clínica {codigo_clinica}")
                time.sleep(8)  # Esperar a que se complete la descarga
                
                # Recargar la página para la siguiente clínica (solo si no es la última)
                if i < len(CODIGOS_CLINICAS) - 1:
                    try:
                        pyautogui.hotkey("ctrl", "r")
                        time.sleep(10)  # Esperar a que la página se recargue completamente
                    except Exception:
                        try:
                            pyautogui.press("f5")
                            time.sleep(10)
                        except Exception:
                            pass
                
                mostrar_toast(f"Informe {i+1}/9 descargado (Clínica {codigo_clinica})", 
                             tipo="success", titulo="Progreso")
            
            print("✅ Descarga de informes completada")
            mostrar_toast("Descarga de 9 informes completada ✅", tipo="success", titulo="Éxito")
            
            # Habilitar botón de procesar
            root.after(0, lambda: btn_procesar_informes.configure(state="normal"))
            
        except Exception as e:
            print(f"❌ Error en descarga de informes: {e}")
            mostrar_toast(f"Error en descarga de informes:\n{e}", tipo="error", titulo="Error")
        finally:
            # Reactivar botón
            root.after(0, lambda: btn_descargar_informes.configure(state="normal"))
    
    btn_descargar_informes.configure(state="disabled")
    threading.Thread(target=_proceso_descarga, daemon=True).start()
def procesar_informes_inventario():
    print("🔄 Procesando informes descargados...")
    mostrar_toast("Procesando informes descargados...", tipo="info", titulo="Procesamiento")
    
    def _proceso_procesamiento():
        try:
            # Ruta donde se descargan los informes
            carpeta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")
            
            # Buscar archivos Excel recientes
            archivos_excel = []
            for archivo in os.listdir(carpeta_descargas):
                if archivo.endswith(('.xlsx', '.xls')) and any(keyword in archivo.lower() for keyword in ['inventario', 'stock', 'posicion', 'estoque']):
                    archivos_excel.append(os.path.join(carpeta_descargas, archivo))
            
            if len(archivos_excel) < 9:
                mostrar_toast(f"Solo se encontraron {len(archivos_excel)} archivos. Se esperaban 9.", 
                             tipo="warning", titulo="Advertencia")
                # Continuar con los que se encontraron
                if len(archivos_excel) == 0:
                    mostrar_toast("No se encontraron archivos para procesar", tipo="error", titulo="Error")
                    return
            
            # Tomar los 9 archivos más recientes (o los que haya)
            archivos_excel = sorted(archivos_excel, key=os.path.getmtime, reverse=True)[:9]
            
            datos_combinados = []
            
            for i, archivo in enumerate(archivos_excel):
                print(f"📊 Procesando archivo {i+1}/{len(archivos_excel)}: {os.path.basename(archivo)}")
                
                try:
                    # Leer el archivo Excel
                    df = pd.read_excel(archivo)
                    
                    # Definir columnas importantes para inventario
                    # Ajusta estas columnas según la estructura real de tus archivos
                    posibles_columnas_importantes = [
                        'Código', 'Código Producto', 'Producto', 'Nombre Producto', 
                        'Cantidad', 'Stock', 'Stock Actual', 'Existencia',
                        'Ubicación', 'Almacén', 'Unidad', 'Clínica'
                    ]
                    
                    # Filtrar solo las columnas importantes que existan en el archivo
                    columnas_existentes = [col for col in posibles_columnas_importantes if col in df.columns]
                    
                    if columnas_existentes:
                        # Agregar columna con el nombre del archivo de origen
                        df_filtrado = df[columnas_existentes].copy()
                        df_filtrado['Archivo_Origen'] = os.path.basename(archivo)
                        datos_combinados.append(df_filtrado)
                        print(f"✅ Procesadas {len(df_filtrado)} filas del archivo {os.path.basename(archivo)}")
                    else:
                        print(f"⚠️ No se encontraron columnas importantes en {archivo}")
                        # Si no encuentra columnas específicas, usar todas las columnas
                        df_filtrado = df.copy()
                        df_filtrado['Archivo_Origen'] = os.path.basename(archivo)
                        datos_combinados.append(df_filtrado)
                        print(f"✅ Usadas todas las {len(df_filtrado)} columnas del archivo {os.path.basename(archivo)}")
                        
                except Exception as e:
                    print(f"❌ Error procesando {archivo}: {e}")
                    mostrar_toast(f"Error procesando {os.path.basename(archivo)}", tipo="warning", titulo="Advertencia")
                    continue
            
            if datos_combinados:
                # Combinar todos los datos
                df_final = pd.concat(datos_combinados, ignore_index=True)
                
                # Eliminar duplicados exactos si es necesario
                df_final = df_final.drop_duplicates()
                
                # Crear nombre del archivo con fecha actual
                fecha_actual = datetime.now().strftime("%Y-%m-%d")
                nombre_archivo = f"Inventario de Stock {fecha_actual}.xlsx"
                
                # Guardar el archivo combinado
                ruta_guardado = os.path.join(carpeta_descargas, nombre_archivo)
                df_final.to_excel(ruta_guardado, index=False)
                
                print(f"✅ Archivo combinado guardado: {ruta_guardado}")
                print(f"📊 Total de registros consolidados: {len(df_final)}")
                
                mostrar_toast(f"Inventario consolidado guardado:\n{nombre_archivo}\nTotal: {len(df_final)} registros", 
                             tipo="success", titulo="Proceso Completado")
                
                # Actualizar la interfaz con la información del archivo creado
                root.after(0, lambda: actualizar_info_inventario(nombre_archivo, len(df_final)))
                
            else:
                mostrar_toast("No se pudieron procesar los archivos", tipo="error", titulo="Error")
                
        except Exception as e:
            print(f"❌ Error en procesamiento de informes: {e}")
            mostrar_toast(f"Error en procesamiento:\n{e}", tipo="error", titulo="Error")
        finally:
            # Reactivar botón
            root.after(0, lambda: btn_procesar_informes.configure(state="normal"))
    
    btn_procesar_informes.configure(state="disabled")
    threading.Thread(target=_proceso_procesamiento, daemon=True).start()

def actualizar_info_inventario(nombre_archivo="", total_items=0):
    try:
        lbl_info_inventario.configure(text=f"Archivo creado: {nombre_archivo}")
        lbl_estado_inventario.configure(text=f"Total de items procesados: {total_items}")
    except Exception as e:
        print(f"Error actualizando info inventario: {e}")

def actualizar_tabla_inventario():
    # Esta función puede usarse para mostrar los datos procesados en la tabla
    # Por ahora la dejamos como placeholder
    pass

# ========= FIN NUEVAS FUNCIONES ======================================

# FUNCIONES AUXILIARES ORIGINALES 

def buscar_y_click(imagen, nombre, confianza=0.9, intentos=3, esperar=5):
    for i in range(intentos):
        print(f"🔎 Buscando botón '{nombre}' (Intento {i+1}/{intentos})...")
        try:
            ubicacion = pyautogui.locateCenterOnScreen(
                os.path.join(CARPETA_BOTONES, imagen), confidence=confianza
            )
            if ubicacion:
                pyautogui.moveTo(ubicacion, duration=0.8)
                pyautogui.click()
                print(f"✅ Botón '{nombre}' encontrado y clickeado.")
                return True
        except pyautogui.ImageNotFoundException:
            pass
        print("⚠️ No encontrado, reintentando...")
        time.sleep(esperar)
    print(f"❌ No se encontró el botón '{nombre}'.")
    return False

def cargar_datos_desde_supabase():
    global df_facturas, productos_por_factura, codigo_clinica, origen_datos

    try:
        facturas_data = supabase.table("facturas").select("*").execute()
        productos_data = supabase.table("catalogo_productos").select("*").execute()

        print("🧩 Ejemplo de datos en 'productos':")
        if productos_data.data:
            print(productos_data.data[0])
        else:
            print("⚠️ No hay productos en Supabase.")

        if not facturas_data.data:
            mostrar_toast("No se encontraron registros en 'facturas'.", tipo="error", titulo="Sin registros")
            return
        if not productos_data.data:
            mostrar_toast("No se encontraron registros en 'productos'.", tipo="error", titulo="Sin registros")
            return

        facturas = []
        productos_por_factura = {}

        codigo_clinica = str(
            facturas_data.data[0].get("codigo_clinica", "0000")
        ).strip()

        for f in facturas_data.data:
            facturas.append({
            "ID_Factura": str(f["id"]),
            "N° Factura": str(f["numero_factura"]),
            "Fecha": str(f["fecha_factura"]),
            "Empresa": f.get("proveedores", {}).get("nombre", ""),
            "NIT": f.get("proveedores", {}).get("nit", "")
        })


            productos_por_factura[numero_factura] = []

        facturas_por_id = {
            str(f["id"]).strip(): str(f["numero_factura"]).strip()
            for f in facturas_data.data
        }

        for p in productos_data.data:
            factura_id = str(p.get("factura_id", "")).strip()
            if not factura_id:
                continue

            numero_factura = facturas_por_id.get(factura_id)
            if not numero_factura:
                continue 

            if numero_factura not in productos_por_factura:
                productos_por_factura[numero_factura] = []

            productos_por_factura[numero_factura].append(
                {
                    "Código Producto": str(p.get("codigo_producto", "")).strip(),
                    "Nombre Producto": str(p.get("nombre", "")).strip(),
                    "Cantidad": str(p.get("cantidad", "")).replace(",", ".").strip(),
                    "Precio": str(p.get("precio", "")).replace(",", ".").strip(),
                }
            )

        df_facturas = pd.DataFrame(facturas)
        origen_datos = "supabase"

        print(f"✅ Facturas cargadas: {len(facturas)}")
        print("✅ Productos agrupados por factura:")
        for k, v in productos_por_factura.items():
            print(f"  - {k}: {len(v)} productos")

        # --- Actualizar interfaz ---
        actualizar_tabla_facturas()
        try:
            btn_iniciar.configure(state="normal")
            # Animación de confirmación rápida
            quick_pulse_animation(btn_iniciar, "#4CAF50")
        except Exception:
            pass
        try:
            lbl_codigo.configure(text=f"🏥 Datos desde Supabase (Clínica {codigo_clinica})")
        except Exception:
            pass

        mostrar_toast("Datos cargados desde Supabase ✅", tipo="success", titulo="Éxito")

    except Exception as e:
        mostrar_toast(f"No se pudo conectar a Supabase:\n{e}", tipo="error", titulo="Error")

def cargar_datos_desde_excel():
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")],
    )
    if not archivo:
        mostrar_toast("No se seleccionó ningún archivo.", tipo="warning", titulo="Aviso")
        return

    global archivo_excel, df_facturas, productos_por_factura, codigo_clinica, origen_datos
    archivo_excel = archivo

    try:
        df = pd.read_excel(archivo, dtype=str).fillna("")
        columnas_esperadas = [
            "Tipo",
            "N° Factura",
            "Fecha",
            "Empresa",
            "NIT",
            "Código Producto",
            "Nombre Producto",
            "Cantidad",
            "Precio",
        ]
        for col in columnas_esperadas:
            if col not in df.columns:
                mostrar_toast(f"Falta la columna '{col}' en el archivo.", tipo="error", titulo="Error de formato")
                return

        facturas = []
        productos_por_factura = {}
        factura_actual = None

        for _, fila in df.iterrows():
            tipo = str(fila["Tipo"]).strip().upper()

            if tipo == "FACTURA":
                factura_actual = str(fila["N° Factura"]).strip()
                if not factura_actual:
                    continue
                facturas.append(
                    {
                        "ID_Factura": factura_actual,
                        "N° Factura": factura_actual,
                        "Fecha": fila["Fecha"].strip(),
                        "Empresa": fila["Empresa"].strip(),
                        "NIT": fila["NIT"].strip(),
                    }
                )
                productos_por_factura[factura_actual] = []
            elif tipo == "PRODUCTO" and factura_actual:
                productos_por_factura[factura_actual].append(
                    {
                        "Código Producto": str(fila["Código Producto"]).strip(),
                        "Nombre Producto": str(fila["Nombre Producto"]).strip(),
                        "Cantidad": str(fila["Cantidad"]).replace(",", ".").strip(),
                        "Precio": str(fila["Precio"]).replace(",", ".").strip(),
                    }
                )

        df_facturas = pd.DataFrame(facturas)
        codigo_clinica = "".join(
            [c for c in os.path.basename(archivo) if c.isdigit()][-4:]
        )
        if not codigo_clinica:
            codigo_clinica = "0000"

        origen_datos = "excel"
        actualizar_tabla_facturas()
        try:
            btn_iniciar.configure(state="normal")
            quick_pulse_animation(btn_iniciar, "#4CAF50")
        except Exception:
            pass
        try:
            lbl_codigo.configure(text=f"🏥 Código clínica detectado: {codigo_clinica}")
        except Exception:
            pass
        mostrar_toast("Archivo cargado correctamente ✅", tipo="success", titulo="Éxito")

    except Exception as e:
        mostrar_toast(f"No se pudo leer el archivo:\n{e}", tipo="error", titulo="Error")

def actualizar_tabla_facturas():
    for row in tree_facturas.get_children():
        tree_facturas.delete(row)
    for _, fila in df_facturas.iterrows():
        tree_facturas.insert("", "end", values=list(fila))

def mostrar_productos(event):
    item = tree_facturas.selection()
    if not item:
        return
    valores = tree_facturas.item(item[0], "values")
    if len(valores) < 2:
        return
    numero_factura = valores[1]

    productos = productos_por_factura.get(numero_factura, [])
    for row in tree_productos.get_children():
        tree_productos.delete(row)

    for prod in productos:
        tree_productos.insert(
            "",
            "end",
            values=(
                prod["Código Producto"],
                prod["Nombre Producto"],
                prod["Cantidad"],
                prod["Precio"],
            ),
        )

    try:
        lbl_productos.configure(
            text=f"🛒 Productos de la factura seleccionada (Total: {len(productos)})"
        )
    except Exception:
        pass
    print(f"🧾 Factura seleccionada: {numero_factura}")
    print(f"📦 Productos encontrados: {len(productos)}")

def iniciar_proceso():
    if df_facturas.empty:
        mostrar_toast("Primero carga datos desde Excel o Supabase.", tipo="error", titulo="Error")
        return

    print(f"🏁 Iniciando proceso para la clínica {codigo_clinica} ({origen_datos.upper()})...")
    print("🔹 Abriendo TecFood en Edge...")
    subprocess.Popen(["C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe", URL_TECFOOD])
    time.sleep(20)

    # Bucle principal: una iteración por cada fila (factura) en df_facturas
    for index, factura_actual in df_facturas.iterrows():
        numero_factura = str(factura_actual["N° Factura"]).strip()
        empresa = str(factura_actual.get("Empresa", "")).strip()
        nit = str(factura_actual.get("NIT", "")).strip()
        productos = productos_por_factura.get(numero_factura, [])

        print(f"\n\n🔁 Procesando factura {index+1}/{len(df_facturas)} → {numero_factura}")
        print(f"   Empresa: {empresa}  NIT: {nit}  Productos: {len(productos)}")

        if not productos:
            print(f"⚠️ Sin productos para {numero_factura}, se salta.")
            continue
        

        if not buscar_y_click("unidad_select.png", "unidad_select", confianza=0.6):
            mostrar_toast("No se encontró el botón unidad_select.png", tipo="warning", titulo="Botón no encontrado")
            continue

        # Selección de clínica
        try:
            time.sleep(3)
            pyautogui.typewrite(codigo_clinica, interval=0.1)
            time.sleep(2)
            x, y = pyautogui.position()
            pyautogui.moveTo(x, y + 50, duration=0.5)
            pyautogui.click()
            print(f"✅ Clínica {codigo_clinica} seleccionada correctamente.")
        except Exception as e:
            print("Error al seleccionar clínica:", e)
            continue
        pyautogui.press("tab")
        pyautogui.press("tab")
        pyautogui.press("tab")
        time.sleep(2)
        try:
            fecha = pd.to_datetime(factura_actual["Fecha"])
            fecha_formateada = fecha.strftime("%d%m%Y")
            pyautogui.typewrite(fecha_formateada, interval=0.1)
            pyautogui.typewrite(fecha_formateada, interval=0.1)
            pyautogui.press("tab")
            print(f"📅 Fecha ingresada: {fecha_formateada}")
        except Exception as e:
            print("⚠️ Error al procesar la fecha:", e)
            continue

        # Aplicar filtro
        time.sleep(3)
        if not (buscar_y_click("aplicar_filtro.png", "aplicar_filtro") or buscar_y_click("aplicar_filtro_en.png", "aplicar_filtro_en.png")):
            mostrar_toast("No se encontró el botón 'aplicar_filtro'.", tipo="warning", titulo="Botón no encontrado")
            continue

        # Añadir factura
        time.sleep(3)
        if not buscar_y_click("anadir.png", "anadir"):
            mostrar_toast("No se encontró el botón 'anadir.png'.", tipo="warning", titulo="Botón no encontrado")
            continue
        time.sleep(4)
        buscar_y_click("seleccionar_archivo.png", "seleccionar")
        time.sleep(1)

        # Buscar y seleccionar PDF correspondiente a esta factura
        try:
            empresa_limpia = "".join(c for c in empresa if c.isalnum() or c.isspace()).lower()
            nit_limpio = "".join(c for c in nit if c.isalnum()).lower()
            carpeta_pdf = r"C:\Users\Maicol Hernandez\Documents\GitHub\PyHealthy\PDF"
            pdf_encontrado: Optional[str] = None
            for archivo in os.listdir(carpeta_pdf):
                nombre_archivo = archivo.lower()
                if (nit_limpio in nombre_archivo and any(palabra in nombre_archivo for palabra in empresa_limpia.split()) and archivo.endswith(".pdf")):
                    pdf_encontrado = os.path.join(carpeta_pdf, archivo)
                    break

            if pdf_encontrado:
                print(f"📄 PDF encontrado: {pdf_encontrado}")
                pyautogui.typewrite(pdf_encontrado)
                time.sleep(1)
                pyautogui.press("enter")
                print("✅ PDF seleccionado correctamente.")
            else:
                mostrar_toast(f"No se encontró un PDF para '{empresa}' y '{nit}'. Verifica carpeta.", tipo="warning", titulo="PDF no encontrado")
                print("⚠️ PDF no encontrado. Se continúa con la factura (si corresponde).")
                # si quieres saltar la factura si no hay PDF, usa: continue
        except Exception as e:
            mostrar_toast(f"No se pudo procesar la factura o buscar el PDF:\n{e}", tipo="error", titulo="Error")
            continue

        # Remitente y número de factura
        time.sleep(5)
        if not buscar_y_click("remitente.png", "remitente"):
            mostrar_toast("No se encontró 'remitente.png'.", tipo="warning", titulo="Botón no encontrado")
            continue
        pyautogui.typewrite(nit)
        pyautogui.sleep(3)
        pyautogui.press("tab")
        pyautogui.sleep(7)
        pyautogui.typewrite(numero_factura)
        pyautogui.press("tab")
        pyautogui.sleep(2)

        if not buscar_y_click("serie.png", "serie"):
            mostrar_toast("No se encontró 'serie.png'.", tipo="warning", titulo="Botón no encontrado")
            continue
        x, y = pyautogui.position()
        pyautogui.moveTo(x, y + 40, duration=0.5)
        pyautogui.click()
        if not buscar_y_click("emision.png", "emision"):
            mostrar_toast("No se encontró 'emision.png'.", tipo="warning", titulo="Botón no encontrado")
            continue
        x, y = pyautogui.position()
        pyautogui.moveTo(x + 50, y + 30, duration=0.5)
        pyautogui.click()
        pyautogui.press("backspace", presses=10)
        pyautogui.typewrite(fecha_formateada)

        if not buscar_y_click("valor.png", "valor"):
            mostrar_toast("No se encontró 'valor.png'.", tipo="warning", titulo="Botón no encontrado")
            continue

        pyautogui.typewrite("0")
        pyautogui.press("tab")

        if not buscar_y_click("grabar.png", "grabar"):
            mostrar_toast("No se encontró 'grabar.png'.", tipo="warning", titulo="Botón no encontrado")
            continue

        pyautogui_sleep = getattr(pyautogui, "sleep", None)
        if pyautogui_sleep:
            pyautogui_sleep(12)
        else:
            time.sleep(12)

        if not buscar_y_click("productos.png", "productos"):
            mostrar_toast("No se encontró 'productos.png'.", tipo="warning", titulo="Botón no encontrado")
            continue

        time.sleep(6)

        # Agregar productos
        for i, producto in enumerate(productos):
            try:
                codigo_producto = str(producto["Código Producto"]).strip()
                cantidad = str(producto["Cantidad"]).strip().replace(",", ".")
                valor_unitario = str(producto["Precio"]).strip().replace(",", ".")

                print(f"➕ Agregando producto {i+1}/{len(productos)}: {codigo_producto}")
                if not buscar_y_click("anadir.png", "añadir"):
                    mostrar_toast("No se encontró 'anadir.png'.", tipo="warning", titulo="Botón no encontrado")
                    raise Exception("boton_anadir_no_encontrado")
                time.sleep(11)

                if not buscar_y_click("anadir_producto.png", "añadir_producto"):
                    mostrar_toast("No se encontró 'anadir_producto.png'.", tipo="warning", titulo="Botón no encontrado")
                    raise Exception("boton_anadir_producto_no_encontrado")
                time.sleep(11)

                pyautogui.typewrite(codigo_producto)
                time.sleep(3)
                x, y = pyautogui.position()
                pyautogui.moveTo(x, y + 40, duration=0.5)
                pyautogui.click()
                time.sleep(3)

                if not buscar_y_click("cantidad.png", "cantidad"):
                    mostrar_toast("No se encontró 'cantidad.png'.", tipo="warning", titulo="Botón no encontrado")
                    raise Exception("boton_cantidad_no_encontrado")
                pyautogui.typewrite(cantidad)
                pyautogui.press("tab")
                time.sleep(5)

                pyautogui.typewrite(valor_unitario)
                pyautogui.press("tab")
                time.sleep(5)

                if not buscar_y_click("grabar.png", "grabar"):
                    mostrar_toast("No se encontró 'grabar.png'.", tipo="warning", titulo="Botón no encontrado")
                    raise Exception("boton_grabar_no_encontrado")

                print(f"✅ Producto {codigo_producto} grabado correctamente.")
                time.sleep(5)
                pyautogui.press("esc")
                pyautogui_sleep = getattr(pyautogui, "sleep", None)
                if pyautogui_sleep:
                    pyautogui_sleep(5)
                else:
                    time.sleep(5)
            except Exception as e:
                print(f"❌ Error al agregar producto {codigo_producto}: {e}")
                continue

        # Finalizar esta factura
        if not buscar_y_click("FinalizarF.png", "FinalizarF"):
            mostrar_toast("No se encontró 'FinalizarF.png'.", tipo="warning", titulo="Botón no encontrado")
            continue
        time.sleep(5)
        if not buscar_y_click("si.png", "si"):
            mostrar_toast("No se encontró 'si.png'.", tipo="warning", titulo="Botón no encontrado")
            continue
        time.sleep(2)
        
        print(f"✅ Factura {numero_factura} procesada correctamente (productos: {len(productos)})")

        # --- Recargar la página para la siguiente factura ---
        time.sleep(2)
        try:
            pyautogui.hotkey("ctrl", "r")
            time.sleep(10)
        except Exception:
            try:
                pyautogui.press("f5")
                time.sleep(10)
            except Exception:
                pass
    print("🏁 Proceso completado para todas las facturas.")

# INTERFAZ CON CUSTOMTKINTER
try:
    root
except NameError:
    root = ctk.CTk()
    root.title("DataSpectra - Sistema de Inventario")
    root.attributes("-fullscreen", True)
    root.configure(fg_color="#0E0F12")

# Sistema simple de multipantallas
contenedor = ctk.CTkFrame(root, fg_color=BG)
contenedor.pack(fill="both", expand=True)

pantallas = {}
def mostrar_pantalla(nombre):
    for p in pantallas.values():
        p.pack_forget()
    pantallas[nombre].pack(fill="both", expand=True)

# PANTALLA PRINCIPAL
pantalla_menu = ctk.CTkFrame(contenedor, fg_color=BG)
pantallas["menu"] = pantalla_menu

# Header menu
header_menu = ctk.CTkFrame(pantalla_menu, fg_color=CARD_BG, corner_radius=18)
header_menu.pack(fill="x", padx=20, pady=16)

titulo_menu = ctk.CTkLabel(header_menu, text="DataSpectra — Panel Principal", font=("Segoe UI", 22, "bold"), text_color=TEXT_MAIN)
titulo_menu.pack(side="left", padx=18, pady=10)

# Exit button
def cerrar_app_wrapper():
    if confirmar_salida("Salir", "¿Deseas cerrar la aplicación?"):
        mostrar_toast("Cerrando aplicación...", tipo="info", titulo="Cerrando")
        root.after(300, root.destroy)
    else:
        mostrar_toast("Salida cancelada", tipo="info", titulo="Cancelado")

btn_exit_menu = ctk.CTkButton(header_menu, text="✕", width=44, height=38, corner_radius=14,
                              fg_color="#FF5252", hover_color="#FF6B6B", font=("Segoe UI", 14, "bold"),
                              command=cerrar_app_wrapper)
btn_exit_menu.pack(side="right", padx=14, pady=8)

# Añadir hover simple
btn_exit_menu.bind("<Enter>", lambda e: simple_button_hover(btn_exit_menu, True))
btn_exit_menu.bind("<Leave>", lambda e: simple_button_hover(btn_exit_menu, False))

content_menu = ctk.CTkScrollableFrame(pantalla_menu, fg_color=BG, corner_radius=0)
content_menu.pack(fill="both", expand=True, padx=20, pady=(10,20))

intro = ctk.CTkLabel(content_menu, text="Selecciona un proceso", font=("Segoe UI", 28, "bold"), text_color=TEXT_MAIN)
intro.pack(pady=(12,6))
sub = ctk.CTkLabel(content_menu, text="Funciones disponibles (las futuras aparecerán habilitadas cuando el administrador las implemente).", font=("Segoe UI", 15), text_color=TEXT_SECOND)
sub.pack(pady=(0,12))

# Lista vertical
def crear_item_lista(parent, icon, titulo, subtitulo, command=None, enabled=True):
    item = ctk.CTkFrame(parent, fg_color="#141518", corner_radius=16)
    item.pack(fill="x", padx=12, pady=8)
    item.configure(height=92)
    item.grid_propagate(False)

    # Icono a la izquierda
    lbl_icon = ctk.CTkLabel(item, text=icon, font=("Segoe UI Emoji", 32), text_color=ACCENT, fg_color="#141518")
    lbl_icon.place(x=18, y=28)

    # Textos
    lbl_tit = ctk.CTkLabel(item, text=titulo, font=("Segoe UI", 15, "bold"), text_color=TEXT_MAIN, fg_color="#141518")
    lbl_tit.place(x=80, y=18)
    lbl_sub = ctk.CTkLabel(item, text=subtitulo, font=("Segoe UI", 11), text_color=TEXT_SECOND, fg_color="#141518")
    lbl_sub.place(x=80, y=44)

    #indicador
    lbl_chev = ctk.CTkLabel(item, text="›", font=("Segoe UI", 22, "bold"), text_color="#8F9AA6", fg_color="#141518")
    lbl_chev.place(relx=0.96, rely=0.5, anchor="e")

    # Hover
    def on_enter(e):
        if enabled:
            simple_item_hover(item, True)
    def on_leave(e):
        if enabled:
            simple_item_hover(item, False)
    item.bind("<Enter>", on_enter)
    item.bind("<Leave>", on_leave)

    # Click
    if enabled and callable(command):
        def onclick(e=None):
            try:
                quick_pulse_animation(item, ACCENT)
                mostrar_toast(f"Abriendo: {titulo}", tipo="info", titulo="Abriendo Operación")
                command()
            except Exception as err:
                mostrar_toast(f"Ocurrió un error: {err}", tipo="error", titulo="Error")
        item.bind("<Button-1>", onclick)
        lbl_icon.bind("<Button-1>", onclick)
        lbl_tit.bind("<Button-1>", onclick)
        lbl_sub.bind("<Button-1>", onclick)
        lbl_chev.bind("<Button-1>", onclick)
    else:
        lbl_tit.configure(text_color="#6F777E")
        lbl_sub.configure(text_color="#4F5559")
        lbl_chev.configure(text_color="#3B4043")

    return item

# Crear 3 items
crear_item_lista(content_menu, "🧾", "Cargar facturas", "Automatiza la carga de facturas en TecFood", command=lambda: mostrar_pantalla("facturas"), enabled=True)
crear_item_lista(content_menu, "📦", "Sistema de Inventario Automatizado", "Descarga y consolida informes de inventario de stock", command=lambda: mostrar_pantalla("inventario"), enabled=True)
crear_item_lista(content_menu, "⚙️", "Función futura #2", "Implementación posterior", command=None, enabled=False)

# footer pequeño
footer_menu = ctk.CTkLabel(content_menu, text="Desarrollado por Area de desarrollo Healthy", font=("Segoe UI", 11), text_color=TEXT_SECOND)
footer_menu.pack(pady=18)

# =========================
# PANTALLA DE FACTURAS
# =========================
pantalla_facturas = ctk.CTkFrame(contenedor, fg_color=BG)
pantallas["facturas"] = pantalla_facturas

# === HEADER SUPERIOR ===
header = ctk.CTkFrame(pantalla_facturas, fg_color=CARD_BG, corner_radius=20)
header.pack(fill="x", padx=20, pady=12)

lbl_titulo = ctk.CTkLabel(
    header,
    text="DataSpectra - Carga de Facturas TecFood",
    font=("Segoe UI", 24, "bold"),
    text_color=TEXT_MAIN,
)
lbl_titulo.pack(side="left", padx=20, pady=12)

# Botón volver
def _on_volver():
    mostrar_toast("Volviendo al menú", tipo="info", titulo="Volver")
    mostrar_pantalla("menu")

btn_volver = ctk.CTkButton(
    header,
    text="↩ Volver",
    command=_on_volver,
    width=90,
    height=36,
    corner_radius=16,
    fg_color="#5EA0FF",
    hover_color="#7AB8FF",
    font=("Segoe UI", 13, "bold"),
)
btn_volver.pack(side="right", padx=10, pady=10)

# Añadir hover simple
btn_volver.bind("<Enter>", lambda e: simple_button_hover(btn_volver, True))
btn_volver.bind("<Leave>", lambda e: simple_button_hover(btn_volver, False))

# Botón cerrar
def cerrar_app_facturas():
    if confirmar_salida("Salir", "¿Deseas cerrar la aplicación?"):
        mostrar_toast("Cerrando aplicación...", tipo="info", titulo="Cerrando")
        root.after(300, root.destroy)
    else:
        mostrar_toast("Salida cancelada", tipo="info", titulo="Cancelado")

btn_close = ctk.CTkButton(
    header,
    text="✕",
    command=cerrar_app_facturas,
    width=40,
    height=40,
    corner_radius=20,
    fg_color="#FF5252",
    hover_color="#FF6B6B",
    font=("Segoe UI", 14, "bold"),
)
btn_close.pack(side="right", padx=10, pady=10)

# Añadir hover simple
btn_close.bind("<Enter>", lambda e: simple_button_hover(btn_close, True))
btn_close.bind("<Leave>", lambda e: simple_button_hover(btn_close, False))

# === CUERPO PRINCIPAL ===
main_frame = ctk.CTkScrollableFrame(pantalla_facturas, fg_color=BG, corner_radius=0)
main_frame.pack(fill="both", expand=True, padx=30, pady=10)

# --- BOTONES DE CARGA ---
btn_frame = ctk.CTkFrame(main_frame, fg_color=CARD_BG, corner_radius=25)
btn_frame.pack(pady=18)

def _on_cargar_excel():
    mostrar_toast("Selecciona un archivo para cargar", tipo="info", titulo="Cargar Excel")
    quick_pulse_animation(btn_excel, "#FFD54F")
    try:
        cargar_datos_desde_excel()
    except Exception as e:
        mostrar_toast(f"Error al ejecutar cargar_datos_desde_excel: {e}", tipo="error", titulo="Error")

def _on_cargar_supabase():
    mostrar_toast("Cargando desde Supabase...", tipo="info", titulo="Supabase")
    quick_pulse_animation(btn_supabase, "#4FC3F7")
    try:
        cargar_datos_desde_supabase()
    except Exception as e:
        mostrar_toast(f"Error al ejecutar cargar_datos_desde_supabase: {e}", tipo="error", titulo="Error")

btn_excel = ctk.CTkButton(
    btn_frame,
    text="📂 Cargar Excel",
    command=_on_cargar_excel,
    corner_radius=25,
    fg_color=ACCENT,
    hover_color=ACCENT_HOVER,
    text_color="#000000",
    font=("Segoe UI", 14, "bold"),
    width=200,
)
btn_excel.pack(side="left", padx=12, pady=12)

btn_supabase = ctk.CTkButton(
    btn_frame,
    text="🗄️ Cargar desde Supabase",
    command=_on_cargar_supabase,
    corner_radius=25,
    fg_color=PRIMARY,
    hover_color=PRIMARY_HOVER,
    text_color="#FFFFFF",
    font=("Segoe UI", 14, "bold"),
    width=250,
)
btn_supabase.pack(side="left", padx=12, pady=12)

btn_excel.bind("<Enter>", lambda e: simple_button_hover(btn_excel, True))
btn_excel.bind("<Leave>", lambda e: simple_button_hover(btn_excel, False))
btn_supabase.bind("<Enter>", lambda e: simple_button_hover(btn_supabase, True))
btn_supabase.bind("<Leave>", lambda e: simple_button_hover(btn_supabase, False))

# --- CÓDIGO DE CLÍNICA ---
lbl_codigo = ctk.CTkLabel(
    main_frame,
    text="Código de clínica detectado: ----",
    font=("Segoe UI", 14, "bold"),
    text_color=TEXT_SECOND,
)
lbl_codigo.pack(pady=6)

# === SECCIÓN DE TABLAS ===
tables_frame = ctk.CTkFrame(main_frame, fg_color=CARD_BG, corner_radius=20)
tables_frame.pack(fill="both", expand=True, padx=10, pady=10)

def crear_titulo(text):
    return ctk.CTkLabel(
        tables_frame,
        text=text,
        font=("Segoe UI", 16, "bold"),
        text_color=ACCENT,
        anchor="w",
    )

# --- FACTURAS ---
lbl_facturas = crear_titulo("📑 Facturas detectadas")
lbl_facturas.pack(anchor="w", pady=(10, 0), padx=12)

tree_facturas = ttk.Treeview(
    tables_frame,
    columns=("ID_Factura", "N° Factura", "Fecha", "Empresa", "NIT"),
    show="headings",
    height=8,
)
for col in ("ID_Factura", "N° Factura", "Fecha", "Empresa", "NIT"):
    tree_facturas.heading(col, text=col)
    tree_facturas.column(col, anchor="center", width=170)
tree_facturas.pack(fill="x", padx=12, pady=8)
if callable(globals().get("mostrar_productos")):
    tree_facturas.bind("<Double-1>", mostrar_productos)

# --- PRODUCTOS ---
lbl_productos = crear_titulo("🛒 Productos")
lbl_productos.pack(anchor="w", pady=(12, 0), padx=12)

tree_productos = ttk.Treeview(
    tables_frame,
    columns=("Código Producto", "Nombre Producto", "Cantidad", "Precio"),
    show="headings",
    height=8,
)
for col in ("Código Producto", "Nombre Producto", "Cantidad", "Precio"):
    tree_productos.heading(col, text=col)
    tree_productos.column(col, anchor="center", width=170)
tree_productos.pack(fill="x", padx=12, pady=8)

# --- ESTILO DE TABLAS ---
style = ttk.Style()
style.theme_use("clam")
style.configure(
    "Treeview",
    background="#22252A",
    fieldbackground="#22252A",
    foreground="#FFFFFF",
    rowheight=28,
    borderwidth=0,
    font=("Segoe UI", 11),
)
style.configure(
    "Treeview.Heading",
    background="#2E3238",
    foreground="#FFA726",
    font=("Segoe UI", 12, "bold"),
    borderwidth=0,
)
style.map("Treeview", background=[("selected", PRIMARY_HOVER)])

# --- BOTÓN INICIAR ---

def _run_iniciar_proceso_thread():
    try:
        root.after(0, lambda: btn_iniciar.configure(state="disabled"))
        root.after(0, lambda: mostrar_toast("Proceso iniciado en background", tipo="info", titulo="Proceso"))
        iniciar_proceso()
        root.after(0, lambda: btn_iniciar.configure(state="normal"))
        root.after(0, lambda: mostrar_toast("Proceso completado ✅", tipo="success", titulo="Completado"))
        # Animación de éxito rápida
        root.after(0, lambda: quick_pulse_animation(btn_iniciar, "#4CAF50"))
    except Exception as e:
        root.after(0, lambda: mostrar_toast(f"Error en proceso: {e}", tipo="error", titulo="Error"))
        root.after(0, lambda: btn_iniciar.configure(state="normal"))
        # Animación de error rápida
        root.after(0, lambda: quick_pulse_animation(btn_iniciar, "#F44336"))

def _on_iniciar_proceso():
    mostrar_toast("Iniciando proceso...", tipo="info", titulo="Iniciar")
    quick_pulse_animation(btn_iniciar, "#FFD54F")
    t = threading.Thread(target=_run_iniciar_proceso_thread, daemon=True)
    t.start()

btn_iniciar = ctk.CTkButton(
    main_frame,
    text="🚀 Iniciar proceso",
    command=_on_iniciar_proceso,
    fg_color=PRIMARY,
    hover_color=PRIMARY_HOVER,
    text_color="#FFFFFF",
    font=("Segoe UI", 15, "bold"),
    corner_radius=24,
    width=250,
    height=48,
    state="disabled",
)
btn_iniciar.pack(pady=16)

btn_iniciar.bind("<Enter>", lambda e: simple_button_hover(btn_iniciar, True))
btn_iniciar.bind("<Leave>", lambda e: simple_button_hover(btn_iniciar, False))

# --- FOOTER ---
footer = ctk.CTkFrame(pantalla_facturas, fg_color=CARD_BG, corner_radius=0)
footer.pack(fill="x", side="bottom")
lbl_footer = ctk.CTkLabel(footer, text="Desarrollado por Area de desarrollo de Healthy", font=("Segoe UI", 11), text_color=TEXT_SECOND)
lbl_footer.pack(pady=8)

# =========================
# PANTALLA DE INVENTARIO
# =========================
pantalla_inventario = ctk.CTkFrame(contenedor, fg_color=BG)
pantallas["inventario"] = pantalla_inventario

# === HEADER SUPERIOR ===
header_inventario = ctk.CTkFrame(pantalla_inventario, fg_color=CARD_BG, corner_radius=20)
header_inventario.pack(fill="x", padx=20, pady=12)

lbl_titulo_inventario = ctk.CTkLabel(
    header_inventario,
    text="Sistema de Inventario Automatizado",
    font=("Segoe UI", 24, "bold"),
    text_color=TEXT_MAIN,
)
lbl_titulo_inventario.pack(side="left", padx=20, pady=12)

# Botón volver
def _on_volver_inventario():
    mostrar_toast("Volviendo al menú", tipo="info", titulo="Volver")
    mostrar_pantalla("menu")

btn_volver_inventario = ctk.CTkButton(
    header_inventario,
    text="↩ Volver",
    command=_on_volver_inventario,
    width=90,
    height=36,
    corner_radius=16,
    fg_color="#5EA0FF",
    hover_color="#7AB8FF",
    font=("Segoe UI", 13, "bold"),
)
btn_volver_inventario.pack(side="right", padx=10, pady=10)

btn_volver_inventario.bind("<Enter>", lambda e: simple_button_hover(btn_volver_inventario, True))
btn_volver_inventario.bind("<Leave>", lambda e: simple_button_hover(btn_volver_inventario, False))

# Botón cerrar
def cerrar_app_inventario():
    if confirmar_salida("Salir", "¿Deseas cerrar la aplicación?"):
        mostrar_toast("Cerrando aplicación...", tipo="info", titulo="Cerrando")
        root.after(300, root.destroy)
    else:
        mostrar_toast("Salida cancelada", tipo="info", titulo="Cancelado")

btn_close_inventario = ctk.CTkButton(
    header_inventario,
    text="✕",
    command=cerrar_app_inventario,
    width=40,
    height=40,
    corner_radius=20,
    fg_color="#FF5252",
    hover_color="#FF6B6B",
    font=("Segoe UI", 14, "bold"),
)
btn_close_inventario.pack(side="right", padx=10, pady=10)

btn_close_inventario.bind("<Enter>", lambda e: simple_button_hover(btn_close_inventario, True))
btn_close_inventario.bind("<Leave>", lambda e: simple_button_hover(btn_close_inventario, False))

# === CUERPO PRINCIPAL ===
main_frame_inventario = ctk.CTkScrollableFrame(pantalla_inventario, fg_color=BG, corner_radius=0)
main_frame_inventario.pack(fill="both", expand=True, padx=30, pady=10)

# --- BOTONES DE PROCESO ---
btn_frame_inventario = ctk.CTkFrame(main_frame_inventario, fg_color=CARD_BG, corner_radius=25)
btn_frame_inventario.pack(pady=18)

def _on_descargar_informes():
    mostrar_toast("Iniciando descarga de informes...", tipo="info", titulo="Descarga")
    quick_pulse_animation(btn_descargar_informes, "#FFD54F")
    try:
        descargar_informes_inventario()
    except Exception as e:
        mostrar_toast(f"Error al descargar informes: {e}", tipo="error", titulo="Error")

btn_descargar_informes = ctk.CTkButton(
    btn_frame_inventario,
    text="📥 Descargar Informes",
    command=_on_descargar_informes,
    corner_radius=25,
    fg_color=ACCENT,
    hover_color=ACCENT_HOVER,
    text_color="#000000",
    font=("Segoe UI", 14, "bold"),
    width=220,
)
btn_descargar_informes.pack(side="left", padx=12, pady=12)

def _on_procesar_informes():
    mostrar_toast("Procesando informes descargados...", tipo="info", titulo="Procesamiento")
    quick_pulse_animation(btn_procesar_informes, "#4FC3F7")
    try:
        procesar_informes_inventario()
    except Exception as e:
        mostrar_toast(f"Error al procesar informes: {e}", tipo="error", titulo="Error")

btn_procesar_informes = ctk.CTkButton(
    btn_frame_inventario,
    text="🔄 Procesar Informes",
    command=_on_procesar_informes,
    corner_radius=25,
    fg_color=PRIMARY,
    hover_color=PRIMARY_HOVER,
    text_color="#FFFFFF",
    font=("Segoe UI", 14, "bold"),
    width=220,
    state="disabled",
)
btn_procesar_informes.pack(side="left", padx=12, pady=12)

# Añadir hover simple
btn_descargar_informes.bind("<Enter>", lambda e: simple_button_hover(btn_descargar_informes, True))
btn_descargar_informes.bind("<Leave>", lambda e: simple_button_hover(btn_descargar_informes, False))
btn_procesar_informes.bind("<Enter>", lambda e: simple_button_hover(btn_procesar_informes, True))
btn_procesar_informes.bind("<Leave>", lambda e: simple_button_hover(btn_procesar_informes, False))

# --- INFORMACIÓN DE ARCHIVO Y ESTADO ---
info_frame_inventario = ctk.CTkFrame(main_frame_inventario, fg_color=BG)
info_frame_inventario.pack(pady=12, fill="x")

lbl_info_inventario = ctk.CTkLabel(
    info_frame_inventario,
    text="Esperando descarga de informes...",
    font=("Segoe UI", 14, "bold"),
    text_color=TEXT_SECOND,
)
lbl_info_inventario.pack(anchor="w", pady=(0, 4))

lbl_estado_inventario = ctk.CTkLabel(
    info_frame_inventario,
    text="Estado: Listo para comenzar",
    font=("Segoe UI", 12),
    text_color=TEXT_SECOND,
)
lbl_estado_inventario.pack(anchor="w")

# === SECCIÓN INFORMATIVA ===
info_frame = ctk.CTkFrame(main_frame_inventario, fg_color=CARD_BG, corner_radius=20)
info_frame.pack(fill="x", padx=10, pady=10)

lbl_info_titulo = ctk.CTkLabel(
    info_frame,
    text="📋 Proceso de Inventario Automatizado",
    font=("Segoe UI", 16, "bold"),
    text_color=ACCENT,
    anchor="w",
)
lbl_info_titulo.pack(anchor="w", pady=(10, 5), padx=12)

lbl_info_desc = ctk.CTkLabel(
    info_frame,
    text="Este proceso automatizado:\n• Descarga 9 informes de inventario (uno por cada clínica)\n• Escribe automáticamente los códigos de cada clínica\n• Llena los campos necesarios y descarga en formato XLSX\n• Combina todos los datos en un solo archivo Excel\n• Guarda el resultado como 'Inventario de Stock {Fecha}'",
    font=("Segoe UI", 12),
    text_color=TEXT_SECOND,
    anchor="w",
    justify="left",
)
lbl_info_desc.pack(anchor="w", pady=(0, 10), padx=12)

# --- FOOTER ---
footer_inventario = ctk.CTkFrame(pantalla_inventario, fg_color=CARD_BG, corner_radius=0)
footer_inventario.pack(fill="x", side="bottom")
lbl_footer_inventario = ctk.CTkLabel(footer_inventario, text="Desarrollado por Area de desarrollo de Healthy", font=("Segoe UI", 11), text_color=TEXT_SECOND)
lbl_footer_inventario.pack(pady=8)

# MOSTRAR PANTALLA INICIAL
mostrar_pantalla("menu")

# Solo iniciamos mainloop si no hay otro en el archivo
if not hasattr(root, "_DataSpectra_mainloop_started"):
    root._DataSpectra_mainloop_started = True
    root.mainloop()