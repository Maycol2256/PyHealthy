import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
import subprocess
import time
import pyautogui
from supabase import create_client, Client
from typing import Optional

# === CONFIGURACIÓN ===
CARPETA_BOTONES = "Buttons"
URL_TECFOOD = "https://food.teknisa.com//df/#/df_entrada#dfe11000_lancamento_entrada"

SUPABASE_URL = "https://ulboklgzjriatmaxzpsi.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVsYm9rbGd6anJpYXRtYXh6cHNpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjI3ODQxNDAsImV4cCI6MjA3ODM2MDE0MH0.gY6_K4JQoJxPZmdXMIbFZfiJAOdavbg8jDJW1rOUSPk"
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# === VARIABLES GLOBALES ===
archivo_excel = None
df_facturas = pd.DataFrame()
productos_por_factura = {}
codigo_clinica = ""
origen_datos = ""


# === FUNCIONES ===
def buscar_y_click(imagen, nombre, confianza=0.9, intentos=3, esperar=5):
    """Busca una imagen en pantalla con varios intentos antes de fallar."""
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
    """Carga las facturas y productos directamente desde Supabase."""
    global df_facturas, productos_por_factura, codigo_clinica, origen_datos

    try:
        facturas_data = supabase.table("facturas").select("*").execute()
        productos_data = supabase.table("productos").select("*").execute()

        print("🧩 Ejemplo de datos en 'productos':")
        if productos_data.data:
            print(productos_data.data[0])
        else:
            print("⚠️ No hay productos en Supabase.")

        if not facturas_data.data:
            messagebox.showerror("Error", "No se encontraron registros en 'facturas'.")
            return
        if not productos_data.data:
            messagebox.showerror("Error", "No se encontraron registros en 'productos'.")
            return

        facturas = []
        productos_por_factura = {}

        # Obtener el código de clínica desde la primera factura
        codigo_clinica = str(
            facturas_data.data[0].get("codigo_clinica", "0000")
        ).strip()

        # --- Construir tabla de facturas ---
        for f in facturas_data.data:
            id_factura = str(f.get("id", "")).strip()
            numero_factura = str(f.get("numero_factura", "")).strip()

            facturas.append(
                {
                    "ID_Factura": id_factura,
                    "N° Factura": numero_factura,
                    "Fecha": str(f.get("fecha", "")).strip(),
                    "Empresa": str(f.get("nombre_empresa", "")).strip(),
                    "NIT": str(f.get("nit", "")).strip(),
                }
            )

            # Inicializar lista vacía de productos por número de factura
            productos_por_factura[numero_factura] = []

        # --- Asociar productos con sus facturas usando factura_id ---
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
                continue  # Si no hay coincidencia, se omite

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

        # --- Convertir lista de facturas en DataFrame ---
        df_facturas = pd.DataFrame(facturas)
        origen_datos = "supabase"

        # --- Mostrar conteo y depuración ---
        print(f"✅ Facturas cargadas: {len(facturas)}")
        print("✅ Productos agrupados por factura:")
        for k, v in productos_por_factura.items():
            print(f"  - {k}: {len(v)} productos")

        # --- Actualizar interfaz ---
        actualizar_tabla_facturas()
        btn_iniciar.config(state="normal")
        lbl_codigo.config(text=f"🏥 Datos desde Supabase (Clínica {codigo_clinica})")

        messagebox.showinfo("Éxito", "Datos cargados desde Supabase ✅")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo conectar a Supabase:\n{e}")


def cargar_datos_desde_excel():
    """Carga facturas y productos desde Excel."""
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")],
    )
    if not archivo:
        messagebox.showwarning("Aviso", "No se seleccionó ningún archivo.")
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
                messagebox.showerror(
                    "Error", f"Falta la columna '{col}' en el archivo."
                )
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
        btn_iniciar.config(state="normal")
        lbl_codigo.config(text=f"🏥 Código clínica detectado: {codigo_clinica}")
        messagebox.showinfo("Éxito", "Archivo cargado correctamente ✅")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")


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

    lbl_productos.config(
        text=f"🛒 Productos de la factura seleccionada (Total: {len(productos)})"
    )
    print(f"🧾 Factura seleccionada: {numero_factura}")
    print(f"📦 Productos encontrados: {len(productos)}")


def iniciar_proceso():
    """Usa los datos cargados (de Excel o Supabase)."""
    if df_facturas.empty:
        messagebox.showerror("Error", "Primero carga datos desde Excel o Supabase.")
        return

    print(
        f"🏁 Iniciando proceso para la clínica {codigo_clinica} ({origen_datos.upper()})..."
    )
    print("🔹 Abriendo TecFood en Edge...")
    subprocess.Popen(
        ["C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe", URL_TECFOOD]
    )

    time.sleep(20)

    if not buscar_y_click("unidad_select.png", "unidad_select", confianza=0.6):
        messagebox.showwarning(
            "Advertencia", "No se encontró el botón unidad_select.png"
        )
        return

    # Selección de clínica
    time.sleep(5)
    pyautogui.typewrite(codigo_clinica, interval=0.1)
    time.sleep(2)
    x, y = pyautogui.position()
    pyautogui.moveTo(x, y + 50, duration=0.5)
    pyautogui.click()
    print(f"✅ Clínica {codigo_clinica} seleccionada correctamente.")

    # Aplicar filtro
    time.sleep(5)
    if not (
        buscar_y_click("aplicar_filtro.png", "aplicar_filtro")
        or buscar_y_click("aplicar_filtro_en.png", "aplicar_filtro_en.png")
    ):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'aplicar_filtro.png' ni 'aplicar_filtro_en.png'.",
        )
        return

    # Añadir factura
    time.sleep(5)
    if not buscar_y_click("anadir.png", "anadir"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'anadir.png'.",
        )
        return

    # Seleccionar archivo PDF
    time.sleep(8)
    buscar_y_click("seleccionar_archivo.png", "seleccionar")

    time.sleep(5)

    if not buscar_y_click("Escritorio.png", "Escritorio"):
        messagebox.showwarning("Advertencia", "No se encontró 'Escritorio.png'.")
        return
    time.sleep(2)

    if not buscar_y_click("Carpeta.png", "Carpeta"):
        messagebox.showwarning("Advertencia", "No se encontró 'Carpeta.png'.")
        return

    pyautogui.doubleClick()
    x, y = pyautogui.position()
    pyautogui.moveTo(x, y + 50, duration=0.5)

    time.sleep(2)

    if not buscar_y_click("carpetapdf.png", "carpetapdf"):
        messagebox.showwarning("Advertencia", "No se encontró 'carpetapdf.png'.")
        return

    pyautogui.doubleClick()
    time.sleep(2)

    # Buscar y seleccionar PDF correspondiente
    try:
        factura_actual = df_facturas.iloc[0]
        empresa = str(factura_actual["Empresa"]).strip()
        nit = str(factura_actual["NIT"]).strip()
        numero_factura = str(factura_actual["N° Factura"]).strip()

        productos = productos_por_factura.get(numero_factura, [])
        if productos:
            primer_producto = productos[0]
            codigo_producto = str(primer_producto["Código Producto"]).strip()
            cantidad = str(primer_producto["Cantidad"]).strip().replace(",", ".")
            valor_unitario = str(primer_producto["Precio"]).strip().replace(",", ".")
        else:
            codigo_producto = cantidad = valor_unitario = ""
            print(f"⚠️ No se encontraron productos para la factura {numero_factura}")

        empresa_limpia = "".join(
            c for c in empresa if c.isalnum() or c.isspace()
        ).lower()
        nit_limpio = "".join(c for c in nit if c.isalnum()).lower()

        carpeta_pdf = (
            r"O:\Perfil\Rogers Allan Merchan Sepulveda\Escritorio\BotHealthyBeta01\PDF"
        )

        pdf_encontrado: Optional[str] = None
        for archivo in os.listdir(carpeta_pdf):
            nombre_archivo = archivo.lower()
            if (
                nit_limpio in nombre_archivo
                and any(palabra in nombre_archivo for palabra in empresa_limpia.split())
                and archivo.endswith(".pdf")
            ):
                pdf_encontrado = os.path.join(carpeta_pdf, archivo)
                break

        if pdf_encontrado:
            print(f"📄 PDF encontrado: {pdf_encontrado}")
            pyautogui.typewrite(pdf_encontrado)
            time.sleep(1)
            pyautogui.press("enter")
            print("✅ PDF seleccionado correctamente.")
        else:
            messagebox.showwarning(
                "Advertencia",
                f"No se encontró un PDF que contenga '{empresa}' y '{nit}' en su nombre.\n"
                "Verifica que el archivo esté en la carpeta configurada.",
            )
            print("⚠️ PDF no encontrado.")

    except Exception as e:
        messagebox.showerror(
            "Error", f"No se pudo procesar la factura o buscar el PDF:\n{e}"
        )
        return

    time.sleep(5)
    if not buscar_y_click("remitente.png", "remitente"):
        messagebox.showwarning("Advertencia", "No se encontró 'remitente.png'.")
        return

    pyautogui.typewrite(nit)
    pyautogui.sleep(3)
    pyautogui.press("tab")
    pyautogui.sleep(14)

    if not buscar_y_click("numero_factura.png", "numero_factura"):
        messagebox.showwarning("Advertencia", "No se encontró 'numero_factura.png'.")
        return

    pyautogui.typewrite(numero_factura)
    pyautogui.press("tab")
    pyautogui.sleep(2)

    if not buscar_y_click("serie.png", "serie"):
        messagebox.showwarning("Advertencia", "No se encontró 'serie.png'.")
        return

    x, y = pyautogui.position()
    pyautogui.moveTo(x, y + 40, duration=0.5)
    pyautogui.click()

    if not buscar_y_click("valor.png", "valor"):
        messagebox.showwarning("Advertencia", "No se encontró 'valor.png'.")
        return

    pyautogui.typewrite("0")
    pyautogui.press("tab")

    if not buscar_y_click("grabar.png", "grabar"):
        messagebox.showwarning("Advertencia", "No se encontró 'grabar.png'.")
        return

    pyautogui.sleep(13)

    if not buscar_y_click("productos.png", "productos"):
        messagebox.showwarning("Advertencia", "No se encontró 'productos.png'.")
        return

    pyautogui.sleep(6)

    # Agregar productos
    for i, producto in enumerate(productos):
        codigo_producto = str(producto["Código Producto"]).strip()
        cantidad = str(producto["Cantidad"]).strip().replace(",", ".")
        valor_unitario = str(producto["Precio"]).strip().replace(",", ".")

        print(f"➕ Agregando producto {i+1}/{len(productos)}: {codigo_producto}")

        if not buscar_y_click("anadir.png", "añadir"):
            messagebox.showwarning("Advertencia", "No se encontró 'anadir.png'.")
            return
        time.sleep(10)

        if not buscar_y_click("anadir_producto.png", "añadir_producto"):
            messagebox.showwarning(
                "Advertencia", "No se encontró 'anadir_producto.png'."
            )
            return
        time.sleep(10)

        pyautogui.typewrite(codigo_producto)
        time.sleep(4)
        x, y = pyautogui.position()
        pyautogui.moveTo(x, y + 40, duration=0.5)
        pyautogui.click()
        time.sleep(2)

        if not buscar_y_click("cantidad.png", "cantidad"):
            messagebox.showwarning("Advertencia", "No se encontró 'cantidad.png'.")
            return

        pyautogui.typewrite(cantidad)
        pyautogui.press("tab")
        time.sleep(5)

        pyautogui.typewrite(valor_unitario)
        pyautogui.press("tab")
        time.sleep(5)

        if not buscar_y_click("grabar.png", "grabar"):
            messagebox.showwarning("Advertencia", "No se encontró 'grabar.png'.")
            return

        print(f"✅ Producto {codigo_producto} grabado correctamente.")
        time.sleep(5)
        pyautogui.press("esc")
        pyautogui.sleep(5)

    if not buscar_y_click("FinalizarF.png", "FinalizarF"):
        messagebox.showwarning("Advertencia", "No se encontró 'FinalizarF.png'.")
        return

    print(f"Todos los {len(productos)} productos se agregaron correctamente ✅")


# === INTERFAZ TKINTER ===
root = tk.Tk()
root.title("BotHealthy Beta - Carga de Facturas TecFood")
root.geometry("900x600")

lbl_titulo = tk.Label(root, text="BotHealthy Beta", font=("Segoe UI", 24, "bold"))
lbl_titulo.pack(pady=12)

frame_botones = tk.Frame(root)
frame_botones.pack(pady=5)

btn_excel = tk.Button(
    frame_botones,
    text="📂 Cargar archivo Excel",
    command=cargar_datos_desde_excel,
    font=("Segoe UI", 11),
    background="green",
    foreground="white",
)
btn_excel.pack(side="left", padx=5)

btn_supabase = tk.Button(
    frame_botones,
    text="🗄️ Cargar desde Supabase",
    command=cargar_datos_desde_supabase,
    font=("Segoe UI", 11),
    background="blue",
    foreground="white",
)
btn_supabase.pack(side="left", padx=5)

lbl_codigo = tk.Label(
    root, text="💻 Código de clínica detectado: ----", font=("Segoe UI", 11)
)
lbl_codigo.pack(pady=5)


frame_tablas = tk.Frame(root)
frame_tablas.pack(pady=10, fill="both", expand=True)

# Facturas
lbl_facturas = tk.Label(
    frame_tablas, text="📑 Facturas detectadas", font=("Segoe UI", 12, "bold")
)
lbl_facturas.pack()

columnas_fact = ("ID_Factura", "N° Factura", "Fecha", "Empresa", "NIT")
tree_facturas = ttk.Treeview(
    frame_tablas, columns=columnas_fact, show="headings", height=6
)
for col in columnas_fact:
    tree_facturas.heading(col, text=col)
    tree_facturas.column(col, width=150)
tree_facturas.pack(side="top", fill="both", expand=True)
tree_facturas.bind("<Double-1>", mostrar_productos)

# Productos
lbl_productos = tk.Label(
    frame_tablas, text="🛒 Productos", font=("Segoe UI", 12, "bold")
)
lbl_productos.pack()

columnas_prod = ("Código Producto", "Nombre Producto", "Cantidad", "Precio")
tree_productos = ttk.Treeview(
    frame_tablas, columns=columnas_prod, show="headings", height=6
)
for col in columnas_prod:
    tree_productos.heading(col, text=col)
    tree_productos.column(col, width=150)
tree_productos.pack(side="bottom", fill="both", expand=True)
btn_iniciar = tk.Button(
    root,
    text="🚀 Comenzar",
    background="orange",
    state="disabled",
    command=iniciar_proceso,
    font=("Segoe UI", 12, "bold"),
)
btn_iniciar.pack(pady=15)

root.mainloop()
