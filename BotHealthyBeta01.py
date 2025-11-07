import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
import subprocess
import time
import pyautogui

# === CONFIGURACIÓN ===
CARPETA_BOTONES = "Buttons"
URL_TECFOOD = "https://food.teknisa.com//df/#/df_entrada#dfe11000_lancamento_entrada"


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
        print(f"⚠️ No encontrado, reintentando...")
        time.sleep(esperar)
    print(f"❌ No se encontró el botón '{nombre}'.")
    return False


def abrir_excel():
    """Carga el archivo Excel, detecta facturas y asocia productos."""
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")],
    )
    if not archivo:
        messagebox.showwarning("Aviso", "No se seleccionó ningún archivo.")
        return

    global archivo_excel, df_facturas, productos_por_factura, codigo_clinica
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

            # Detectar una nueva factura
            if tipo == "FACTURA":
                factura_actual = str(fila["N° Factura"]).strip()
                if not factura_actual:
                    continue
                facturas.append(
                    {
                        "N° Factura": factura_actual,
                        "Fecha": fila["Fecha"].strip(),
                        "Empresa": fila["Empresa"].strip(),
                        "NIT": fila["NIT"].strip(),
                    }
                )
                productos_por_factura[factura_actual] = []

            # Detectar productos asociados
            elif tipo == "PRODUCTO" or (tipo == "" and factura_actual):
                if not factura_actual:
                    continue
                codigo = str(fila["Código Producto"]).strip()
                nombre = str(fila["Nombre Producto"]).strip()
                cantidad = str(fila["Cantidad"]).replace(",", ".").strip()
                precio = str(fila["Precio"]).strip()
                if codigo or nombre:
                    productos_por_factura[factura_actual].append(
                        {
                            "Código Producto": codigo,
                            "Nombre Producto": nombre,
                            "Cantidad": cantidad,
                            "Precio": precio,
                        }
                    )

        df_facturas = pd.DataFrame(facturas)

        if df_facturas.empty:
            messagebox.showerror("Error", "No se encontraron facturas en el archivo.")
            return

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")
        return

    # Código de clínica desde el nombre del archivo
    codigo_clinica = "".join([c for c in os.path.basename(archivo) if c.isdigit()][-4:])
    if not codigo_clinica:
        codigo_clinica = "0000"
    lbl_codigo.config(text=f"🏥 Código de clínica detectado: {codigo_clinica}")

    actualizar_tabla_facturas()
    btn_iniciar.config(state="normal")
    messagebox.showinfo("Éxito", "Archivo cargado correctamente ✅")


def actualizar_tabla_facturas():
    """Muestra las facturas en la tabla principal."""
    for row in tree_facturas.get_children():
        tree_facturas.delete(row)
    for _, fila in df_facturas.iterrows():
        tree_facturas.insert("", "end", values=list(fila))


def mostrar_productos(event):
    """Muestra los productos asociados a la factura seleccionada."""
    item = tree_facturas.selection()
    if not item:
        return
    factura_id = tree_facturas.item(item[0], "values")[0]
    productos = productos_por_factura.get(factura_id, [])

    # Limpiar tabla de productos
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


def iniciar_proceso():
    """Ejecuta la secuencia automática en TecFood."""
    if not archivo_excel:
        messagebox.showerror("Error", "Primero debes cargar un archivo Excel.")
        return

    print(f"🏁 Iniciando proceso para la clínica {codigo_clinica}...")

    # Abrir navegador Edge
    print("🔹 Abriendo TecFood en Edge...")
    subprocess.Popen(
        ["C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe", URL_TECFOOD]
    )

    # Esperar carga completa de la página
    time.sleep(18)

    # 🔹 Ajuste específico: botón de unidad con menor confianza
    if not buscar_y_click("unidad_select.png", "Selector de unidad", confianza=0.7):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'unidad_select.png'. Verifica que la página esté visible.",
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
    buscar_y_click("aplicar_filtro.png", "Aplicar filtro")

    # Añadir factura
    time.sleep(5)
    if not buscar_y_click("anadir.png", "anadir"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'anadir.png'. Verifica que la página esté visible.",
        )
        return

    # Añadir PDF
    time.sleep(8)
    buscar_y_click("seleccionar_archivo.png", "seleccionar")

    time.sleep(5)
    if not buscar_y_click("Escritorio.png", "Escritorio"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'Escritorio.png'. Verifica que la página esté visible.",
        )
        return
    time.sleep(2)
    if not buscar_y_click("Carpeta.png", "Carpeta"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'Carpeta.png'. Verifica que la página esté visible.",
        )
        return
    pyautogui.doubleClick()
    x, y = pyautogui.position()
    pyautogui.moveTo(x, y + 50, duration=0.5)

    time.sleep(2)
    if not buscar_y_click("carpetapdf.png", "carpetapdf"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'carpetapdf.png'. Verifica que la página esté visible.",
        )
        return

    pyautogui.doubleClick()
    time.sleep(2)

    # === Buscar y seleccionar el PDF correspondiente ===
    try:
        factura_actual = df_facturas.iloc[0]
        empresa = str(factura_actual["Empresa"]).strip()
        nit = str(factura_actual["NIT"]).strip()
        numero_factura = str(factura_actual["N° Factura"]).strip()

        # 🧩 Extraer datos de productos asociados a la factura
        productos = productos_por_factura.get(numero_factura, [])
        if productos:
            primer_producto = productos[0]
            codigo_producto = str(primer_producto["Código Producto"]).strip()
            cantidad = str(primer_producto["Cantidad"]).strip().replace(",", ".")
            valor_unitario = str(primer_producto["Precio"]).strip().replace(",", ".")
        else:
            codigo_producto = cantidad = valor_unitario = ""
            print(f"⚠️ No se encontraron productos para la factura {numero_factura}")

        # Limpiar caracteres problemáticos para búsqueda
        empresa_limpia = "".join(
            c for c in empresa if c.isalnum() or c.isspace()
        ).lower()
        nit_limpio = "".join(c for c in nit if c.isalnum()).lower()

        # Carpeta donde buscar los PDF
        carpeta_pdf = r"O:\Perfil\Rogers Allan Merchan Sepulveda\Escritorio\BotHealthyBeta01\PDF"

        # Buscar coincidencia en nombres de archivo
        pdf_encontrado = None
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
                "Verifica que el archivo esté en la carpeta de PDF configurada.",
            )
            print("⚠️ PDF no encontrado.")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo procesar la factura o buscar el PDF:\n{e}")
        return

    # === Continuación del flujo ===
    time.sleep(5)
    if not buscar_y_click("remitente.png", "remitente"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'remitente.png'. Verifica que la página esté visible.",
        )
        return
    pyautogui.typewrite(nit)
    pyautogui.sleep(3)
    pyautogui.press("tab")
    pyautogui.sleep(14)

    if not buscar_y_click("numero_factura.png", "numero_factura"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'numero_factura.png'. Verifica que la página esté visible.",
        )
        return
    pyautogui.typewrite(numero_factura)
    pyautogui.press("tab")
    pyautogui.sleep(2)

    if not buscar_y_click("serie.png", "serie"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'serie.png'. Verifica que la página esté visible.",
        )
        return

    x, y = pyautogui.position()
    pyautogui.moveTo(x, y + 40, duration=0.5)
    pyautogui.click()

    if not buscar_y_click("valor.png", "valor"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'valor.png'. Verifica que la página esté visible.",
        )
        return
    pyautogui.typewrite("0")
    pyautogui.press("tab")

    if not buscar_y_click("grabar.png", "grabar"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'grabar.png'. Verifica que la página esté visible.",
        )
        return

    pyautogui.sleep(13)

    if not buscar_y_click("productos.png", "productos"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'productos.png'. Verifica que la página esté visible.",
        )
        return

    pyautogui.sleep(6)

    if not buscar_y_click("anadir.png", "anadir"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'anadir.png'. Verifica que la página esté visible.",
        )
        return

    pyautogui.sleep(6)

    if not buscar_y_click("anadir_producto.png", "anadir_producto"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'anadir_producto.png'. Verifica que la página esté visible.",
        )
        return

    pyautogui.typewrite(codigo_producto)
    pyautogui.press("enter")

    pyautogui.sleep(5)

    if not buscar_y_click("cantidad.png", "cantidad"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'cantidad.png'. Verifica que la página esté visible.",
        )
        return
    pyautogui.typewrite(cantidad)
    pyautogui.press("tab")
    pyautogui.sleep(2)
    pyautogui.typewrite(valor_unitario)
    pyautogui.press("tab")
    pyautogui.sleep(2)

    if not buscar_y_click("grabar.png", "grabar"):
        messagebox.showwarning(
            "Advertencia",
            "No se encontró el botón 'grabar.png'. Verifica que la página esté visible.",
        )
        return

    messagebox.showinfo("Finalizado", "Primer paso completado correctamente ✅")


# === INTERFAZ TKINTER ===
root = tk.Tk()
root.title("BotHealthy Beta - Carga de Facturas TecFood")

# Tamaño y centrado
ancho_ventana = 900
alto_ventana = 600
pantalla_ancho = root.winfo_screenwidth()
pantalla_alto = root.winfo_screenheight()
x = (pantalla_ancho // 2) - (ancho_ventana // 2)
y = (pantalla_alto // 2) - (alto_ventana // 2)
root.geometry(f"{ancho_ventana}x{alto_ventana}+{x}+{y}")

# === Elementos de la interfaz ===
lbl_titulo = tk.Label(root, text="BotHealthy Beta", font=("Segoe UI", 18, "bold"))
lbl_titulo.pack(pady=10)

btn_cargar = tk.Button(
    root,
    text="📂 Cargar archivo Excel",
    command=abrir_excel,
    font=("Segoe UI", 11),
    background="green",
    foreground="white",
)
btn_cargar.pack(pady=5)

lbl_codigo = tk.Label(
    root, text="💻 Código de clínica detectado: ----", font=("Segoe UI", 11)
)
lbl_codigo.pack(pady=5)

# Frame de tablas
frame_tablas = tk.Frame(root)
frame_tablas.pack(pady=10, fill="both", expand=True)

# Tabla de facturas
frame_facturas = tk.Frame(frame_tablas)
frame_facturas.pack(side="top", fill="both", expand=True)

lbl_facturas = tk.Label(
    frame_facturas, text="📑 Facturas detectadas", font=("Segoe UI", 12, "bold")
)
lbl_facturas.pack()

columnas_fact = ("N° Factura", "Fecha", "Empresa", "NIT")
tree_facturas = ttk.Treeview(
    frame_facturas, columns=columnas_fact, show="headings", height=6
)
for col in columnas_fact:
    tree_facturas.heading(col, text=col)
    tree_facturas.column(col, width=200)
tree_facturas.pack(side="left", fill="both", expand=True)

scroll_fact = ttk.Scrollbar(
    frame_facturas, orient="vertical", command=tree_facturas.yview
)
scroll_fact.pack(side="right", fill="y")
tree_facturas.configure(yscroll=scroll_fact.set)

tree_facturas.bind("<Double-1>", mostrar_productos)

# Tabla de productos
frame_productos = tk.Frame(frame_tablas)
frame_productos.pack(side="bottom", fill="both", expand=True, pady=10)

lbl_productos = tk.Label(
    frame_productos,
    text="🛒 Productos de la factura seleccionada",
    font=("Segoe UI", 12, "bold"),
)
lbl_productos.pack()

columnas_prod = ("Código Producto", "Nombre Producto", "Cantidad", "Precio")
tree_productos = ttk.Treeview(
    frame_productos, columns=columnas_prod, show="headings", height=6
)
for col in columnas_prod:
    tree_productos.heading(col, text=col)
    tree_productos.column(col, width=200)
tree_productos.pack(side="left", fill="both", expand=True)

scroll_prod = ttk.Scrollbar(
    frame_productos, orient="vertical", command=tree_productos.yview
)
scroll_prod.pack(side="right", fill="y")
tree_productos.configure(yscroll=scroll_prod.set)

# Botón iniciar
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

