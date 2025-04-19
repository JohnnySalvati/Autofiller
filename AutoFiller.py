# pyinstaller --add-data "C:\Users\Johnny\AppData\Local\ms-playwright;.local\ms-playwright" --onefile AutoFiller.py
# chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\ChromeProfile" --no-first-run --no-default-browser-check


import asyncio
from playwright.async_api import async_playwright
import tkinter as tk
from tkinter import filedialog
import pdfplumber
import re
import pikepdf
from datetime import date 
import subprocess
import time
import socket
import sys


async def main(usuario, contraseña, factura):
    chrome_path = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    user_data_dir = r"C:\\ChromeProfile"
    debugging_port = 9222


    chrome_process = subprocess.Popen([
        chrome_path,
        f"--remote-debugging-port={debugging_port}",
        f"--user-data-dir={user_data_dir}",
        "--no-first-run",
        "--no-default-browser-check"
    ], shell=True)

    # Función para verificar si el puerto está abierto
    def is_port_open(host: str, port: int, timeout: float = 0.5) -> bool:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(timeout)
            return sock.connect_ex((host, port)) == 0

    # Esperar hasta que el puerto 9222 esté disponible
    while not is_port_open("localhost", debugging_port):
        time.sleep(0.5)

    async with async_playwright() as p:

      # Conectarse al navegador en ejecución sin abrir una nueva ventana
        browser = await p.chromium.connect_over_cdp("http://localhost:9222")
        
        # Obtener el contexto de la sesión actual
        context = browser.contexts[0] if browser.contexts else await browser.new_context()
        page = await context.new_page()

        # Navegar a la aplicación
        await page.goto("http://vpn.aomaosam.org.ar:8081/sisaludevo/servlet/cargarapidacomprobantescompra?10,0")

        await page.get_by_role("textbox", name = "ingrese con la forma: usuario").fill(usuario)
        await page.locator("input[name=\"W0030vPASS\"]").fill(contraseña)
        await page.get_by_role("button", name = "entrar").click()
        
        await page.locator("a").filter(has_text="Prestadores").click()
        await page.get_by_role("cell", name="Carga Rapida Comprobantes").click()

        # Click the button to open the modal
        await page.locator("#PROMPTIMGENTIDAD").click()

        await page.locator("iframe[title=\"Promptentidad\\?34\\,\\,0\\,10\\,c\\,\\,gxPopupLevel\\%3D0\\%3B\"]").content_frame.locator("#vENTIDADCUIT").fill(factura.cuit)
        await page.locator("iframe[title=\"Promptentidad\\?34\\,\\,0\\,10\\,c\\,\\,gxPopupLevel\\%3D0\\%3B\"]").content_frame.get_by_role("button", name="Buscar").click()
        await page.locator("iframe[title=\"Promptentidad\\?34\\,\\,0\\,10\\,c\\,\\,gxPopupLevel\\%3D0\\%3B\"]").content_frame.get_by_role("link").click()

        await page.locator("#vTIPOCOMPROBANTECODIGO").click()
        lista_tipo_comprobante = page.locator("#vTIPOCOMPROBANTECODIGO")
        options = await lista_tipo_comprobante.locator("option").all()
        opcion = await options[factura.tipo_comprobante].text_content()
        await lista_tipo_comprobante.select_option(label=opcion)
        await page.evaluate('document.querySelector("#vTIPOCOMPROBANTECODIGO").dispatchEvent(new Event("change"))')
        await page.locator("#vTIPOCOMPROBANTECODIGO").click()
        await page.keyboard.press("Enter")
        
        await page.locator("#vCOMPROBANTEPREFIJO").fill(factura.punto_venta)
        await page.locator("#vCOMPROBANTECODIGO").fill(factura.nro_factura)
        await page.locator("#vCOMPROBANTEFECHARECEPCION").fill(factura.fecha_recepcion)
        await page.locator("#vCOMPROBANTEFECHAEMISION").fill(factura.fecha_emision)
        await page.locator("#vCOMPROBANTECUOTAVTO").fill(factura.fecha_vencimiento)
        await page.locator("#vCOMPROBANTEDEVENGAMIENTO").fill(factura.fecha_devengamiento)
        await page.locator("#vCOMPROBANTECAE").fill(factura.cae)

        # await page.locator("#vCOMPROBANTEDETALLEDESCRIPCION").fill(factura.descripcion)
        await page.locator("#vEXENTOCOMPROBANTEDETALLEDESCRIPCION").fill(factura.descripcion)
        # await page.locator("#vCOMPROBANTEDETALLEPRECIOUNITARIO").fill(factura.importe)
        await page.locator("#vEXENTOCOMPROBANTEDETALLEPRECIOUNITARIO").fill(factura.importe)
        # await page.locator("#IMAGE2").click()
        await page.locator("#IMAGE3").click()

def decrypt_pdf(input_path):
    """Decrypts a PDF and saves it to the decrypted_folder."""
    try:
        with pikepdf.open(input_path) as pdf:
            pdf.save("decrypted.pdf")  
            pdf.close()
        return True
    except pikepdf.PasswordError:
        print(f"Error: Could not open {input_path} - incorrect password or strong encryption.")
        error_message = f"Error: Could not open {input_path} - incorrect password or strong encryption."
        # error_console.insert(tk.END, error_message + "\n")
        # error_console.see(tk.END)
        return None
    except Exception as e:
        print(f"Error decrypting {input_path}: {e}")
        error_message = f"Error decrypting {input_path}: {e}"
        # error_console.insert(tk.END, error_message + "\n")
        # error_console.see(tk.END)
        return None
    
def extract_information():
    """Extracts information from a decrypted PDF and saves it with a new filename in output_folder."""
    tipo_comprobante, punto_venta, nro_factura, cuit, punto_venta, nro_factura, fecha_recepcion, fecha_emision, fecha_vencimiento, fecha_devengamiento, cae, descripcion = None, None, None, None, None, None, None, None, None, None, None, ""
    translate = {"6": 2,
                 "11": 2,
                 "15": 4}
    try:
        with pdfplumber.open("decrypted.pdf") as pdf:

            text = "".join(page.extract_text() or "" for page in pdf.pages)
            
            # Regex to find Codigo
            codigo_match = re.search(r'COD\.?\s*0*(\d+)', text, re.IGNORECASE)
            if codigo_match:
                tipo_comprobante = codigo_match.group(1)
            else:
                codigo_match = re.search(r'Codigo\s*nº\s*(\d+)', text, re.IGNORECASE)
                tipo_comprobante = codigo_match.group(1) if codigo_match else None
            # translate codigo 
            tipo_comprobante = translate[tipo_comprobante]


            cuit_match = re.search(r'CUIT\:?\s*(\d{11})', text)
            if cuit_match:
                cuit = cuit_match.group(1)
                # cuit1 = cuit_match1.group(1) if cuit_match1 and len(cuit_match1.group(1))==11 else None
            else:
                cuit_match = re.search(r':\s*(\d{2})-(\d{8})-(\d)', text)
                cuit = cuit_match.group(1) + cuit_match.group(2) + cuit_match.group(3) if cuit_match else None

            # if len(cuit) != 11:
            #     cuit = None

            punto_venta_match = re.search(r'Punto de Venta:\s*(\d+)', text)
            if punto_venta_match:
                nro_factura_match = re.search(r'Comp\. Nro:\s*(\d+)', text)
                punto_venta = punto_venta_match.group(1).lstrip('0') if punto_venta_match else None
                nro_factura = nro_factura_match.group(1).lstrip('0') if nro_factura_match else None
            else:
                punto_venta_factura_match = re.search(r'\D*0*(\d{5})\s*-\s*0*(\d{8})', text)
                if punto_venta_factura_match:
                    punto_venta = punto_venta_factura_match.group(1).lstrip('0')
                    nro_factura = punto_venta_factura_match.group(2).lstrip('0')
                else:
                    punto_venta = None
                    nro_factura = None

            fecha_recepcion =  date.today().strftime("%d/%m/%Y")

            fecha_emision_match = re.search(r'Fecha de Emisión:\s*(\d{2}/\d{2}/\d{4})', text)
            if fecha_emision_match:
                fecha_emision = fecha_emision_match.group(1)

            fecha_hasta_match = re.search(r'Hasta:\s*(\d{2}/\d{2}/\d{4})', text)
            if fecha_hasta_match:
                fecha_hasta = fecha_hasta_match.group(1)

            fecha_vencimiento = fecha_hasta
            fecha_devengamiento = fecha_hasta

            cae_match = re.search(r'CAE N°:\s*(\d+)', text)
            cae = cae_match.group(1) if cae_match else None

            regex = r"(?<=Subtotal)(.*?)(?=Subtotal)"
            descripcion_match = re.search(regex, text, re.DOTALL | re.IGNORECASE)
            if descripcion_match:
                descripcion = descripcion_match.group(1).strip()

            # Expresión regular para capturar el "Importe Total"
            regex_importe = r"Importe Total:\s*\$?\s*([\d,]+(?:\.\d+)?)"
            importe_match = re.search(regex_importe, text)
            if importe_match:
                importe = importe_match.group(1)

        if cuit and tipo_comprobante and punto_venta and nro_factura:
            return Factura(cuit,tipo_comprobante, punto_venta, nro_factura, fecha_recepcion, fecha_emision, fecha_vencimiento, fecha_devengamiento, cae, descripcion, importe )
        else:
            print(f"Error: CUIT: {cuit} Cod: {tipo_comprobante} Punto de Venta: {punto_venta} Nro Factura: {nro_factura}")
            # error_message = f"Error: {decrypted_path} CUIT: {cuit} Cod: {codigo} Punto de Venta: {punto_venta} Nro Factura: {nro_factura}"
            # error_console.insert(tk.END, error_message + "\n")
            # error_console.see(tk.END)
            return None

    except Exception as e:
        print(f"Error extracting information: {e}")
        # error_message = f"Error extracting information: {e} {decrypted_path}"
        # error_console.insert(tk.END, error_message + "\n")
        # error_console.see(tk.END)
        return None

class Factura():
    def __init__(self,cuit,tipo_comprobante, punto_venta, nro_factura, fecha_recepcion, fecha_emision, fecha_vencimiento, fecha_devengamiento, cae, descripcion, importe):
        self.cuit = cuit
        self.tipo_comprobante = tipo_comprobante
        self.punto_venta = punto_venta
        self.nro_factura = nro_factura
        self.fecha_recepcion = fecha_recepcion
        self.fecha_emision = fecha_emision
        self.fecha_vencimiento = fecha_vencimiento
        self.fecha_devengamiento = fecha_devengamiento
        self.cae = cae
        self.descripcion = descripcion
        self.importe = importe

def start_processing(factura):
    asyncio.run(main(user_var.get(), pass_var.get() , factura))
    sys.exit(0)

# Crear ventana principal
app = tk.Tk()
app.title("AutoFiller")
app.geometry("1100x380")
app.configure(bg="#f7f7f9")  # Fondo suave

# Variables de entrada
user_var = tk.StringVar()
pass_var = tk.StringVar()

user_var.set("MACEVEDO@OSAM")
pass_var.set("Leon5513")

# Estilo de colores
bg_color = "#ffffff"  # Blanco para los paneles
border_color = "#d1d5db"  # Gris claro para bordes
title_color = "#4b5563"  # Gris oscuro para el título
button_color = "#2563eb"  # Azul para botones
button_text_color = "#ffffff"  # Blanco para texto de botones

# Estilo de bordes
frame_style = {
    "bg": bg_color,
    "highlightthickness": 1,
    "highlightbackground": border_color
}

# Título
title_frame = tk.Frame(app, pady=10, **frame_style)
tk.Label(title_frame, text="AutoFiller", font=("Helvetica", 16), fg=title_color, bg=bg_color).pack()
tk.Label(title_frame, text="Desarrollado por InSoft", bg=bg_color).pack()
title_frame.pack(pady=10, fill="x", padx=10)

# Usuario y contraseña
user_frame = tk.Frame(app, pady=10, **frame_style)
tk.Label(user_frame, text="Usuario", bg=bg_color).grid(row=0, column=0, pady=5, padx=5, sticky="e")
tk.Entry(user_frame, textvariable=user_var).grid(row=0, column=1, pady=5, padx=5)
tk.Label(user_frame, text="Contraseña", bg=bg_color).grid(row=1, column=0, pady=5, padx=5, sticky="e")
tk.Entry(user_frame, textvariable=pass_var, show="*").grid(row=1, column=1, pady=5, padx=5)
user_frame.pack(pady=10, fill="x", padx=10)

# Selección de archivo
file_frame = tk.Frame(app, pady=10, **frame_style)
# tk.Label(file_frame, text="Seleccione el archivo PDF:", bg=bg_color).grid(row=0, column=0, pady=5, padx=5, sticky="w")
process_button = None
file_label= None

def select_pdf():
    global process_button, file_label  # Reference to the global variables

    # Remove previous labels and buttons
    if file_label:
        file_label.pack_forget()
    if process_button:
        process_button.pack_forget()

    file_path = filedialog.askopenfilename(filetypes=(("Facturas", "*.pdf"),))
    if file_path:
        decrypt_pdf(file_path)
        factura = extract_information()
        if factura:
            file_label = tk.Label(file_frame, text=f"Archivo seleccionado: {file_path}", bg=bg_color)
            file_label.pack(pady=5, padx=5)
            process_button = tk.Button(app, text="Procesar", command=lambda: start_processing(factura), bg=button_color, fg=button_text_color)
            process_button.pack(pady=10)
        else:
            error_label = tk.Label(file_frame, text="Error: Factura inválida.", fg="red", bg=bg_color)
            error_label.pack(pady=5, padx=5)

tk.Button(file_frame, text="Seleccionar", command=select_pdf, bg=button_color, fg=button_text_color).pack(pady=5, padx=5)
file_frame.pack(pady=10, fill="x", padx=10)

# Ejecutar la aplicación
app.mainloop()
