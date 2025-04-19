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


        """
        # browser = await p.firefox.launch(headless=False)
        browser = await p.chromium.launch(
            user_data_dir="c:\\ChromeProfile",
            # executable_path="C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            headless=False,
        #     # args=["--force-device-scale-factor=1", "--disable-features=OverlayScrollbar"]
            # args= ["--start-maximized"],
            channel="chrome"
        )
        
        context = await browser.new_context(viewport={"width": 1200, "height": 800})
        page = await browser.new_page()
     """
        # context = await browser.new_context()
        # page = await context.new_page()

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

        # await page.pause()

        # await page.locator("#vCOMPROBANTEDETALLEDESCRIPCION").fill(factura.descripcion)
        await page.locator("#vEXENTOCOMPROBANTEDETALLEDESCRIPCION").fill(factura.descripcion)
        # await page.locator("#vCOMPROBANTEDETALLEPRECIOUNITARIO").fill(factura.importe)
        await page.locator("#vEXENTOCOMPROBANTEDETALLEPRECIOUNITARIO").fill(factura.importe)
        # await page.locator("#IMAGE2").click()
        await page.locator("#IMAGE3").click()

        # await page.pause()
        input("Presiona enter para cerrar el navegador")
        # Close the browser
        # subprocess.run(["taskkill", "/F", "/IM", "chrome.exe"], shell=True)
        chrome_process.kill()
        await browser.close()