# AutoFiller — Deploy a producción

Producción corre en la **misma VM Ubuntu 24.04 que Balance360 y FactuMov**
(`192.168.100.16`), detrás de **srv-nginx** (`192.168.100.9`), que termina el HTTPS y
proxea por HTTP a la VM. Son dos máquinas distintas y es fácil confundirlas: el
certificado y el server block van en la `.9`; el compose y el `.env`, en la `.16`.

```
navegador → srv-nginx .9 :443 → VM .16 :8002 → app (uvicorn + FastAPI)
                                                 ├─ /api/…  extracción
                                                 └─ /       la web (StaticFiles)
```

**Es el más simple de los tres deploys de la VM**: un solo servicio, sin base de datos
—el servidor no guarda nada—, sin SPA que compilar y sin Alembic.

**Y es solo la mitad de la app.** La otra mitad es el **agente**, que corre en la PC de
cada operador y es lo único que toca SISalud. Sin agente instalado, la web lee
comprobantes pero no carga ninguno. Ver § 5.

| Servicio | Qué es | Puerto |
|---|---|---|
| `app` | uvicorn + FastAPI, imagen de `servidor/Dockerfile` | **8002** |

---

## 1. Antes de empezar

- **DNS**: `autofiller.insoft.net.ar` tiene que apuntar a donde apunta `insoft.net.ar`,
  o sea a srv-nginx.
- **El 8000 es de Balance360 y el 8001 de FactuMov.** AutoFiller publica el **8002**. Si
  alguna vez hay que cambiarlo, se cambia en `docker-compose.prod.yml` y en el
  `proxy_pass` de srv-nginx.
- **En el gateway no hay que abrir el 8002.** Ese puerto no sale nunca de la LAN: es el
  salto de srv-nginx a la VM. La puerta desde internet es el 443 de la `.9`, ya
  reenviado desde antes.
- **Espacio en la VM.** Van a ser tres apps Python conviviendo. La imagen de AutoFiller
  no lleva librerías de sistema, pero suma:
  ```bash
  df -h /              # espacio
  docker system df     # cuánto de eso es basura de builds viejos
  ```
- **Rotar la contraseña de SISalud.** Salió del código el 2026-09-10, pero sigue en el
  commit `486e2a8` y adentro del `dist/AutoFiller.exe` ya distribuido. El repo es
  privado, así que el historial no es el problema urgente; el `.exe` que ya está en las
  PCs, sí.

> `chromedriver.exe` (18 MB, no se usaba) se sacó del árbol el 2026-09-10. Sigue en el
> historial, así que el `git clone` de la VM lo baja igual —comprimido, unos 9 MB— y la
> única forma de evitarlo sería reescribir el historial. No vale la pena por una vez.

---

## 2. Primer arranque en la VM

```bash
ssh johnny@192.168.100.16
git clone https://github.com/JohnnySalvati/Autofiller.git
cd Autofiller
cp .env.example .env
nano .env            # al menos una de las dos claves de lectura (ver abajo)
mkdir -p publicacion # de donde se sirve el agente (§ 5). Antes del up: si la crea
                     # Docker, queda de root y el scp del agente no entra.
docker compose -f docker-compose.prod.yml up -d --build
docker compose -f docker-compose.prod.yml logs -f app
```

**Las fotos necesitan una clave de lectura, y hay dos motores**: `GOOGLE_VISION_API_KEY`
(OCR de Cloud Vision, el camino por defecto: gratis hasta 1000 páginas por mes y
determinístico) y `ANTHROPIC_API_KEY` (modelo de visión de Claude, el respaldo para
cuando el OCR devuelve algo que los regex de ARCA no parsean). Con cualquiera de las dos
la app lee fotos; lo recomendado es tener las dos.

**Sin ninguna la app arranca igual** y lee PDFs, pero no comprobantes en foto — que es
justamente lo que destraba a los dos operadores que hoy no usan AutoFiller. La web lo
avisa con el chip «Fotos: no», y es fácil no mirarlo: verificalo con el punto 2 de la
§ 4.

Del `.env`, `docker compose` solo le pasa al contenedor las variables declaradas en
`env_file`. Una variable con el valor vacío llega como cadena vacía, que para
las dos claves de lectura es lo mismo que no tenerlas.

---

## 3. El server block de srv-nginx

Va en `192.168.100.9` (`administrator@192.168.100.9`, el mismo server de la landing).
La convención de ese server es un archivo por dominio en `sites-available/` con su
symlink en `sites-enabled/`.

Primero el archivo de contraseñas. Es la **única** autenticación del servidor: la
pantalla «Entrar» de la app pide las credenciales de SISalud, que viajan al agente en la
PC del operador y nunca al servidor, así que sin esto `/api/extraer` queda abierto a
internet y cualquiera puede gastar tu cuota de lectura de comprobantes.

```bash
sudo apt install apache2-utils          # si no está
sudo htpasswd -c /etc/nginx/.htpasswd-autofiller operador
sudo htpasswd    /etc/nginx/.htpasswd-autofiller otro-operador   # el -c crea de cero y pisa: va solo la primera vez
```

`/etc/nginx/sites-available/autofiller.insoft.net.ar`, con `listen 80` y no `listen 443`:
un bloque con `listen 443 ssl` y sin `ssl_certificate` no pasa el `nginx -t`, y el
certificado todavía no existe. Certbot reescribe el archivo después.

```nginx
server {
    listen 80;
    server_name autofiller.insoft.net.ar;

    # Las DOS únicas rutas sin auth_basic, y es a propósito: son por donde el
    # agente de cada PC se entera de que hay una versión nueva y se la baja, y el
    # agente no tiene esas credenciales (pedírselas sería volver a meter un
    # secreto adentro de un ejecutable que se distribuye).
    #
    # Lo que queda expuesto es el instalador del agente, que no lleva ningún
    # secreto adentro: las credenciales de SISalud las tipea el operador y la
    # clave de lectura vive en el .env de la VM. El auth_basic sigue tapando
    # /api/extraer, que es lo que protege la cuota de lectura de comprobantes.
    #
    # Van con `=` (coincidencia exacta), que en nginx tiene prioridad sobre el
    # prefijo `location /`: no abren nada más que estas dos direcciones.
    location = /api/agente {
        proxy_pass http://192.168.100.16:8002;
        proxy_set_header Host $host;
    }

    location = /descargas/AutoFillerAgente.zip {
        proxy_pass http://192.168.100.16:8002;
        proxy_set_header Host $host;
        # 60 MB por una conexión de oficina pueden pasar del minuto, y el default
        # de proxy_read_timeout es 60 s: sin esto la descarga se corta por la
        # mitad y el agente la descarta por el sha256, una y otra vez.
        proxy_read_timeout 300s;
    }

    location / {
        # El auth_basic va ACÁ ADENTRO y no a nivel server, y no es cuestión de estilo:
        # certbot valida el dominio por HTTP-01 insertando un `location = /.well-known/…`
        # en este mismo server block. Las directivas de un location no se heredan a otro
        # location, pero las del server sí: puesto arriba, el desafío contestaría 401 y
        # la RENOVACIÓN fallaría en silencio hasta que el certificado venza.
        auth_basic           "AutoFiller";
        auth_basic_user_file /etc/nginx/.htpasswd-autofiller;

        proxy_pass http://192.168.100.16:8002;
        proxy_set_header Host              $host;
        proxy_set_header X-Real-IP         $remote_addr;
        proxy_set_header X-Forwarded-For   $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;

        # La web manda un comprobante por pedido, así que el cuerpo tiene el techo del
        # AUTOFILLER_TAMANIO_MAXIMO_MB de la app (25 MB) más el sobre del multipart. El
        # default de nginx es 1 MB, o sea que sin esta línea cualquier foto de celular
        # muere en un 413 antes de llegar a la app.
        client_max_body_size 30m;

        # Cada pedido es un comprobante, pero si es foto hay una llamada al modelo de
        # visión en el medio. Los 60 s de default alcanzan casi siempre y fallan justo
        # cuando el modelo tarda: el operador se come un 504 con la lectura ya pagada.
        proxy_read_timeout 120s;
    }
}
```

```bash
sudo ln -s /etc/nginx/sites-available/autofiller.insoft.net.ar /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
sudo certbot --nginx -d autofiller.insoft.net.ar
sudo nginx -t && sudo systemctl reload nginx
```

Los `X-Forwarded-*` acá son informativos: AutoFiller no usa la IP del cliente para nada
(no hay rate limiter ni sesiones propias). Van igual porque son la convención de los
otros dos, y porque el día que hagan falta no se acuerda nadie.

---

## 4. Verificación después de un deploy

Ninguna de estas cosas avisa sola si está mal.

```bash
# 1. El contenedor arriba y "healthy" (tarda hasta 20 s en pasar de "starting")
docker compose -f docker-compose.prod.yml ps

# 2. La lectura de fotos está habilitada, y con qué motor. Si "lectura_de_fotos" es
#    false, no hay NINGUNA de las dos claves en el .env, y el servidor no lo dice por
#    ningún otro lado. (Nada de `head -c`: los 24 centros de costo ocupan el medio del
#    JSON y el campo que importa queda después.)
curl -s localhost:8002/api/opciones | grep -o '"lectura_de_fotos":[a-z]*\|"motor_lectura":"[^"]*"'

# 3. La web se sirve
curl -s -o /dev/null -w '%{http_code}\n' localhost:8002/
```

Y desde afuera, con el dominio: `https://autofiller.insoft.net.ar/` tiene que pedir
usuario y contraseña, y recién después mostrar la pantalla de Entrar.

Lo que ninguna verificación técnica cubre: **subir un comprobante de verdad** y ver que
vuelven los campos. Es lo único que ejercita la extracción entera.

---

## 5. El agente, en la PC de cada operador

El agente es lo único que toca SISalud y **no se despliega en la VM**: necesita el Chrome
y la sesión del operador.

Lo que hay que configurarle sí o sí antes de arrancarlo es de dónde se sirve la web, o va
a rechazar los pedidos:

```bat
set AUTOFILLER_ORIGENES=https://autofiller.insoft.net.ar
```

Es una lista blanca a propósito y **nunca un comodín**: el agente escucha en `127.0.0.1`,
así que sin ella cualquier página abierta en el navegador podría hacerle cargar
comprobantes.

Tiene que quedar abierto mientras se usa la web — acceso directo en el Inicio de Windows,
o el operador se va a encontrar con «No se detecta el agente».

### Instalar el agente en la PC de un operador

No necesita Python ni pip: es un `.exe`.

1. Que baje `AutoFillerAgente.zip` (60 MB) de
   `https://autofiller.insoft.net.ar/descargas/AutoFillerAgente.zip` —esa dirección no
   pide usuario ni contraseña— y lo descomprima donde quiera que viva, por ejemplo
   `C:\AutoFiller`. Es la única vez que hay que hacer esto: de ahí en más se actualiza
   solo.
2. Doble clic en **`configurar.bat`**. Deja `AUTOFILLER_ORIGENES` como variable de
   usuario y crea el acceso directo en la carpeta Inicio, para que el agente arranque
   con Windows. Acepta el origen como argumento si hay que apuntar a otro lado.
3. Doble clic en **`AutoFillerAgente.exe`**. No se abre ninguna ventana: queda como
   ícono al lado del reloj, atrás de la flechita de «Iconos ocultos». Lo que antes se
   veía en la consola está en `%LOCALAPPDATA%\AutoFiller\agente.log`, y el menú del
   ícono lo muestra en vivo («Ver la actividad»).

**Para verificar la instalación sin cargar un comprobante de verdad**:
`AutoFillerAgente.exe --probar`. Ejercita las dos cosas que pueden faltar en una PC
nueva —el driver de Playwright, que es un `node.exe` empaquetado, y la ruta de
Chrome— y dice qué falta. No abre el navegador ni toca SISalud.

Chrome tiene que estar instalado; si no está en la ruta estándar, se le indica con
`AUTOFILLER_CHROME`.

### Generar el zip

`agente\empaquetar.bat` compila, corre el autodiagnóstico sobre lo compilado —si falla,
no genera el zip— y deja `AutoFillerAgente.zip` listo.

Dos cosas del build que no son opcionales y están en el script: **onedir** y no
`--onefile` (arranca más rápido y los antivirus lo marcan menos), y
**`--collect-all playwright`**, porque Playwright trae su propio `node.exe` y
PyInstaller no se lo lleva solo — sin eso el agente compila bien y falla recién al
intentar abrir Chrome, en la PC del operador. Los navegadores que Playwright descarga
siguen sin hacer falta: el agente se engancha por CDP al Chrome que ya está.

### Publicar una versión nueva del agente

Desde la 2.2 **el agente se actualiza solo**: no hay que ir PC por PC ni pasar el zip por
mail. Subir una versión es esto:

```bat
agente\empaquetar.bat
```

que compila, corre el autodiagnóstico, arma `AutoFillerAgente.zip` y escribe
`AutoFillerAgente.json` (versión + sha256). Y después, a la VM:

```bash
scp AutoFillerAgente.zip AutoFillerAgente.json johnny@192.168.100.16:~/Autofiller/publicacion/
```

Eso es todo: **sin rebuild y sin restart**, porque el servidor lee esa carpeta en cada
consulta. Cada agente instalado lo ve dentro de las 4 horas, o antes si la PC se reinicia.

- La carpeta `publicacion/` está montada en el contenedor como `/publicacion` (ver
  `docker-compose.prod.yml`) y **tiene que existir antes del primer `up`**, o Docker la
  crea como root y el `scp` después rebota.
- **Lo que se compara es `VERSION` en `agente/main.py`.** Un zip nuevo con la misma
  versión no actualiza a nadie: para el agente instalado no hay nada nuevo.
- Verificar que quedó publicado, desde afuera y sin credenciales (tienen que ser las dos
  únicas rutas que contestan sin el usuario y la contraseña):

  ```bash
  curl -s https://autofiller.insoft.net.ar/api/agente
  curl -sI https://autofiller.insoft.net.ar/descargas/AutoFillerAgente.zip | head -3
  ```

**Cómo se aplica, en la PC del operador** (`agente/actualizacion.py`): baja el zip,
verifica el sha256, lo descomprime en `%LOCALAPPDATA%\AutoFiller\actualizacion`, deja
andando un `.bat` suelto y se apaga; el `.bat` espera a que el `.exe` cierre —Windows no
deja reemplazar un ejecutable tomado—, copia encima de la instalación y vuelve a
arrancar el agente. El ícono de la bandeja desaparece unos segundos y vuelve con la
versión nueva. Queda registrado en `%LOCALAPPDATA%\AutoFiller\actualizacion.log`.

**Nunca se actualiza a mitad de camino**: solo con el agente libre y sin sesión de
SISalud abierta (reiniciar borra las credenciales, que viven en memoria). La ventana que
siempre existe es el arranque de Windows.

> **El salto a la 2.2 hay que hacerlo a mano, una sola vez.** Los agentes instalados hoy
> son 2.1 y no traen la actualización automática: no hay nada que les avise. Hay que
> pasarles el zip de la 2.2 (o el link de descarga) y que lo descompriman encima, como
> siempre. De esa en más, se actualizan solos.

**Marcha atrás**: publicar de nuevo el zip y el json de la versión anterior no alcanza
—las PCs solo van hacia adelante—, así que hay que subir el arreglo como versión nueva.
Para volver atrás en una PC puntual, descomprimirle el zip viejo encima a mano.

---

## 6. Deploys siguientes

En dev, antes de pushear: probar la extracción contra las muestras (ver `CLAUDE.md`,
«Cómo probar»). No hay tests automatizados todavía.

Y en la VM:

```bash
cd ~/Autofiller
git pull
docker compose -f docker-compose.prod.yml up -d --build
docker compose -f docker-compose.prod.yml logs -f app
```

Si `requirements.txt` no cambió, la capa de dependencias sale de cache y el build es de
segundos. No hay volúmenes ni base: un redeploy es reemplazar el contenedor y nada más.

---

## 7. Cosas que muerden

- **Deployaste y en la app no cambió nada**, sin error ni nada raro en los logs: casi
  siempre es un `up -d` **sin `--build`**. Si el `up` dice `Running` en lugar de
  `Recreated`, no se reconstruyó nada.
- **«No se detecta el agente» en la pantalla de Entrar.** Por orden de probabilidad: el
  agente no está abierto en esa PC; está abierto pero sin `AUTOFILLER_ORIGENES`
  apuntando a `https://autofiller.insoft.net.ar`; o Chrome está bloqueando la llamada de
  una página HTTPS a `http://127.0.0.1:8765`. Las dos primeras se ven en la consola del
  navegador como error de CORS. La tercera es la que hay que probar **antes** de dar el
  despliegue por bueno: `127.0.0.1` es origen confiable para Chrome y el permiso de
  Private Network Access ya está resuelto en el código (`allow_private_network=True`),
  pero las versiones recientes sumaron un prompt de acceso a red local que puede
  aparecerle al operador.
- **413 al subir un comprobante.** Hay dos techos y corta el más bajo: el
  `client_max_body_size` de srv-nginx (30m) y el `AUTOFILLER_TAMANIO_MAXIMO_MB` de la
  app (25 MB). El que tiene que contestar el caso normal es el segundo, que es el único
  que trae un mensaje en castellano en la fila del comprobante. Si subís uno, subí el
  otro.
- **504 leyendo una foto.** Es el `proxy_read_timeout`. Como la web manda un
  comprobante por pedido, un 504 es una foto sola que tardó de más, no una tanda: se
  reintenta esa y las demás no se tocan.
- **«Fotos: no» en la web.** No hay ninguna de las dos claves de lectura
  (`GOOGLE_VISION_API_KEY` / `ANTHROPIC_API_KEY`) en el `.env` de la VM. Si el chip dice
  «Fotos: sí», al pasarle el mouse aclara con cuál de los dos motores se van a leer. Ojo:
  `servidor/.env` existe en la máquina de desarrollo pero **no lo lee nadie** —el
  proyecto no usa `python-dotenv`, las claves salen del entorno—, así que copiarlo a la
  VM no alcanza. Lo que manda es el `.env` de la raíz, que lee `docker compose`.
- **La pantalla de SISalud queda trabada con la máscara gris puesta.** No es del deploy:
  es GeneXus. Se libera con F5, y suele pasar por tener varias pestañas de la Carga
  Rápida abiertas. Ver `CLAUDE.md`.
