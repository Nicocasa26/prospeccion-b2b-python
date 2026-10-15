# -*- coding: utf-8 -*-
"""
Prospección B2B - Versión 3
---------------------------
Ventana de escritorio (tkinter). Buscás por TAG (transporte, fabrica, industria...)
en todas las zonas del GBA/CABA, o solo en las zonas que escribas.

Instalación (una sola vez):
    pip install requests ddgs

Ejecución:
    python prospeccion_b2b.py        (o botón "Run Python File" en VS Code)

Archivos que genera (junto a este script):
    prospectos.csv            -> Categoría ; Nombre ; Correo ; Otros correos
    historial_busquedas.txt   -> búsquedas (tag + zona) ya completadas
    dominios_visitados.txt    -> webs ya revisadas (no se vuelven a abrir)
"""

import csv
import html
import os
import queue
import re
import subprocess
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from urllib.parse import urljoin, urlparse
import tkinter as tk
from tkinter import scrolledtext, messagebox

import requests

try:
    from ddgs import DDGS
except ImportError:  # compatibilidad con el paquete viejo
    from duckduckgo_search import DDGS

# =========================================================
# CONFIGURACIÓN
# =========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(BASE_DIR, "prospectos.csv")
HISTORIAL_PATH = os.path.join(BASE_DIR, "historial_busquedas.txt")
DOMINIOS_PATH = os.path.join(BASE_DIR, "dominios_visitados.txt")

CABECERAS = ["Categoría", "Nombre", "Correo", "Otros correos"]
# Excel en Argentina abre bien los CSV separados por ";" (con "," queda todo en una columna).
DELIM = ";"

MAX_CORREOS_POR_EMPRESA = 5
HILOS_RASTREO = 6

ZONAS_PREDETERMINADAS = [
    # Zona Sur
    "Quilmes", "Bernal", "Don Bosco", "Ezpeleta", "Villa la Florida", "San Francisco Solano",
    "Berazategui", "Florencio Varela", "Avellaneda", "Sarandí", "Wilde", "Gerli", "Lanús",
    "Remedios de Escalada", "Banfield", "Lomas de Zamora", "Temperley", "Turdera", "Llavallol",
    "Luis Guillón", "Monte Grande", "El Jagüel", "Canning", "Ezeiza", "Tristán Suárez",
    "Carlos Spegazzini", "Adrogué", "Burzaco", "Longchamps", "Glew", "Claypole",
    "Rafael Calzada", "Guernica",
    # Zona Oeste
    "González Catán", "Gregorio de Laferrere", "Isidro Casanova", "Ciudad Evita", "La Tablada",
    "San Justo", "Ramos Mejía", "Ciudadela", "Haedo", "Morón", "Castelar", "Ituzaingó",
    "Padre Stefenelli", "Merlo", "Paso del Rey", "Moreno", "El Palomar", "Caseros", "Hurlingham",
    # CABA / Límite General Paz
    "Buenos Aires", "Liniers", "Villa Devoto", "Saavedra", "General San Martín", "Villa Ballester",
    # Zona Norte
    "Vicente López", "Olivos", "Martínez", "San Isidro", "Victoria", "San Fernando", "Tigre",
]

# Directorios, redes, portales de empleo, diarios: no son la web de la empresa.
DOMINIOS_BLOQUEADOS = (
    "facebook.com", "instagram.com", "linkedin.com", "youtube.com", "twitter.com", "x.com",
    "tiktok.com", "pinterest.com", "wikipedia.org", "google.com", "bing.com", "yahoo.com",
    "paginasamarillas.com.ar", "mercadolibre.com.ar", "mercadolibre.com", "olx.com.ar",
    "cylex-argentina.com", "cylex.com.ar", "guialocal.com.ar", "yelp.com", "tripadvisor.com",
    "tripadvisor.com.ar", "foursquare.com", "mapcarta.com", "cybo.com", "infoisinfo.com.ar",
    "computrabajo.com.ar", "bumeran.com.ar", "zonajobs.com.ar", "indeed.com", "glassdoor.com",
    "cuitonline.com", "dateas.com", "opencorporates.com", "solostocks.com.ar", "yellowpages.com",
    "infobae.com", "clarin.com", "lanacion.com.ar", "pagina12.com.ar", "ambito.com",
    "cronista.com", "wikiwand.com", "scribd.com", "slideshare.net", "issuu.com",
    "amazon.com", "aliexpress.com", "booking.com", "despegar.com", "waze.com",
)
SUFIJOS_BLOQUEADOS = (".gob.ar", ".gov.ar", ".edu.ar", ".mil.ar")

DOMINIOS_BASURA_EMAIL = (
    "example.com", "email.com", "domain.com", "yourdomain.com", "sentry.io", "wixpress.com",
    "godaddy.com", "wordpress.com", "wordpress.org", "w3.org", "schema.org", "gstatic.com",
    "googleapis.com", "cloudflare.com", "sentry-next.wixpress.com", "test.com", "correo.com",
)
PROVEEDORES_GRATIS = (
    "gmail.com", "hotmail.com", "hotmail.com.ar", "outlook.com", "outlook.com.ar", "yahoo.com",
    "yahoo.com.ar", "live.com", "live.com.ar", "icloud.com", "speedy.com.ar", "fibertel.com.ar",
    "arnet.com.ar", "ciudad.com.ar", "msn.com",
)
EXTENSIONES_NO_EMAIL = (
    ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".css", ".js", ".ico", ".woff",
    ".woff2", ".ttf", ".pdf", ".mp4", ".avif",
)
PREFIJOS_PREFERIDOS = ("info", "contacto", "ventas", "comercial", "administracion", "consultas", "rrhh")
PALABRAS_CONTACTO = ("contact", "contacto", "nosotros", "empresa", "about", "quienes",
                     "ubicacion", "sucursal", "rrhh", "trabaja")
NOMBRES_GENERICOS = {"inicio", "home", "bienvenidos", "bienvenido", "contacto", "principal",
                     "página principal", "pagina principal", "index", "welcome"}

PATRON_EMAIL = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9\-]+(?:\.[a-zA-Z0-9\-]+)*\.[a-zA-Z]{2,}")
HEADERS_HTTP = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/125.0 Safari/537.36"),
    "Accept-Language": "es-AR,es;q=0.9",
}


# =========================================================
# UTILIDADES DE URL / EMAIL
# =========================================================
def normalizar_url(url):
    url = (url or "").strip()
    if not url:
        return ""
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url.lstrip("/")
    return url


def dominio_de(url):
    try:
        netloc = urlparse(url).netloc.lower().split(":")[0]
    except Exception:
        return ""
    return netloc[4:] if netloc.startswith("www.") else netloc


def dominio_bloqueado(dom):
    if not dom:
        return True
    if dom.endswith(SUFIJOS_BLOQUEADOS):
        return True
    return any(dom == b or dom.endswith("." + b) for b in DOMINIOS_BLOQUEADOS)


def _decodificar_cfemail(hex_str):
    try:
        clave = int(hex_str[:2], 16)
        return bytes(int(hex_str[i:i + 2], 16) ^ clave for i in range(2, len(hex_str), 2)).decode("utf-8")
    except Exception:
        return ""


def preparar_texto(raw):
    """Desescapa HTML, decodifica mails protegidos por Cloudflare y formas '[at]' / '(arroba)'."""
    texto = html.unescape(raw or "")
    texto = texto.replace("%40", "@")
    extras = [_decodificar_cfemail(h) for h in re.findall(r'data-cfemail="([0-9a-fA-F]+)"', texto)]
    texto = re.sub(r"\s*[\[\(\{]\s*(?:at|arroba)\s*[\]\)\}]\s*", "@", texto, flags=re.I)
    texto = re.sub(r"\s*[\[\(\{]\s*(?:dot|punto)\s*[\]\)\}]\s*", ".", texto, flags=re.I)
    return texto + "\n" + "\n".join(e for e in extras if e)


def normalizar_email(email):
    email = (email or "").strip().strip(" <>[](){}\"'.,;:").lower()
    if not PATRON_EMAIL.fullmatch(email):
        return ""
    local, dominio = email.split("@", 1)
    if email.endswith(EXTENSIONES_NO_EMAIL):
        return ""
    if any(dominio == d or dominio.endswith("." + d) for d in DOMINIOS_BASURA_EMAIL):
        return ""
    if any(x in local for x in ("noreply", "no-reply", "donotreply", "webmaster", "postmaster")):
        return ""
    return email


def email_pertenece_al_sitio(email, dom_sitio):
    """Acepta mails del mismo dominio que la web o de proveedores gratuitos (gmail, hotmail...)."""
    dom_email = email.split("@", 1)[1]
    if dom_email in PROVEEDORES_GRATIS:
        return True
    if not dom_sitio:
        return False
    return dom_email == dom_sitio or dom_email.endswith("." + dom_sitio) or dom_sitio.endswith("." + dom_email)


def extraer_emails(texto, dom_sitio):
    encontrados = []
    for bruto in PATRON_EMAIL.findall(preparar_texto(texto)):
        email = normalizar_email(bruto)
        if email and email not in encontrados and email_pertenece_al_sitio(email, dom_sitio):
            encontrados.append(email)
    return encontrados


def ordenar_emails(emails, dom_sitio):
    def clave(e):
        local, dom = e.split("@", 1)
        mismo = 0 if (dom == dom_sitio or dom.endswith("." + dom_sitio)) else 1
        try:
            prio = next(i for i, p in enumerate(PREFIJOS_PREFERIDOS) if local.startswith(p))
        except StopIteration:
            prio = 99
        return (mismo, prio, e)
    return sorted(emails, key=clave)


# =========================================================
# DESCARGA Y RASTREO DE WEBS
# =========================================================
_local = threading.local()


def _sesion():
    if not hasattr(_local, "s"):
        _local.s = requests.Session()
        _local.s.headers.update(HEADERS_HTTP)
    return _local.s


def descargar(url, timeout=(5, 9)):
    """Devuelve (texto_html, url_final) o (None, url) si falla."""
    try:
        r = _sesion().get(url, timeout=timeout, allow_redirects=True)
        ctype = r.headers.get("Content-Type", "").lower()
        if r.status_code >= 400 or (ctype and "html" not in ctype):
            return None, r.url
        codificacion = r.encoding if "charset" in ctype and r.encoding else "utf-8"
        try:
            texto = r.content[:1_500_000].decode(codificacion, errors="replace")
        except LookupError:
            texto = r.content[:1_500_000].decode("utf-8", errors="replace")
        return texto, r.url
    except Exception:
        return None, url


def limpiar_titulo(titulo):
    titulo = re.sub(r"\s+", " ", html.unescape(titulo or "")).strip()
    partes = re.split(r"\s+[\-|–—:·•»]\s+|\s*\|\s*", titulo)
    for p in partes:
        p = p.strip()
        if len(p) >= 3 and p.lower() not in NOMBRES_GENERICOS:
            return p[:80]
    return titulo[:80]


def extraer_nombre(html_text, dom):
    patrones = (
        r'<meta[^>]+property=["\']og:site_name["\'][^>]+content=["\']([^"\']+)',
        r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:site_name["\']',
        r"<title[^>]*>(.*?)</title>",
    )
    for pat in patrones:
        m = re.search(pat, html_text or "", re.I | re.S)
        if m:
            nombre = limpiar_titulo(m.group(1))
            if nombre and nombre.lower() not in NOMBRES_GENERICOS:
                return nombre
    return ""


def rastrear_sitio(url):
    """Abre la home y hasta 3 páginas de contacto. Devuelve (nombre, [emails]) o None si no abre."""
    html0, url_final = descargar(url)
    if html0 is None:
        return None
    dom = dominio_de(url_final)
    textos = [html0]

    candidatos = []
    for href in re.findall(r"""href\s*=\s*["']([^"'#]+)""", html0, re.I):
        h = href.lower()
        if h.startswith(("mailto:", "tel:", "javascript:")):
            continue
        if any(p in h for p in PALABRAS_CONTACTO):
            destino = urljoin(url_final, href)
            if dominio_de(destino) == dom and destino != url_final and destino not in candidatos:
                candidatos.append(destino)
    if not candidatos:
        candidatos = [urljoin(url_final, "/contacto"), urljoin(url_final, "/contact")]

    for destino in candidatos[:3]:
        t, _ = descargar(destino, timeout=(4, 7))
        if t:
            textos.append(t)

    emails = []
    for t in textos:
        for e in extraer_emails(t, dom):
            if e not in emails:
                emails.append(e)
    emails = ordenar_emails(emails, dom)[:MAX_CORREOS_POR_EMPRESA]
    return extraer_nombre(html0, dom), emails


# =========================================================
# INTERFAZ
# =========================================================
class AppProspeccion:
    def __init__(self, root):
        self.root = root
        self.root.title("Prospección B2B - Buscador de contactos")
        self.root.geometry("960x720")
        self.root.minsize(820, 600)
        self.root.configure(bg="#1e1e1e")

        self.cola = queue.Queue()
        self.ejecutando = False
        self.deadline = None
        self.lock = threading.Lock()
        self.buffer_csv = []
        self.emails_vistos = set()
        self.dominios = set()
        self.reservados = set()
        self.n_empresas = 0
        self.n_correos = 0

        self._construir_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.root.after(100, self._vaciar_cola)

    # ---------- construcción de la ventana ----------
    def _construir_ui(self):
        fg, bg, bg2 = "#ffffff", "#1e1e1e", "#2d2d2d"

        head = tk.Frame(self.root, bg=bg2, pady=10, padx=12)
        head.pack(fill=tk.X)
        tk.Label(head, text="🚀 Prospección B2B por TAG y zona", font=("Consolas", 15, "bold"),
                 fg="#00ffcc", bg=bg2).pack(anchor=tk.W)

        cfg = tk.LabelFrame(self.root, text=" Búsqueda ", font=("Segoe UI", 10, "bold"),
                            fg=fg, bg=bg, padx=10, pady=10)
        cfg.pack(fill=tk.X, padx=15, pady=8)
        cfg.columnconfigure(1, weight=1)

        def fila(r, texto, valor="", ancho=None):
            tk.Label(cfg, text=texto, font=("Segoe UI", 9, "bold"), fg="#cccccc", bg=bg).grid(
                row=r, column=0, sticky=tk.W, pady=4)
            e = tk.Entry(cfg, font=("Segoe UI", 10), bg=bg2, fg=fg, insertbackground="white",
                         relief=tk.FLAT, width=ancho or 60)
            e.grid(row=r, column=1, sticky=tk.W if ancho else tk.EW, padx=10, pady=4)
            e.insert(0, valor)
            return e

        self.ent_tags = fila(0, "Tags (separados por coma):", "transporte, fabrica, industria")
        self.ent_zonas = fila(1, "Zonas (vacío = todas):", "")
        self.ent_horas = fila(2, "Duración en horas (0 = sin límite):", "3", ancho=8)

        tk.Label(cfg, text=(f"💡 Con 'Zonas' vacío busca en las {len(ZONAS_PREDETERMINADAS)} localidades "
                            "(Sur, Oeste, CABA y Norte). Si escribís zonas, busca solo ahí."),
                 font=("Segoe UI", 8, "italic"), fg="#00ffcc", bg=bg).grid(
            row=3, column=1, sticky=tk.W, padx=10)

        barra = tk.Frame(self.root, bg=bg)
        barra.pack(fill=tk.X, padx=15, pady=4)

        def boton(texto, color, cmd, estado=tk.NORMAL):
            b = tk.Button(barra, text=texto, font=("Segoe UI", 10, "bold"), bg=color, fg="white",
                          activeforeground="white", relief=tk.FLAT, padx=12, pady=7,
                          command=cmd, state=estado)
            b.pack(side=tk.LEFT, padx=4)
            return b

        self.btn_iniciar = boton("▶ Iniciar", "#007acc", self.iniciar)
        self.btn_detener = boton("⏹ Detener", "#d9534f", self.detener, tk.DISABLED)
        self.btn_csv = boton("📂 Abrir CSV", "#444444", self.abrir_csv)
        self.btn_reset = boton("🗑 Borrar historial", "#444444", self.borrar_historial)

        self.lbl_estado = tk.Label(barra, text="Listo", font=("Segoe UI", 10, "bold"),
                                   fg="#00ffcc", bg=bg)
        self.lbl_estado.pack(side=tk.RIGHT, padx=8)

        self.lbl_stats = tk.Label(self.root, text="Consultas 0/0  ·  Empresas 0  ·  Correos 0",
                                  font=("Consolas", 10), fg="#cccccc", bg=bg)
        self.lbl_stats.pack(anchor=tk.W, padx=20)

        marco = tk.LabelFrame(self.root, text=" Monitor ", font=("Segoe UI", 10, "bold"),
                              fg=fg, bg=bg, padx=5, pady=5)
        marco.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        self.txt = scrolledtext.ScrolledText(marco, wrap=tk.WORD, font=("Consolas", 9),
                                             bg="#0d1117", fg="#00ff66", insertbackground="white")
        self.txt.pack(fill=tk.BOTH, expand=True)

    # ---------- comunicación hilo -> ventana ----------
    def log(self, msg):
        self.cola.put(("log", f"[{datetime.now():%H:%M:%S}] {msg}\n"))

    def _stats(self, hechas, total):
        self.cola.put(("stats", f"Consultas {hechas}/{total}  ·  Empresas {self.n_empresas}  ·  "
                                f"Correos {self.n_correos}"))

    def _vaciar_cola(self):
        try:
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "log":
                    self.txt.insert(tk.END, dato)
                    self.txt.see(tk.END)
                    if int(self.txt.index("end-1c").split(".")[0]) > 3000:
                        self.txt.delete("1.0", "500.0")
                elif tipo == "stats":
                    self.lbl_stats.config(text=dato)
                elif tipo == "fin":
                    self.btn_iniciar.config(state=tk.NORMAL)
                    self.btn_detener.config(state=tk.DISABLED)
                    self.lbl_estado.config(text="Listo", fg="#00ffcc")
        except queue.Empty:
            pass
        self.root.after(100, self._vaciar_cola)

    # ---------- botones ----------
    def iniciar(self):
        tags = []
        for t in re.split(r"[,;]", self.ent_tags.get()):
            t = t.strip()
            if t and t.casefold() not in [x.casefold() for x in tags]:
                tags.append(t)
        if not tags:
            messagebox.showerror("Falta el tag", "Escribí al menos un tag (ej.: transporte).")
            return

        zonas_txt = self.ent_zonas.get().strip()
        if zonas_txt:
            zonas = list(dict.fromkeys(z.strip() for z in re.split(r"[,;]", zonas_txt) if z.strip()))
        else:
            zonas = list(ZONAS_PREDETERMINADAS)

        try:
            horas = float(self.ent_horas.get().strip().replace(",", "."))
            if horas < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Duración inválida", "Poné un número de horas (ej.: 1, 2.5 o 0 para sin límite).")
            return

        self.ejecutando = True
        self.deadline = time.time() + horas * 3600 if horas > 0 else None
        self.btn_iniciar.config(state=tk.DISABLED)
        self.btn_detener.config(state=tk.NORMAL)
        self.lbl_estado.config(text="Buscando...", fg="#ffcc00")
        threading.Thread(target=self._ejecutar, args=(tags, zonas), daemon=True).start()

    def detener(self):
        self.ejecutando = False
        self.lbl_estado.config(text="Deteniendo...", fg="#ff4444")
        self.log("🛑 Detención pedida: termina lo que está en curso.")

    def abrir_csv(self):
        if not os.path.exists(CSV_PATH):
            messagebox.showinfo("Sin CSV", "Todavía no se generó prospectos.csv.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(CSV_PATH)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", CSV_PATH])
            else:
                subprocess.Popen(["xdg-open", CSV_PATH])
        except Exception as e:
            messagebox.showerror("Error", f"No pude abrir el archivo:\n{e}")

    def borrar_historial(self):
        if self.ejecutando:
            messagebox.showwarning("En curso", "Detené la búsqueda antes de borrar el historial.")
            return
        if messagebox.askyesno("Borrar historial",
                               "Se olvidan las búsquedas y webs ya revisadas, así que se van a repetir.\n"
                               "El CSV NO se borra y no se duplican correos que ya estén en él.\n\n¿Continuar?"):
            for p in (HISTORIAL_PATH, DOMINIOS_PATH):
                if os.path.exists(p):
                    os.remove(p)
            self.log("🗑 Historial borrado.")

    def _cerrar(self):
        self.ejecutando = False
        self.root.destroy()

    # ---------- archivos ----------
    @staticmethod
    def _cargar_lineas(path):
        if not os.path.exists(path):
            return set()
        with open(path, encoding="utf-8") as f:
            return {l.strip().lower() for l in f if l.strip()}

    @staticmethod
    def _agregar_linea(path, linea):
        with open(path, "a", encoding="utf-8") as f:
            f.write(linea + "\n")

    def _inicializar_csv(self):
        if os.path.exists(CSV_PATH) and os.path.getsize(CSV_PATH) > 0:
            with open(CSV_PATH, newline="", encoding="utf-8-sig") as f:
                primera = next(csv.reader(f, delimiter=DELIM), [])
            if primera == CABECERAS:
                return
            respaldo = CSV_PATH.replace(".csv", f"_respaldo_{datetime.now():%Y%m%d_%H%M%S}.csv")
            os.replace(CSV_PATH, respaldo)
            self.log(f"📦 El CSV existente tenía otro formato; quedó guardado como {os.path.basename(respaldo)}")
        with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f, delimiter=DELIM).writerow(CABECERAS)

    def _cargar_emails_existentes(self):
        vistos = set()
        try:
            with open(CSV_PATH, newline="", encoding="utf-8-sig") as f:
                for fila in list(csv.reader(f, delimiter=DELIM))[1:]:
                    for celda in fila[2:]:
                        for e in PATRON_EMAIL.findall(celda or ""):
                            vistos.add(e.lower())
        except Exception as e:
            self.log(f"⚠️ No pude leer el CSV existente: {e}")
        return vistos

    def _guardar_filas(self, filas, final=False):
        self.buffer_csv.extend(filas)
        if not self.buffer_csv:
            return
        try:
            with open(CSV_PATH, "a", newline="", encoding="utf-8-sig") as f:
                csv.writer(f, delimiter=DELIM).writerows(self.buffer_csv)
            self.buffer_csv.clear()
        except PermissionError:
            if final:
                alterno = os.path.join(BASE_DIR, f"prospectos_pendientes_{datetime.now():%Y%m%d_%H%M%S}.csv")
                with open(alterno, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.writer(f, delimiter=DELIM)
                    w.writerow(CABECERAS)
                    w.writerows(self.buffer_csv)
                self.log(f"⚠️ El CSV seguía abierto. Guardé {len(self.buffer_csv)} filas en {os.path.basename(alterno)}")
                self.buffer_csv.clear()
            else:
                self.log("⚠️ prospectos.csv está abierto en Excel: cerralo. Las filas se guardan en la próxima ronda.")

    # ---------- lógica de búsqueda ----------
    def _activo(self):
        return self.ejecutando and (self.deadline is None or time.time() < self.deadline)

    def _esperar(self, segundos):
        fin = time.time() + segundos
        while time.time() < fin and self._activo():
            time.sleep(0.5)

    def _buscar(self, consulta):
        """Devuelve lista de resultados ([] si no hay) o None si el buscador falló."""
        for intento in range(3):
            if not self._activo():
                return None
            try:
                with DDGS() as d:
                    return list(d.text(consulta, region="ar-es", max_results=10))
            except Exception as e:
                msg = str(e)
                if "no results" in msg.lower():
                    return []
                self.log(f"⚠️ Buscador (intento {intento + 1}/3): {msg[:120]}")
                self._esperar(20 * (intento + 1))
        return None

    def _trabajo(self, url):
        if not self._activo():
            return None
        return rastrear_sitio(url)

    def _procesar_resultados(self, resultados, tag):
        trabajos = []
        for r in resultados:
            url = normalizar_url(r.get("href", ""))
            dom = dominio_de(url)
            if dominio_bloqueado(dom):
                continue
            with self.lock:
                if dom in self.dominios or dom in self.reservados:
                    continue
                self.reservados.add(dom)
            trabajos.append((url, dom, r.get("title", "")))

        filas = []
        if not trabajos:
            self.log("   — Sin webs nuevas en esta búsqueda.")
            return filas

        tag_csv = tag[:1].upper() + tag[1:]
        with ThreadPoolExecutor(max_workers=HILOS_RASTREO) as ex:
            futuros = {ex.submit(self._trabajo, url): (dom, titulo) for url, dom, titulo in trabajos}
            for fut in as_completed(futuros):
                dom, titulo = futuros[fut]
                try:
                    res = fut.result()
                except Exception:
                    res = None
                if res is None:
                    self.log(f"   ✖ {dom}: no se pudo abrir")
                    continue
                nombre, emails = res
                # Web revisada: no se vuelve a abrir en futuras corridas.
                self.dominios.add(dom)
                self._agregar_linea(DOMINIOS_PATH, dom)

                nuevos = [e for e in emails if e not in self.emails_vistos]
                if not nuevos:
                    self.log(f"   · {dom}: sin correos nuevos")
                    continue
                nombre = nombre or limpiar_titulo(titulo) or dom.split(".")[0].title()
                self.emails_vistos.update(nuevos)
                filas.append([tag_csv, nombre, nuevos[0], ", ".join(nuevos[1:])])
                self.n_empresas += 1
                self.n_correos += len(nuevos)
                self.log(f"   ✅ {nombre} → {', '.join(nuevos)}")
        return filas

    def _ejecutar(self, tags, zonas):
        try:
            self._inicializar_csv()
            self.emails_vistos = self._cargar_emails_existentes()
            self.dominios = self._cargar_lineas(DOMINIOS_PATH)
            self.reservados = set()
            self.n_empresas = self.n_correos = 0
            historial = self._cargar_lineas(HISTORIAL_PATH)

            todas = [(t, z) for t in tags for z in zonas]
            pendientes = [(t, z) for t, z in todas if f"{t}|{z}".lower() not in historial]
            self.log(f"🏷  Tags: {', '.join(tags)}")
            self.log(f"📍 Zonas: {len(zonas)}  ·  Combinaciones: {len(todas)}  ·  "
                     f"Pendientes (no repetidas): {len(pendientes)}")
            if not pendientes:
                self.log("⚠️ Ya hiciste todas estas búsquedas. Probá otro tag/zona o 'Borrar historial'.")
                return

            fallos = 0
            for i, (tag, zona) in enumerate(pendientes, 1):
                if not self._activo():
                    break
                self._stats(i - 1, len(pendientes))
                self.log(f"\n🔎 [{i}/{len(pendientes)}] {tag} · {zona}")
                consulta = f"{tag} {zona} Buenos Aires empresa contacto"
                resultados = self._buscar(consulta)
                if resultados is None:
                    fallos += 1
                    if self._activo():
                        self.log("⚠️ Búsqueda fallida: no se marca como hecha, se reintentará en otra corrida.")
                        self._esperar(90 if fallos >= 3 else 15)
                    continue
                fallos = 0

                filas = self._procesar_resultados(resultados, tag)
                self._guardar_filas(filas)
                if filas:
                    self.log(f"💾 {len(filas)} empresa(s) guardada(s).")

                if self._activo():  # solo se marca completa si no se cortó a la mitad
                    self._agregar_linea(HISTORIAL_PATH, f"{tag}|{zona}".lower())
                self._esperar(1.5)

            self._stats(len(pendientes), len(pendientes))
        except Exception as e:
            self.log(f"❌ Error crítico: {e}")
        finally:
            try:
                self._guardar_filas([], final=True)
            except Exception as e:
                self.log(f"❌ No pude guardar filas pendientes: {e}")
            self.log(f"\n🏁 Fin. Empresas nuevas: {self.n_empresas} · Correos nuevos: {self.n_correos}")
            self.log(f"📄 Archivo: {CSV_PATH}")
            self.ejecutando = False
            self.cola.put(("fin", None))


if __name__ == "__main__":
    ventana = tk.Tk()
    AppProspeccion(ventana)
    ventana.mainloop()