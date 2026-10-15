# -*- coding: utf-8 -*-
"""
Enviador de correos B2B (Gmail + Gemini)
----------------------------------------
Lee prospectos.csv (generado por prospeccion_b2b.py), te deja escribir el mail
(o pedirle un borrador a Gemini) y lo envía de a poco desde tu Gmail,
anotando a quién le mandó en envios.csv para no repetir nunca.

Instalación (una sola vez):
    pip install requests

Antes de usarlo necesitás una "contraseña de aplicación" de Gmail:
    1) Activá la verificación en 2 pasos en tu cuenta de Google.
    2) Entrá a https://myaccount.google.com/apppasswords y generá una (16 letras).
    3) Pegala en la pestaña "Cuenta y envío" (NO uses tu contraseña normal).

API key de Gemini: https://aistudio.google.com/apikey  (opcional, solo para redactar).
Podés ponerlas como variables de entorno GMAIL_APP_PASSWORD y GEMINI_API_KEY
para no tener que tipearlas. El programa NUNCA las guarda en disco.

Archivos (junto a este script):
    prospectos.csv      <- entrada (Categoría ; Nombre ; Correo ; Otros correos)
    envios.csv          <- registro: fecha, correo, estado (enviado/rebotado/baja/error)
    bajas.txt           <- correos que pidieron no recibir más
    config_mailer.json  <- último borrador y ajustes (sin contraseñas)
"""

import csv
import email
import imaplib
import json
import os
import queue
import random
import re
import smtplib
import ssl
import subprocess
import sys
import threading
import time
from datetime import datetime, timedelta
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid, parseaddr
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog

import requests

# =========================================================
# CONFIGURACIÓN
# =========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROSPECTOS_CSV = os.path.join(BASE_DIR, "prospectos.csv")
ENVIOS_CSV = os.path.join(BASE_DIR, "envios.csv")
BAJAS_TXT = os.path.join(BASE_DIR, "bajas.txt")
CONFIG_JSON = os.path.join(BASE_DIR, "config_mailer.json")

DELIM = ";"
CABECERAS_ENVIOS = ["Fecha", "Correo", "Nombre", "Categoría", "Asunto", "Estado", "Detalle"]
ESTADOS_FINALES = {"enviado", "rebotado", "baja"}  # a estos NO se les vuelve a escribir

SMTP_HOST, SMTP_PORT = "smtp.gmail.com", 465
IMAP_HOST = "imap.gmail.com"
MODELO_GEMINI_DEFECTO = "gemini-2.5-flash"

PATRON_EMAIL = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9\-]+(?:\.[a-zA-Z0-9\-]+)*\.[a-zA-Z]{2,}")
PATRON_BAJA = re.compile(
    r"\b(baja|unsubscribe|desuscrib\w*|no me (?:env\w+|escrib\w+|contact\w+)|"
    r"dejen? de (?:env\w+|escrib\w+)|no (?:deseo|quiero) recibir|quit\w+ de (?:la|su|sus) lista)\b",
    re.I)

CUERPO_EJEMPLO = (
    "Hola, equipo de {nombre}:\n\n"
    "Soy [tu nombre] de [tu empresa]. Nos dedicamos a [qué ofrecés] y creemos que puede "
    "servirles porque [motivo concreto].\n\n"
    "¿Les interesaría que les envíe más información?\n\n"
    "Saludos."
)
FIRMA_EJEMPLO = "[Tu nombre]\n[Tu empresa]\n[Teléfono]"
BAJA_EJEMPLO = (
    "Te escribimos porque este correo figura públicamente como contacto de tu empresa. "
    "Si no querés recibir más mensajes nuestros, respondé BAJA a este correo y te quitamos de la lista."
)

MESES_IMAP = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# =========================================================
# LÓGICA (sin interfaz)
# =========================================================
def leer_prospectos():
    """Devuelve [{'categoria','nombre','correos':[principal, otros...]}]"""
    filas = []
    if not os.path.exists(PROSPECTOS_CSV):
        return filas
    with open(PROSPECTOS_CSV, newline="", encoding="utf-8-sig") as f:
        lector = csv.reader(f, delimiter=DELIM)
        next(lector, None)
        for r in lector:
            if len(r) < 3 or not r[2].strip():
                continue
            otros = []
            if len(r) > 3:
                otros = [e.strip().lower() for e in re.split(r"[,;\s]+", r[3]) if e.strip()]
            filas.append({"categoria": r[0].strip(), "nombre": r[1].strip(),
                          "correos": [r[2].strip().lower()] + otros})
    return filas


def cargar_bajas():
    if not os.path.exists(BAJAS_TXT):
        return set()
    with open(BAJAS_TXT, encoding="utf-8") as f:
        return {l.strip().lower() for l in f if l.strip()}


def cargar_envios():
    datos = {"hechos": set(), "enviados": set(), "rebotados": set(), "hoy": 0, "total": 0}
    if not os.path.exists(ENVIOS_CSV):
        return datos
    hoy = datetime.now().strftime("%Y-%m-%d")
    with open(ENVIOS_CSV, newline="", encoding="utf-8-sig") as f:
        lector = csv.reader(f, delimiter=DELIM)
        next(lector, None)
        for r in lector:
            if len(r) < 6:
                continue
            fecha, correo, estado = r[0], r[1].strip().lower(), r[5].strip().lower()
            if estado in ESTADOS_FINALES:
                datos["hechos"].add(correo)
            if estado == "enviado":
                datos["enviados"].add(correo)
                datos["total"] += 1
                if fecha.startswith(hoy):
                    datos["hoy"] += 1
            if estado == "rebotado":
                datos["rebotados"].add(correo)
    return datos


class Registro:
    """Anota cada envío en envios.csv (a prueba de Excel abierto) y recuerda las bajas."""

    def __init__(self):
        self.lock = threading.Lock()
        self.pendientes = []
        self.recargar()

    def recargar(self):
        with self.lock:
            d = cargar_envios()
            self.hechos, self.enviados, self.rebotados = d["hechos"], d["enviados"], d["rebotados"]
            self.hoy, self.total = d["hoy"], d["total"]
            self.bajas = cargar_bajas()

    def agregar_baja(self, correo):
        correo = correo.strip().lower()
        with self.lock:
            if correo in self.bajas:
                return
            self.bajas.add(correo)
            with open(BAJAS_TXT, "a", encoding="utf-8") as f:
                f.write(correo + "\n")

    def anotar(self, correo, nombre, categoria, asunto, estado, detalle=""):
        """Devuelve False si no pudo escribir en disco todavía (queda en memoria)."""
        fila = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), correo, nombre, categoria,
                asunto, estado, detalle]
        with self.lock:
            if estado in ESTADOS_FINALES:
                self.hechos.add(correo)
            if estado == "enviado":
                self.enviados.add(correo)
                self.hoy += 1
                self.total += 1
            if estado == "rebotado":
                self.rebotados.add(correo)
            self.pendientes.append(fila)
            return self._volcar()

    def _volcar(self):
        if not self.pendientes:
            return True
        try:
            nuevo = not os.path.exists(ENVIOS_CSV) or os.path.getsize(ENVIOS_CSV) == 0
            with open(ENVIOS_CSV, "a", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=DELIM)
                if nuevo:
                    w.writerow(CABECERAS_ENVIOS)
                w.writerows(self.pendientes)
            self.pendientes.clear()
            return True
        except PermissionError:
            return False

    def cerrar(self):
        """Último intento de escritura; si falla, vuelca a un archivo alterno. Devuelve su ruta o None."""
        with self.lock:
            if self._volcar() or not self.pendientes:
                return None
            alterno = os.path.join(BASE_DIR, f"envios_pendientes_{datetime.now():%Y%m%d_%H%M%S}.csv")
            with open(alterno, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=DELIM)
                w.writerow(CABECERAS_ENVIOS)
                w.writerows(self.pendientes)
            self.pendientes.clear()
            return alterno


def renderizar(plantilla, nombre, categoria):
    return (plantilla.replace("{nombre}", nombre or "su empresa")
                     .replace("{categoria}", categoria or ""))


def construir_mensaje(cfg, destino, nombre, categoria, prueba=False):
    cuerpo = renderizar(cfg["cuerpo"], nombre, categoria).strip()
    texto = f"{cuerpo}\n\n{cfg['firma'].strip()}\n\n--\n{cfg['baja'].strip()}\n"
    asunto = renderizar(cfg["asunto"], nombre, categoria).strip()
    msg = EmailMessage()
    msg["From"] = formataddr((cfg["remitente"], cfg["gmail"]))
    msg["To"] = destino
    msg["Subject"] = ("[PRUEBA] " + asunto) if prueba else asunto
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=cfg["gmail"].split("@")[-1])
    # Encabezado estándar de baja: Gmail muestra el link "Cancelar suscripción".
    msg["List-Unsubscribe"] = f"<mailto:{cfg['gmail']}?subject=BAJA>"
    msg.set_content(texto)
    return msg


def parsear_gemini(texto):
    t = (texto or "").strip().replace("**", "")
    m = re.search(r"ASUNTO:\s*(.+?)\s*\n+\s*CUERPO:\s*(.*)", t, re.S | re.I)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return "", t


def llamar_gemini(api_key, modelo, prompt):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{modelo}:generateContent"
    r = requests.post(url, headers={"x-goog-api-key": api_key, "Content-Type": "application/json"},
                      json={"contents": [{"parts": [{"text": prompt}]}]}, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"Gemini respondió {r.status_code}: {r.text[:250]}")
    try:
        return r.json()["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError, ValueError):
        raise RuntimeError("Gemini no devolvió texto (¿modelo incorrecto o contenido bloqueado?).")


REGLAS_REDACCION = (
    "Reglas: español rioplatense (Argentina); máximo 120 palabras; texto plano; tono cordial, "
    "profesional y honesto; sin mayúsculas sostenidas, sin exceso de signos de exclamación, sin "
    "promesas exageradas ni urgencia falsa; sin links ni adjuntos; un solo pedido claro al final, "
    "en forma de pregunta; el saludo inicial debe usar el marcador {nombre} tal cual, con llaves "
    "(es el nombre de la empresa destinataria); NO incluyas firma ni texto de baja (el programa los "
    "agrega). El asunto debe ser corto, claro y reflejar el contenido real del mensaje, sin engaños.\n"
    "Respondé EXACTAMENTE con este formato:\nASUNTO: <asunto>\nCUERPO: <cuerpo>\n"
)


def prompt_borrador(idea):
    return ("Redactá un primer correo comercial B2B para enviar a empresas que no nos conocen.\n"
            + REGLAS_REDACCION + f"\nQuién escribe y qué ofrece: {idea}")


def prompt_mejora(asunto, cuerpo):
    return ("Mejorá la redacción del siguiente correo comercial B2B manteniendo su sentido y los "
            "marcadores {nombre} y {categoria} sin modificar.\n" + REGLAS_REDACCION
            + f"\nASUNTO ACTUAL: {asunto}\nCUERPO ACTUAL:\n{cuerpo}")


def _texto_del_mensaje(msg):
    partes = []
    for p in msg.walk():
        tipo = p.get_content_type()
        if tipo.startswith("text/"):
            carga = p.get_payload(decode=True)
            if carga:
                partes.append(carga.decode(p.get_content_charset() or "utf-8", errors="replace"))
        elif tipo == "message/delivery-status":
            partes.append(str(p))
    return "\n".join(partes)


def analizar_mensaje(msg, enviados):
    """('rebote', [correos]) si es un aviso de rebote, ('baja', [correo]) si pide baja, o None."""
    remitente = parseaddr(msg.get("From", ""))[1].lower()
    texto = _texto_del_mensaje(msg)
    if "mailer-daemon" in remitente or "postmaster" in remitente:
        fallidos = {e.lower() for e in PATRON_EMAIL.findall(texto)} & set(enviados)
        return ("rebote", sorted(fallidos)) if fallidos else None
    if remitente in enviados:
        try:
            asunto = str(make_header(decode_header(msg.get("Subject", ""))))
        except Exception:
            asunto = msg.get("Subject", "")
        propio = "\n".join(l for l in texto.splitlines() if not l.lstrip().startswith(">"))[:600]
        if PATRON_BAJA.search(asunto + " " + propio):
            return ("baja", [remitente])
    return None


def fecha_imap(dt):
    return f"{dt.day:02d}-{MESES_IMAP[dt.month - 1]}-{dt.year}"


# =========================================================
# INTERFAZ
# =========================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Enviador de correos B2B - Gmail + Gemini")
        self.root.geometry("980x860")
        self.root.minsize(880, 720)
        self.root.configure(bg="#1e1e1e")

        self.cola = queue.Queue()
        self.reg = Registro()
        self.enviando = False
        self.smtp = None

        self._construir_ui()
        self._cargar_config()
        self._refrescar_categorias()
        self._refrescar_stats()
        self.root.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.root.after(100, self._vaciar_cola)

    # ---------- UI ----------
    def _construir_ui(self):
        fg, bg, bg2 = "#ffffff", "#1e1e1e", "#2d2d2d"
        self.colores = (fg, bg, bg2)

        estilo = ttk.Style()
        estilo.theme_use("clam")
        estilo.configure("TNotebook", background=bg, borderwidth=0)
        estilo.configure("TNotebook.Tab", background=bg2, foreground="#cccccc", padding=(16, 6))
        estilo.map("TNotebook.Tab", background=[("selected", "#007acc")],
                   foreground=[("selected", "white")])

        head = tk.Frame(self.root, bg=bg2, pady=8, padx=12)
        head.pack(fill=tk.X)
        tk.Label(head, text="✉ Enviador B2B (Gmail + Gemini)", font=("Consolas", 14, "bold"),
                 fg="#00ffcc", bg=bg2).pack(anchor=tk.W)

        nb = ttk.Notebook(self.root)
        nb.pack(fill=tk.X, padx=15, pady=8)
        tab_msg = tk.Frame(nb, bg=bg, padx=10, pady=8)
        tab_cfg = tk.Frame(nb, bg=bg, padx=10, pady=8)
        nb.add(tab_msg, text="✉ Mensaje")
        nb.add(tab_cfg, text="⚙ Cuenta y envío")

        # --- pestaña Mensaje ---
        tab_msg.columnconfigure(1, weight=1)
        self.ent_categoria = self._combo(tab_msg, 0, "Enviar a la categoría:")
        self.ent_categoria.bind("<<ComboboxSelected>>", lambda e: self._refrescar_stats())
        self.var_otros = tk.BooleanVar(value=False)
        tk.Checkbutton(tab_msg, text="Incluir también los 'Otros correos' de cada empresa (no recomendado)",
                       variable=self.var_otros, command=self._refrescar_stats, fg="#cccccc", bg=bg,
                       selectcolor=bg2, activebackground=bg, activeforeground=fg).grid(
            row=1, column=1, sticky=tk.W, padx=8)

        self.ent_remitente = self._entry(tab_msg, 2, "Tu nombre / empresa (remitente):")
        self.ent_idea = self._entry(tab_msg, 3, "Idea para Gemini (qué ofrecés):")
        botones_ia = tk.Frame(tab_msg, bg=bg)
        botones_ia.grid(row=4, column=1, sticky=tk.W, padx=8, pady=2)
        self.btn_borrador = self._boton(botones_ia, "✨ Generar borrador", "#6f42c1",
                                        lambda: self._gemini("borrador"))
        self.btn_mejorar = self._boton(botones_ia, "✨ Mejorar mi texto", "#6f42c1",
                                       lambda: self._gemini("mejora"))

        self.ent_asunto = self._entry(tab_msg, 5, "Asunto:")
        tk.Label(tab_msg, text="Cuerpo  (usá {nombre} y {categoria}):", font=("Segoe UI", 9, "bold"),
                 fg="#cccccc", bg=bg).grid(row=6, column=0, sticky=tk.NW, pady=4)
        self.txt_cuerpo = self._texto(tab_msg, 6, 9)
        self.txt_firma = self._texto(tab_msg, 7, 3, "Firma:")
        self.txt_baja = self._texto(tab_msg, 8, 3, "Texto de baja (obligatorio):")

        # --- pestaña Cuenta y envío ---
        tab_cfg.columnconfigure(1, weight=1)
        self.ent_gmail = self._entry(tab_cfg, 0, "Tu Gmail:")
        self.ent_pass = self._entry(tab_cfg, 1, "Contraseña de aplicación Gmail:", show="*")
        self.ent_pass.insert(0, os.getenv("GMAIL_APP_PASSWORD", ""))
        self.ent_gemini = self._entry(tab_cfg, 2, "API key de Gemini:", show="*")
        self.ent_gemini.insert(0, os.getenv("GEMINI_API_KEY", ""))
        self.ent_modelo = self._entry(tab_cfg, 3, "Modelo de Gemini:")
        self.ent_modelo.insert(0, MODELO_GEMINI_DEFECTO)
        self.ent_limite = self._entry(tab_cfg, 4, "Máximo de correos por día:", ancho=10)
        self.ent_pmin = self._entry(tab_cfg, 5, "Pausa mínima entre correos (seg):", ancho=10)
        self.ent_pmax = self._entry(tab_cfg, 6, "Pausa máxima entre correos (seg):", ancho=10)
        for e, v in ((self.ent_limite, "40"), (self.ent_pmin, "60"), (self.ent_pmax, "150")):
            e.insert(0, v)
        tk.Label(tab_cfg, justify=tk.LEFT, fg="#00ffcc", bg=bg, font=("Segoe UI", 8, "italic"),
                 text=("💡 Cuenta nueva o que nunca envió en volumen: empezá con 20-40 por día y subí de a poco.\n"
                       "   Gmail personal permite hasta ~500/día, pero acercarse a ese tope es lo que hace "
                       "que marquen la cuenta como spam.")).grid(row=7, column=1, sticky=tk.W, padx=8, pady=6)

        # --- botones de acción ---
        barra = tk.Frame(self.root, bg=bg)
        barra.pack(fill=tk.X, padx=15, pady=2)
        self.btn_vista = self._boton(barra, "👁 Vista previa", "#444444", self.vista_previa)
        self.btn_prueba = self._boton(barra, "🧪 Enviarme prueba", "#444444", self.enviar_prueba)
        self.btn_iniciar = self._boton(barra, "▶ Iniciar envío", "#007acc", self.iniciar_envio)
        self.btn_detener = self._boton(barra, "⏹ Detener", "#d9534f", self.detener, tk.DISABLED)
        barra2 = tk.Frame(self.root, bg=bg)
        barra2.pack(fill=tk.X, padx=15, pady=2)
        self._boton(barra2, "📥 Revisar bandeja (bajas y rebotes)", "#444444", self.revisar_bandeja)
        self._boton(barra2, "➕ Agregar baja manual", "#444444", self.baja_manual)
        self._boton(barra2, "📂 Abrir envios.csv", "#444444", self.abrir_envios)
        self.lbl_estado = tk.Label(barra2, text="Listo", font=("Segoe UI", 10, "bold"), fg="#00ffcc", bg=bg)
        self.lbl_estado.pack(side=tk.RIGHT, padx=8)

        self.lbl_stats = tk.Label(self.root, text="", font=("Consolas", 9), fg="#cccccc", bg=bg,
                                  justify=tk.LEFT)
        self.lbl_stats.pack(anchor=tk.W, padx=20, pady=2)

        marco = tk.LabelFrame(self.root, text=" Monitor ", font=("Segoe UI", 10, "bold"),
                              fg=fg, bg=bg, padx=5, pady=5)
        marco.pack(fill=tk.BOTH, expand=True, padx=15, pady=8)
        self.txt_log = scrolledtext.ScrolledText(marco, wrap=tk.WORD, font=("Consolas", 9), height=8,
                                                 bg="#0d1117", fg="#00ff66", insertbackground="white")
        self.txt_log.pack(fill=tk.BOTH, expand=True)

    def _entry(self, parent, fila, etiqueta, show=None, ancho=None):
        fg, bg, bg2 = self.colores
        tk.Label(parent, text=etiqueta, font=("Segoe UI", 9, "bold"), fg="#cccccc", bg=bg).grid(
            row=fila, column=0, sticky=tk.W, pady=4)
        e = tk.Entry(parent, font=("Segoe UI", 10), bg=bg2, fg=fg, insertbackground="white",
                     relief=tk.FLAT, show=show or "", width=ancho or 50)
        e.grid(row=fila, column=1, sticky=tk.W if ancho else tk.EW, padx=8, pady=4)
        return e

    def _combo(self, parent, fila, etiqueta):
        fg, bg, bg2 = self.colores
        tk.Label(parent, text=etiqueta, font=("Segoe UI", 9, "bold"), fg="#cccccc", bg=bg).grid(
            row=fila, column=0, sticky=tk.W, pady=4)
        c = ttk.Combobox(parent, state="readonly", width=30, values=["Todas"])
        c.grid(row=fila, column=1, sticky=tk.W, padx=8, pady=4)
        c.set("Todas")
        return c

    def _texto(self, parent, fila, alto, etiqueta=None):
        fg, bg, bg2 = self.colores
        if etiqueta:
            tk.Label(parent, text=etiqueta, font=("Segoe UI", 9, "bold"), fg="#cccccc", bg=bg).grid(
                row=fila, column=0, sticky=tk.NW, pady=4)
        t = tk.Text(parent, height=alto, font=("Segoe UI", 10), bg=bg2, fg=fg,
                    insertbackground="white", relief=tk.FLAT, wrap=tk.WORD, undo=True)
        t.grid(row=fila, column=1, sticky=tk.EW, padx=8, pady=4)
        return t

    def _boton(self, parent, texto, color, cmd, estado=tk.NORMAL):
        b = tk.Button(parent, text=texto, font=("Segoe UI", 9, "bold"), bg=color, fg="white",
                      activeforeground="white", relief=tk.FLAT, padx=10, pady=6, command=cmd,
                      state=estado)
        b.pack(side=tk.LEFT, padx=3)
        return b

    # ---------- comunicación hilo -> ventana ----------
    def log(self, msg):
        self.cola.put(("log", f"[{datetime.now():%H:%M:%S}] {msg}\n"))

    def _vaciar_cola(self):
        try:
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "log":
                    self.txt_log.insert(tk.END, dato)
                    self.txt_log.see(tk.END)
                elif tipo == "refresh":
                    self._refrescar_stats()
                elif tipo == "estado":
                    self.lbl_estado.config(text=dato[0], fg=dato[1])
                elif tipo == "fin_envio":
                    self.btn_iniciar.config(state=tk.NORMAL)
                    self.btn_detener.config(state=tk.DISABLED)
                    self.lbl_estado.config(text="Listo", fg="#00ffcc")
                    self._refrescar_stats()
                elif tipo == "gemini":
                    asunto, cuerpo = dato
                    if asunto:
                        self.ent_asunto.delete(0, tk.END)
                        self.ent_asunto.insert(0, asunto)
                    self.txt_cuerpo.delete("1.0", tk.END)
                    self.txt_cuerpo.insert("1.0", cuerpo)
                    self.lbl_estado.config(text="Listo", fg="#00ffcc")
                elif tipo == "error_box":
                    messagebox.showerror("Error", dato)
        except queue.Empty:
            pass
        self.root.after(100, self._vaciar_cola)

    # ---------- configuración ----------
    def _cargar_config(self):
        cfg = {}
        if os.path.exists(CONFIG_JSON):
            try:
                with open(CONFIG_JSON, encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                cfg = {}

        def poner(entry, clave, defecto=""):
            valor = cfg.get(clave, defecto)
            if valor != "":
                entry.delete(0, tk.END)
                entry.insert(0, str(valor))

        poner(self.ent_gmail, "gmail")
        poner(self.ent_remitente, "remitente")
        poner(self.ent_idea, "idea")
        poner(self.ent_asunto, "asunto")
        poner(self.ent_modelo, "modelo", MODELO_GEMINI_DEFECTO)
        poner(self.ent_limite, "limite", "40")
        poner(self.ent_pmin, "pmin", "60")
        poner(self.ent_pmax, "pmax", "150")
        self.txt_cuerpo.insert("1.0", cfg.get("cuerpo", CUERPO_EJEMPLO))
        self.txt_firma.insert("1.0", cfg.get("firma", FIRMA_EJEMPLO))
        self.txt_baja.insert("1.0", cfg.get("baja", BAJA_EJEMPLO))
        if cfg.get("categoria"):
            self.ent_categoria.set(cfg["categoria"])

    def _guardar_config(self):
        cfg = {
            "gmail": self.ent_gmail.get().strip(), "remitente": self.ent_remitente.get().strip(),
            "idea": self.ent_idea.get().strip(), "asunto": self.ent_asunto.get().strip(),
            "modelo": self.ent_modelo.get().strip(), "limite": self.ent_limite.get().strip(),
            "pmin": self.ent_pmin.get().strip(), "pmax": self.ent_pmax.get().strip(),
            "cuerpo": self.txt_cuerpo.get("1.0", tk.END).strip(),
            "firma": self.txt_firma.get("1.0", tk.END).strip(),
            "baja": self.txt_baja.get("1.0", tk.END).strip(),
            "categoria": self.ent_categoria.get(),
        }
        try:
            with open(CONFIG_JSON, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"⚠️ No pude guardar el borrador: {e}")

    def _leer_cfg(self, requiere_cuenta=True):
        """Lee y valida los campos. Devuelve dict o None (mostrando el error)."""
        cfg = {
            "gmail": self.ent_gmail.get().strip(),
            "password": self.ent_pass.get().replace(" ", ""),
            "remitente": self.ent_remitente.get().strip(),
            "asunto": self.ent_asunto.get().strip(),
            "cuerpo": self.txt_cuerpo.get("1.0", tk.END).strip(),
            "firma": self.txt_firma.get("1.0", tk.END).strip(),
            "baja": self.txt_baja.get("1.0", tk.END).strip(),
        }

        def error(msg):
            messagebox.showerror("Revisá los datos", msg)
            return None

        if requiere_cuenta:
            if "@" not in cfg["gmail"]:
                return error("Escribí tu Gmail en la pestaña 'Cuenta y envío'.")
            if not cfg["password"]:
                return error("Falta la contraseña de aplicación de Gmail (pestaña 'Cuenta y envío').")
        if not cfg["remitente"]:
            return error("Escribí tu nombre o el de tu empresa (remitente).")
        if not cfg["asunto"] or not cfg["cuerpo"]:
            return error("Completá el asunto y el cuerpo del mensaje.")
        if not cfg["baja"]:
            return error("El texto de baja es obligatorio: tiene que haber una forma clara de pedir que no les escribas más.")
        if re.search(r"\[[^\]]+\]", cfg["asunto"] + cfg["cuerpo"] + cfg["firma"]):
            if not messagebox.askyesno("Texto sin completar",
                                       "Todavía hay textos entre [corchetes] (ej.: [tu empresa]).\n"
                                       "¿Querés enviar igual?"):
                return None
        try:
            cfg["limite"] = int(self.ent_limite.get().strip())
            cfg["pmin"] = float(self.ent_pmin.get().strip().replace(",", "."))
            cfg["pmax"] = float(self.ent_pmax.get().strip().replace(",", "."))
        except ValueError:
            return error("Límite diario y pausas deben ser números.")
        if not 1 <= cfg["limite"] <= 500:
            return error("El máximo por día debe estar entre 1 y 500.")
        if cfg["pmin"] < 20:
            return error("La pausa mínima no puede ser menor a 20 segundos (enviar más rápido dispara filtros).")
        if cfg["pmax"] < cfg["pmin"]:
            return error("La pausa máxima tiene que ser mayor o igual que la mínima.")
        cfg["categoria"] = self.ent_categoria.get()
        cfg["otros"] = self.var_otros.get()
        return cfg

    # ---------- datos ----------
    def _refrescar_categorias(self):
        cats = sorted({f["categoria"] for f in leer_prospectos() if f["categoria"]}, key=str.casefold)
        actual = self.ent_categoria.get()
        self.ent_categoria.config(values=["Todas"] + cats)
        self.ent_categoria.set(actual if actual in cats else "Todas")

    def _pendientes(self, categoria, incluir_otros):
        res, vistos = [], set()
        for f in leer_prospectos():
            if categoria not in ("", "Todas") and f["categoria"].casefold() != categoria.casefold():
                continue
            correos = f["correos"] if incluir_otros else f["correos"][:1]
            for c in correos:
                if not PATRON_EMAIL.fullmatch(c) or c in vistos:
                    continue
                vistos.add(c)
                if c in self.reg.hechos or c in self.reg.bajas:
                    continue
                res.append({"correo": c, "nombre": f["nombre"], "categoria": f["categoria"]})
        return res

    def _refrescar_stats(self):
        try:
            limite = int(self.ent_limite.get().strip())
        except ValueError:
            limite = 0
        pend = self._pendientes(self.ent_categoria.get(), self.var_otros.get())
        self.lbl_stats.config(text=(
            f"Pendientes (categoría elegida): {len(pend)}   |   Enviados hoy: {self.reg.hoy}/{limite}   |   "
            f"Enviados en total: {self.reg.total}   |   Bajas: {len(self.reg.bajas)}   |   "
            f"Rebotados: {len(self.reg.rebotados)}"))

    # ---------- acciones ----------
    def vista_previa(self):
        cfg = self._leer_cfg(requiere_cuenta=False)
        if not cfg:
            return
        pend = self._pendientes(cfg["categoria"], cfg["otros"])
        nombre, cat, dest = ("Empresa de Ejemplo S.A.", "Industria", "contacto@ejemplo.com")
        if pend:
            nombre, cat, dest = pend[0]["nombre"], pend[0]["categoria"], pend[0]["correo"]
        cfg.setdefault("gmail", "tucorreo@gmail.com")
        cfg["gmail"] = cfg["gmail"] or "tucorreo@gmail.com"
        msg = construir_mensaje(cfg, dest, nombre, cat)
        ventana = tk.Toplevel(self.root)
        ventana.title("Vista previa")
        ventana.geometry("640x460")
        caja = scrolledtext.ScrolledText(ventana, wrap=tk.WORD, font=("Segoe UI", 10))
        caja.pack(fill=tk.BOTH, expand=True)
        caja.insert("1.0", f"De: {msg['From']}\nPara: {msg['To']}\nAsunto: {msg['Subject']}\n"
                           f"{'-' * 60}\n{msg.get_content()}")
        caja.config(state=tk.DISABLED)

    def enviar_prueba(self):
        cfg = self._leer_cfg()
        if cfg:
            self._guardar_config()
            threading.Thread(target=self._hilo_prueba, args=(cfg,), daemon=True).start()

    def _hilo_prueba(self, cfg):
        self.log("🧪 Enviando prueba a tu propio Gmail...")
        try:
            smtp = self._abrir_smtp(cfg)
            smtp.send_message(construir_mensaje(cfg, cfg["gmail"], "Empresa de Ejemplo S.A.",
                                                "Industria", prueba=True))
            smtp.quit()
            self.log("✅ Prueba enviada. Revisá tu bandeja (y la carpeta Spam) para ver cómo llega.")
        except smtplib.SMTPAuthenticationError:
            self.log("❌ Gmail rechazó el usuario/contraseña. Tiene que ser una CONTRASEÑA DE APLICACIÓN.")
        except Exception as e:
            self.log(f"❌ No pude enviar la prueba: {e}")

    def iniciar_envio(self):
        cfg = self._leer_cfg()
        if not cfg:
            return
        self.reg.recargar()
        pend = self._pendientes(cfg["categoria"], cfg["otros"])
        restante = cfg["limite"] - self.reg.hoy
        if not pend:
            messagebox.showinfo("Nada para enviar", "No quedan destinatarios pendientes en esa categoría.")
            return
        if restante <= 0:
            messagebox.showinfo("Límite diario", f"Ya enviaste {self.reg.hoy} hoy (límite {cfg['limite']}). "
                                                 "Seguí mañana o subí el límite con cuidado.")
            return
        n = min(restante, len(pend))
        minutos = n * (cfg["pmin"] + cfg["pmax"]) / 2 / 60
        if not messagebox.askyesno(
                "Confirmar envío",
                f"Se van a enviar {n} correos desde {cfg['gmail']}\n"
                f"(categoría: {cfg['categoria']}), con pausas de {cfg['pmin']:.0f}-{cfg['pmax']:.0f} s.\n"
                f"Duración estimada: ~{minutos:.0f} minutos. Dejá la ventana abierta.\n\n"
                "Una vez enviado no se puede deshacer. ¿Continuar?"):
            return
        self._guardar_config()
        self.enviando = True
        self.btn_iniciar.config(state=tk.DISABLED)
        self.btn_detener.config(state=tk.NORMAL)
        self.lbl_estado.config(text="Enviando...", fg="#ffcc00")
        threading.Thread(target=self._hilo_envio, args=(cfg, pend[:n]), daemon=True).start()

    def detener(self):
        self.enviando = False
        self.lbl_estado.config(text="Deteniendo...", fg="#ff4444")
        self.log("🛑 Detención pedida: termina el correo en curso y para.")

    # ---------- envío ----------
    def _abrir_smtp(self, cfg):
        smtp = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ssl.create_default_context(), timeout=30)
        smtp.login(cfg["gmail"], cfg["password"])
        return smtp

    def _cerrar_smtp(self):
        try:
            if self.smtp:
                self.smtp.quit()
        except Exception:
            pass
        self.smtp = None

    def _smtp_enviar(self, cfg, msg):
        for intento in (1, 2):
            try:
                if self.smtp is None:
                    self.smtp = self._abrir_smtp(cfg)
                self.smtp.send_message(msg)
                return
            except smtplib.SMTPServerDisconnected:
                self.smtp = None
                if intento == 2:
                    raise
            except smtplib.SMTPException:
                raise
            except (ConnectionError, TimeoutError, ssl.SSLError):
                self._cerrar_smtp()
                if intento == 2:
                    raise

    def _esperar(self, segundos):
        fin = time.time() + segundos
        while time.time() < fin and self.enviando:
            time.sleep(0.5)

    def _hilo_envio(self, cfg, destinos):
        enviados = errores_seguidos = 0
        try:
            self.log(f"📤 Iniciando envío de {len(destinos)} correo(s)...")
            for i, d in enumerate(destinos, 1):
                if not self.enviando:
                    break
                if d["correo"] in self.reg.hechos or d["correo"] in self.reg.bajas:
                    continue
                asunto_real = renderizar(cfg["asunto"], d["nombre"], d["categoria"]).strip()
                try:
                    self._smtp_enviar(cfg, construir_mensaje(cfg, d["correo"], d["nombre"], d["categoria"]))
                    ok = self.reg.anotar(d["correo"], d["nombre"], d["categoria"], asunto_real, "enviado")
                    enviados += 1
                    errores_seguidos = 0
                    self.log(f"✅ [{i}/{len(destinos)}] {d['correo']}  ({d['nombre']})")
                    if not ok:
                        self.log("⚠️ envios.csv está abierto en Excel: cerralo. Los envíos quedan en memoria "
                                 "y se anotan en la próxima ronda (no se repetirán).")
                except smtplib.SMTPAuthenticationError:
                    self.log("❌ Gmail rechazó el usuario/contraseña de aplicación. Se detiene el envío.")
                    break
                except smtplib.SMTPRecipientsRefused:
                    self.reg.anotar(d["correo"], d["nombre"], d["categoria"], asunto_real, "rebotado",
                                    "Dirección rechazada")
                    self.log(f"↩️ [{i}/{len(destinos)}] {d['correo']}: dirección inexistente, marcada como rebotada.")
                except Exception as e:
                    detalle = str(e)[:200]
                    self.reg.anotar(d["correo"], d["nombre"], d["categoria"], asunto_real, "error", detalle)
                    errores_seguidos += 1
                    self.log(f"⚠️ [{i}/{len(destinos)}] {d['correo']}: {detalle}")
                    if "5.4.5" in detalle or "daily" in detalle.lower():
                        self.log("🛑 Gmail indica que llegaste al límite diario. Se detiene el envío.")
                        break
                    if errores_seguidos >= 3:
                        self.log("🛑 3 errores seguidos: se detiene para proteger tu cuenta. Revisá y reintentá.")
                        break
                self.cola.put(("refresh", None))
                if i < len(destinos) and self.enviando:
                    pausa = random.uniform(cfg["pmin"], cfg["pmax"])
                    self.log(f"   ⏳ Pausa de {pausa:.0f} s...")
                    self._esperar(pausa)
        except Exception as e:
            self.log(f"❌ Error crítico: {e}")
        finally:
            self._cerrar_smtp()
            alterno = self.reg.cerrar()
            if alterno:
                self.log(f"⚠️ envios.csv seguía abierto. Registro guardado en {os.path.basename(alterno)}")
            self.log(f"\n🏁 Fin. Enviados en esta tanda: {enviados}. "
                     "Tip: en unas horas usá '📥 Revisar bandeja' para detectar rebotes y bajas.")
            self.enviando = False
            self.cola.put(("fin_envio", None))

    # ---------- bandeja (rebotes y bajas) ----------
    def revisar_bandeja(self):
        cfg = self._leer_cfg()
        if cfg:
            threading.Thread(target=self._hilo_bandeja, args=(cfg,), daemon=True).start()

    def _hilo_bandeja(self, cfg):
        self.log("📥 Revisando tu bandeja (últimos 14 días)...")
        try:
            imap = imaplib.IMAP4_SSL(IMAP_HOST)
            imap.login(cfg["gmail"], cfg["password"])
            imap.select("INBOX", readonly=True)
            _, datos = imap.search(None, "SINCE", fecha_imap(datetime.now() - timedelta(days=14)))
            ids = datos[0].split()[-300:]
            enviados = set(self.reg.enviados)
            rebotes = bajas = 0
            for mid in ids:
                tipo, resp = imap.fetch(mid, "(BODY.PEEK[])")
                if tipo != "OK" or not resp or not isinstance(resp[0], tuple):
                    continue
                res = analizar_mensaje(email.message_from_bytes(resp[0][1]), enviados)
                if not res:
                    continue
                if res[0] == "rebote":
                    for c in res[1]:
                        if c not in self.reg.rebotados:
                            self.reg.anotar(c, "", "", "", "rebotado", "Aviso de rebote en bandeja")
                            rebotes += 1
                            self.log(f"↩️ Rebote detectado: {c}")
                else:
                    for c in res[1]:
                        if c not in self.reg.bajas:
                            self.reg.agregar_baja(c)
                            self.reg.anotar(c, "", "", "", "baja", "Pidió la baja por respuesta")
                            bajas += 1
                            self.log(f"🚫 Baja detectada: {c}")
            imap.logout()
            alterno = self.reg.cerrar()
            if alterno:
                self.log(f"⚠️ envios.csv abierto: registro en {os.path.basename(alterno)}")
            self.log(f"✔ Bandeja revisada: {rebotes} rebote(s) y {bajas} baja(s) nuevos. "
                     "Si alguien pide la baja con otro texto o desde otra dirección, usá '➕ Agregar baja manual'.")
            self.cola.put(("refresh", None))
        except imaplib.IMAP4.error as e:
            self.log(f"❌ Error de acceso a la bandeja: {e}")
        except Exception as e:
            self.log(f"❌ No pude revisar la bandeja: {e}")

    def baja_manual(self):
        correo = simpledialog.askstring("Agregar baja", "Correo que no debe recibir más mensajes:")
        if not correo:
            return
        correo = correo.strip().lower()
        if not PATRON_EMAIL.fullmatch(correo):
            messagebox.showerror("Correo inválido", "Ese correo no parece válido.")
            return
        self.reg.agregar_baja(correo)
        self.reg.anotar(correo, "", "", "", "baja", "Agregada manualmente")
        self.reg.cerrar()
        self.log(f"🚫 Baja agregada: {correo}")
        self._refrescar_stats()

    # ---------- Gemini ----------
    def _gemini(self, modo):
        clave = self.ent_gemini.get().strip()
        modelo = self.ent_modelo.get().strip() or MODELO_GEMINI_DEFECTO
        if not clave:
            messagebox.showerror("Falta la API key", "Pegá tu API key de Gemini en 'Cuenta y envío'.")
            return
        if modo == "borrador":
            idea = self.ent_idea.get().strip()
            if not idea:
                messagebox.showerror("Falta la idea", "Contale a Gemini quién sos y qué ofrecés (campo 'Idea').")
                return
            prompt = prompt_borrador(idea)
        else:
            cuerpo = self.txt_cuerpo.get("1.0", tk.END).strip()
            if not cuerpo:
                messagebox.showerror("Sin texto", "Escribí primero un cuerpo para mejorar.")
                return
            prompt = prompt_mejora(self.ent_asunto.get().strip(), cuerpo)
        self.lbl_estado.config(text="Gemini escribiendo...", fg="#ffcc00")

        def tarea():
            try:
                asunto, cuerpo = parsear_gemini(llamar_gemini(clave, modelo, prompt))
                self.cola.put(("gemini", (asunto, cuerpo)))
                self.log("✨ Gemini terminó. Revisá y editá el texto antes de enviar.")
            except Exception as e:
                self.cola.put(("estado", ("Listo", "#00ffcc")))
                self.cola.put(("error_box", f"No pude usar Gemini:\n{e}"))

        threading.Thread(target=tarea, daemon=True).start()

    # ---------- varios ----------
    def abrir_envios(self):
        if not os.path.exists(ENVIOS_CSV):
            messagebox.showinfo("Sin registro", "Todavía no se envió ningún correo.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(ENVIOS_CSV)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", ENVIOS_CSV])
            else:
                subprocess.Popen(["xdg-open", ENVIOS_CSV])
        except Exception as e:
            messagebox.showerror("Error", f"No pude abrir el archivo:\n{e}")

    def _cerrar(self):
        self.enviando = False
        self._guardar_config()
        self.root.destroy()


if __name__ == "__main__":
    ventana = tk.Tk()
    App(ventana)
    ventana.mainloop()