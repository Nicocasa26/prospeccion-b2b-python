# 🚀 Prospección B2B + Enviador Inteligente de Correos

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?style=for-the-badge&logo=python&logoColor=white)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-orange?style=for-the-badge)
![Gemini API](https://img.shields.io/badge/AI-Gemini%20API-4285F4?style=for-the-badge&logo=google&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

Sistema de automatización de escritorio desarrollado en Python para la **localización, extracción y gestión responsable de prospectos comerciales B2B** en GBA y CABA (Argentina). Integra scraping concurrente determinista y un enviador con gestión de reputación de dominio asistido por la **API de Gemini**.

---

## 📌 Características Principales

El proyecto consta de **dos aplicaciones de escritorio independientes**:

### 1. 🔍 Buscador y Extractor de Prospectos (`buscador.py`)
- **Búsqueda por Tags × Zonas:** Parametrización automática para recorrer 65+ localidades del Gran Buenos Aires y CABA.
- **Scraping Concurrente:** Procesamiento multihilo (`ThreadPoolExecutor` con 6 hilos) de sitios web corporativos y sus páginas internas de contacto.
- **Extracción Determinista:** Parsing estricto de HTML (regex + Cloudflare email decoding) para obtener correos reales descartando falsos positivos (`noreply`, recursos gráficos, CDN, dominios externos).
- **Deduplicación e Historial:** Registro persistente de búsquedas realizadas y dominios inspeccionados para evitar trabajo redundante.

### 2. ✉️ Enviador Inteligente y Control de Reputación (`enviador.py`)
- **Asistente de Redacción IA:** Integración con la API de Gemini (`gemini-2.5-flash`) para redactar borradores y optimizar copys comerciales en español rioplatense.
- **Control de Entregabilidad (Anti-Spam):** Pausas estocásticas (aleatorias entre envíos), límite diario estricto, formato en texto plano y encabezado estándar `List-Unsubscribe` de baja obligatoria.
- **Gestión Automatizada de Bajas y Bounces (IMAP):** Lectura periódica de la bandeja de entrada para detectar automáticamente *Mailer-Daemon* (rebotes) y solicitudes de baja (*opt-out*).
- **Persistencia A prueba de Bloqueos:** Manejo de concurrencia para archivos de salida CSV en caso de que estén abiertos en Excel.

---

## 🛠️ Stack Tecnológico

- **Lenguaje:** Python 3.x
- **Interfaz Gráfica:** `tkinter` + `ttk` (Multithreading + `queue.Queue` para una UI responsiva)
- **Scraping & Conexiones:** `requests`, `duckduckgo_search` (`ddgs`), `urllib.parse`
- **Email & Protocolos:** `smtplib` (SSL/TLS), `imaplib`, `email.message`
- **IA & NLP:** Gemini API (`generativelanguage.googleapis.com`)
- **Empaquetado:** `PyInstaller`

---

## ⚙️ Instalación y Configuración

### Requisitos Previos
- Python 3.8 o superior instalado.
- Cuenta de Gmail con **Verificación en 2 pasos** y una **Contraseña de aplicación** de 16 caracteres.
- *(Opcional)* API Key de [Google AI Studio](https://aistudio.google.com/apikey) para la redacción con Gemini.

### Pasos de Instalación
```bash
# 1. Clonar el repositorio
git clone [https://github.com/tu-usuario/prospeccion-b2b-python.git](https://github.com/tu-usuario/prospeccion-b2b-python.git)
cd prospeccion-b2b-python

# 2. Crear y activar un entorno virtual (Recomendado)
python -m venv venv
# En Windows:
venv\Scripts\activate
# En Linux/Mac:
source venv/bin/activate

# 3. Instalar dependencias
pip install -r requirements.txt
