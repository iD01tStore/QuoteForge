# QuoteForgeAI.py
#
# Brand: iD01t Productions
# App Name: QuoteForgeAI Pro
# Version: 1.0.0
# License: $299 CAD, Lifetime License
#
# A premium desktop application for construction professionals in Canada.
# Features include AI-powered soumission reading, quote building with a
# Canadian tax engine, contract generation, client management, and more.
# All contained within a single, production-grade Python script.

import sys
import subprocess
import os
import json
import sqlite3
import threading
import queue
import logging
from logging.handlers import RotatingFileHandler
import base64
import datetime
import re
from tkinter import (
    Tk,
    Frame,
    Label,
    Button,
    Entry,
    Text,
    Scrollbar,
    StringVar,
    IntVar,
    DoubleVar,
    BooleanVar,
    Toplevel,
    PhotoImage,
    Menu,
    messagebox,
    filedialog,
    simpledialog,
)

# --- DEPENDENCY INSTALLATION ---

# List of required packages
REQUIRED_PACKAGES = [
    ("ttkbootstrap", "ttkbootstrap"),
    ("reportlab", "reportlab"),
    ("requests", "requests"),
    ("pypdf", "pypdf"),
    ("pdfminer.six", "pdfminer.six"),
    ("docx", "python-docx"),
    ("PIL", "Pillow"),
    ("cryptography", "cryptography"),
    ("chardet", "chardet"),
]

# List of optional packages for enhanced features
OPTIONAL_PACKAGES = [
    ("pytesseract", "pytesseract"),
    ("pdf2image", "pdf2image"),
    ("google.generativeai", "google-generativeai"),
]

# Track installed status to avoid re-checking
_deps_checked = False
_optional_deps = {}


def ensure_deps():
    """
    Checks for required Python packages and installs them if missing.
    Prints progress to the console.
    """
    global _deps_checked, _optional_deps
    if _deps_checked:
        return True

    print("QuoteForgeAI Pro: Initializing...")
    print("Checking for required dependencies...")

    all_packages = REQUIRED_PACKAGES + OPTIONAL_PACKAGES
    missing_packages = []

    for import_name, package_name in all_packages:
        try:
            __import__(import_name)
            if import_name in [p[0] for p in OPTIONAL_PACKAGES]:
                _optional_deps[import_name] = True
        except ImportError:
            is_required = import_name in [p[0] for p in REQUIRED_PACKAGES]
            if is_required:
                missing_packages.append(package_name)
            _optional_deps[import_name] = False

    if missing_packages:
        print(f"Missing required packages: {', '.join(missing_packages)}. Attempting to install...")
        total_missing = len(missing_packages)
        for i, pkg in enumerate(missing_packages):
            print(f"Installing dependency {i+1}/{total_missing}: {pkg}...")
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", pkg],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.PIPE,
                )
            except subprocess.CalledProcessError as e:
                print(f"FATAL: Failed to install required package: {pkg}.")
                print(f"Please install it manually using: pip install {pkg}")
                error_output = e.stderr.decode() if e.stderr else str(e)
                print(f"Error: {error_output}")
                sys.exit(1)
        print("All required dependencies installed successfully.")
    else:
        print("All required dependencies are already installed.")

    _deps_checked = True
    print("Initialization complete.")
    return True

# Run dependency check before importing external modules
ensure_deps()


# --- Now that deps are handled, we can import them ---
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.scrolled import ScrolledText, ScrolledFrame
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.colors import navy, black, gray, white, HexColor
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
import requests
from pypdf import PdfReader
from pdfminer.high_level import extract_text as extract_text_from_pdf
from docx import Document
from PIL import Image as PILImage, ImageTk
from cryptography.fernet import Fernet, InvalidToken
import chardet
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Conditional imports for optional features
try:
    import pytesseract
    from pdf2image import convert_from_path
except ImportError:
    pytesseract = None
    convert_from_path = None

try:
    import google.generativeai as genai
except ImportError:
    genai = None


# --- CONSTANTS AND CONFIGURATION ---

APP_NAME = "QuoteForgeAI Pro"
APP_VERSION = "1.0.0"
BRAND_NAME = "iD01t Productions"
DB_FILE = "quoteforge_data.db"
CONFIG_FILE = "config.json"
KEY_FILE = "key.bin"
LOG_FILE = "quoteforge.log"

PROVINCES = {
    "AB": "Alberta",
    "BC": "British Columbia",
    "MB": "Manitoba",
    "NB": "New Brunswick",
    "NL": "Newfoundland and Labrador",
    "NS": "Nova Scotia",
    "NT": "Northwest Territories",
    "NU": "Nunavut",
    "ON": "Ontario",
    "PE": "Prince Edward Island",
    "QC": "Quebec",
    "SK": "Saskatchewan",
    "YT": "Yukon",
}

# Tax rates as of a recent period, subject to change
# Structure: { province_code: { tax_name: rate, ... } }
TAX_STRUCTURE = {
    "AB": {"GST": 0.05},
    "BC": {"GST": 0.05, "PST": 0.07},
    "MB": {"GST": 0.05, "PST": 0.07},
    "NB": {"HST": 0.15},
    "NL": {"HST": 0.15},
    "NS": {"HST": 0.15},
    "NT": {"GST": 0.05},
    "NU": {"GST": 0.05},
    "ON": {"HST": 0.13},
    "PE": {"HST": 0.15},
    "QC": {"GST": 0.05, "QST": 0.09975},
    "SK": {"GST": 0.05, "PST": 0.06},
    "YT": {"GST": 0.05},
}

# Default configuration settings
DEFAULT_CONFIG = {
    "company_info": {
        "name": "Your Company Name",
        "address": "123 Main Street, Anytown, Province, A1B 2C3",
        "phone": "(555) 123-4567",
        "email": "contact@yourcompany.com",
        "logo_path": "",
    },
    "defaults": {
        "province": "ON",
        "currency": "CAD",
        "language": "en",
        "holdback_percent": 10.0,
    },
    "ai_keys": {
        "gemini_api_key": "",
        "openai_api_key": "",
        "anthropic_api_key": "",
        "ollama_base_url": "http://localhost:11434",
        "ollama_model": "llama3",
        "cohere_api_key": "",
        "mistral_api_key": "",
        "azure_openai_api_key": "",
        "azure_openai_endpoint": "",
        "aws_access_key_id": "",
        "aws_secret_access_key": "",
        "aws_region_name": "",
        "groq_api_key": "",
        "nvidia_api_key": "",
    },
    "smtp": {
        "server": "smtp.example.com",
        "port": 587,
        "username": "your_email@example.com",
        "password": "",
    },
    "license": {"key": ""},
    "theme": "darkly",
    "use_encryption": False,
}

# --- I18N LOCALIZATION ---

I18N_STRINGS = {
    "en": {
        "app_title": f"{APP_NAME} by {BRAND_NAME}",
        "dashboard": "Dashboard",
        "soumission_reader": "Soumission Reader",
        "quote_builder": "Quote Builder",
        "contract_builder": "Contract Builder",
        "client_manager": "Client Manager",
        "templates": "Templates",
        "settings": "Settings",
        "logs": "Logs",
        "help": "Help",
        "file": "File",
        "new_quote": "New Quote",
        "new_contract": "New Contract",
        "exit": "Exit",
        "tools": "Tools",
        "bulk_parse": "Bulk Parse Soumissions",
        "welcome": f"Welcome to {APP_NAME}",
        "quick_actions": "Quick Actions",
        "add_client": "Add Client",
        "import_soumission": "Import Soumission",
        "recent_projects": "Recent Projects",
        "project_name": "Project Name",
        "client": "Client",
        "amount": "Amount",
        "status": "Status",
        "created_date": "Created Date",
        "load_file": "Load File (PDF, DOCX, TXT)",
        "parsing_in_progress": "Parsing document, please wait...",
        "ai_extraction_in_progress": "AI is extracting data, this may take a moment...",
        "extraction_review": "Extraction Review & Approval",
        "field": "Field",
        "extracted_value": "Extracted Value",
        "confidence": "Confidence",
        "accept": "Accept",
        "edit": "Edit",
        "reject": "Reject",
        "push_to_quote_builder": "Push Accepted to Quote Builder",
        "line_items_for_quote": "Line Items for Quote",
        "description": "Description",
        "quantity": "Quantity",
        "unit": "Unit",
        "rate": "Rate",
        "line_total": "Line Total",
        "accept_all": "Accept All",
        "reject_all": "Reject All",
        "save": "Save",
        "export_pdf": "Export PDF",
        "send_email": "Send via Email",
        "item_description": "Item Description",
        "add_item": "Add Item",
        "delete_item": "Delete Item",
        "subtotal": "Subtotal",
        "taxes": "Taxes",
        "grand_total": "Grand Total",
        "construction_holdback": "Construction Holdback ({percent}%)",
        "total_with_holdback": "Total with Holdback",
        "total_without_holdback": "Total without Holdback",
        "province": "Province",
        "override_rates": "Override Tax Rates",
        # ... more strings
    },
    "fr": {
        "app_title": f"{APP_NAME} par {BRAND_NAME}",
        "dashboard": "Tableau de bord",
        "soumission_reader": "Lecteur de soumission",
        "quote_builder": "Créateur de devis",
        "contract_builder": "Créateur de contrat",
        "client_manager": "Gestion des clients",
        "templates": "Modèles",
        "settings": "Paramètres",
        "logs": "Journaux",
        "help": "Aide",
        "file": "Fichier",
        "new_quote": "Nouveau devis",
        "new_contract": "Nouveau contrat",
        "exit": "Quitter",
        "tools": "Outils",
        "bulk_parse": "Analyser les soumissions en masse",
        "welcome": f"Bienvenue à {APP_NAME}",
        "quick_actions": "Actions rapides",
        "add_client": "Ajouter un client",
        "import_soumission": "Importer une soumission",
        "recent_projects": "Projets récents",
        "project_name": "Nom du projet",
        "client": "Client",
        "amount": "Montant",
        "status": "Statut",
        "created_date": "Date de création",
        "load_file": "Charger un fichier (PDF, DOCX, TXT)",
        "parsing_in_progress": "Analyse du document en cours, veuillez patienter...",
        "ai_extraction_in_progress": "L'IA extrait les données, cela peut prendre un moment...",
        "extraction_review": "Révision et approbation de l'extraction",
        "field": "Champ",
        "extracted_value": "Valeur extraite",
        "confidence": "Confiance",
        "accept": "Accepter",
        "edit": "Modifier",
        "reject": "Rejeter",
        "push_to_quote_builder": "Pousser vers le créateur de devis",
        "line_items_for_quote": "Lignes d'articles pour le devis",
        "description": "Description",
        "quantity": "Quantité",
        "unit": "Unité",
        "rate": "Taux",
        "line_total": "Total ligne",
        "accept_all": "Tout accepter",
        "reject_all": "Tout rejeter",
        "save": "Enregistrer",
        "export_pdf": "Exporter en PDF",
        "send_email": "Envoyer par courriel",
        "item_description": "Description de l'article",
        "add_item": "Ajouter un article",
        "delete_item": "Supprimer l'article",
        "subtotal": "Sous-total",
        "taxes": "Taxes",
        "grand_total": "Grand Total",
        "construction_holdback": "Retenue de construction ({percent}%)",
        "total_with_holdback": "Total avec retenue",
        "total_without_holdback": "Total sans retenue",
        "province": "Province",
        "override_rates": "Modifier les taux de taxe",
        # ... more strings
    },
}
current_lang = "en"


def i18n(key, **kwargs):
    """Simple i18n lookup function."""
    return I18N_STRINGS.get(current_lang, I18N_STRINGS["en"]).get(key, key).format(**kwargs)


# --- LOGGING SETUP ---

log_formatter = logging.Formatter(
    "%(asctime)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
)
log_handler = RotatingFileHandler(
    LOG_FILE, maxBytes=1 * 1024 * 1024, backupCount=5
)
log_handler.setFormatter(log_formatter)
logger = logging.getLogger("QuoteForgeAI")
logger.setLevel(logging.INFO)
logger.addHandler(log_handler)


class QueueHandler(logging.Handler):
    """Class to send logging records to a queue."""
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))


# --- CORE APPLICATION LOGIC ---


class ConfigManager:
    """Handles loading, saving, and managing application configuration."""

    def __init__(self):
        self.config = {}
        self.fernet = None
        self.load_config()

    def _load_key(self):
        """Loads the encryption key or generates a new one."""
        try:
            with open(KEY_FILE, "rb") as f:
                key = f.read()
            self.fernet = Fernet(key)
        except FileNotFoundError:
            key = Fernet.generate_key()
            with open(KEY_FILE, "wb") as f:
                f.write(key)
            self.fernet = Fernet(key)
            logger.info("New encryption key generated.")

    def _encrypt(self, data):
        if not self.fernet or not data:
            return data
        return self.fernet.encrypt(data.encode()).decode()

    def _decrypt(self, data):
        if not self.fernet or not data:
            return data
        try:
            return self.fernet.decrypt(data.encode()).decode()
        except (InvalidToken, TypeError):
            logger.warning(f"Could not decrypt value, returning as is.")
            return data

    def load_config(self):
        """Loads config from file and overrides with environment variables."""
        self.config = json.loads(json.dumps(DEFAULT_CONFIG)) # Deep copy
        try:
            with open(CONFIG_FILE, "r") as f:
                file_config = json.load(f)
                # Deep merge
                for key, value in file_config.items():
                    if isinstance(value, dict) and key in self.config:
                        self.config[key].update(value)
                    else:
                        self.config[key] = value
            logger.info("Configuration loaded from config.json")
        except (FileNotFoundError, json.JSONDecodeError):
            logger.info("config.json not found or invalid, using defaults.")
            self.save_config()  # Create a default one

        if self.config.get("use_encryption", False):
            self._load_key()
            self.config["smtp"]["password"] = self._decrypt(self.config["smtp"]["password"])
            for key in self.config["ai_keys"]:
                if "key" in key and self.config["ai_keys"][key]:
                    self.config["ai_keys"][key] = self._decrypt(self.config["ai_keys"][key])

        # Environment variable overrides
        self.config["ai_keys"]["gemini_api_key"] = os.environ.get(
            "GEMINI_API_KEY", self.config["ai_keys"]["gemini_api_key"]
        )
        self.config["ai_keys"]["openai_api_key"] = os.environ.get(
            "OPENAI_API_KEY", self.config["ai_keys"]["openai_api_key"]
        )
        self.config["ai_keys"]["anthropic_api_key"] = os.environ.get(
            "ANTHROPIC_API_KEY", self.config["ai_keys"]["anthropic_api_key"]
        )
        # ... other env vars

    def save_config(self):
        """Saves the current configuration to config.json."""
        try:
            config_to_save = json.loads(json.dumps(self.config)) # Deep copy
            if config_to_save.get("use_encryption", False) and self.fernet:
                 config_to_save["smtp"]["password"] = self._encrypt(config_to_save["smtp"]["password"])
                 for key in config_to_save["ai_keys"]:
                    if "key" in key and config_to_save["ai_keys"][key]:
                        config_to_save["ai_keys"][key] = self._encrypt(config_to_save["ai_keys"][key])

            with open(CONFIG_FILE, "w") as f:
                json.dump(config_to_save, f, indent=4)
            logger.info("Configuration saved to config.json")
        except Exception as e:
            logger.error(f"Error saving configuration: {e}")
            messagebox.showerror("Config Error", f"Could not save configuration file.\n{e}")

    def get(self, key_path, default=None):
        """Gets a value from config using dot notation."""
        keys = key_path.split(".")
        val = self.config
        for key in keys:
            val = val.get(key)
            if val is None:
                return default
        return val

    def set(self, key_path, value):
        """Sets a value in config using dot notation."""
        keys = key_path.split(".")
        d = self.config
        for key in keys[:-1]:
            d = d.setdefault(key, {})
        d[keys[-1]] = value


class DBManager:
    """Handles all SQLite database operations."""

    def __init__(self, db_file):
        self.db_file = db_file
        self.conn = None
        try:
            self.conn = sqlite3.connect(self.db_file, check_same_thread=False)
            self.conn.row_factory = sqlite3.Row
            self.init_db()
            logger.info(f"Database connection successful to {db_file}")
        except sqlite3.Error as e:
            logger.error(f"Database connection failed: {e}")
            messagebox.showerror("Database Error", f"Could not connect to database {db_file}.\n{e}")
            raise

    def init_db(self):
        """Initializes the database schema."""
        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS clients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                company TEXT,
                email TEXT,
                phone TEXT,
                address TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                client_id INTEGER,
                status TEXT,
                amount REAL,
                data TEXT, -- JSON blob for soumission extraction, quote details, etc.
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (client_id) REFERENCES clients (id)
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                category TEXT,
                content TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        self.conn.commit()
        logger.info("Database schema initialized.")

    def execute_query(self, query, params=(), fetch=None):
        """Generic query execution method."""
        try:
            cursor = self.conn.cursor()
            cursor.execute(query, params)
            if fetch == "one":
                result = cursor.fetchone()
            elif fetch == "all":
                result = cursor.fetchall()
            else:
                self.conn.commit()
                result = cursor.lastrowid
            return result
        except sqlite3.Error as e:
            logger.error(f"Database query failed: {query} with params {params}\nError: {e}")
            # Avoid showing popups for background errors
            # messagebox.showerror("Database Error", f"A database operation failed.\nCheck logs for details.")
            return None if fetch else False

    def close(self):
        if self.conn:
            self.conn.close()
            logger.info("Database connection closed.")


class AIProvider:
    """Abstraction layer for various AI providers."""

    def __init__(self, config_manager, log_queue):
        self.config = config_manager.config["ai_keys"]
        self.log_queue = log_queue

    def _log(self, message):
        self.log_queue.put(f"[AI] {message}")

    def get_active_provider(self):
        if self.config.get("gemini_api_key"): return "gemini"
        if self.config.get("openai_api_key"): return "openai"
        if self.config.get("anthropic_api_key"): return "anthropic"
        # Check if Ollama is running
        try:
            requests.get(self.config.get("ollama_base_url", "http://localhost:11434"), timeout=1)
            return "ollama"
        except requests.exceptions.RequestException:
            return None

    def ai_complete(self, prompt, max_tokens=4096, system=None, temperature=0.2, is_json=False):
        provider = self.get_active_provider()
        self._log(f"Using AI provider: {provider}")
        if not provider:
            self._log("No AI provider configured or available.")
            return {"error": "No AI provider is configured or available."}

        try:
            if provider == "gemini":
                return self._gemini_complete(prompt, max_tokens, system, temperature, is_json)
            elif provider == "openai":
                return self._openai_complete(prompt, max_tokens, system, temperature, is_json)
            elif provider == "anthropic":
                return self._anthropic_complete(prompt, max_tokens, system, temperature, is_json)
            elif provider == "ollama":
                return self._ollama_complete(prompt, max_tokens, system, temperature, is_json)
        except Exception as e:
            self._log(f"AI completion error with {provider}: {e}")
            logger.error(f"AI completion error with {provider}: {e}")
            return {"error": str(e)}

    def _gemini_complete(self, prompt, max_tokens, system, temperature, is_json):
        api_key = self.config["gemini_api_key"]
        if not genai:
            return {"error": "Google GenerativeAI library not installed. Please install it."}

        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')

        generation_config = genai.types.GenerationConfig(
            max_output_tokens=max_tokens,
            temperature=temperature,
            response_mime_type="application/json" if is_json else "text/plain",
        )

        full_prompt = []
        if system:
            full_prompt.append(system)
        full_prompt.append(prompt)

        response = model.generate_content(full_prompt, generation_config=generation_config)
        return {"content": response.text}

    def _openai_complete(self, prompt, max_tokens, system, temperature, is_json):
        api_key = self.config["openai_api_key"]
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

        messages = []
        if system:
            messages.append({"role": "system", "content": system})
        messages.append({"role": "user", "content": prompt})

        json_data = {
            "model": "gpt-4o",
            "messages": messages,
            "max_tokens": max_tokens,
            "temperature": temperature,
        }
        if is_json:
            json_data["response_format"] = {"type": "json_object"}

        response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=json_data, timeout=120)
        response.raise_for_status()
        return {"content": response.json()["choices"][0]["message"]["content"]}

    def _anthropic_complete(self, prompt, max_tokens, system, temperature, is_json):
        api_key = self.config["anthropic_api_key"]
        headers = {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json"
        }

        messages = [{"role": "user", "content": prompt}]

        json_data = {
            "model": "claude-3-haiku-20240307",
            "messages": messages,
            "max_tokens": max_tokens,
            "temperature": temperature,
        }
        if system:
            json_data["system"] = system

        response = requests.post("https://api.anthropic.com/v1/messages", headers=headers, json=json_data, timeout=120)
        response.raise_for_status()
        return {"content": response.json()["content"][0]["text"]}

    def _ollama_complete(self, prompt, max_tokens, system, temperature, is_json):
        url = self.config.get("ollama_base_url", "http://localhost:11434") + "/api/chat"

        messages = []
        if system:
            messages.append({"role": "system", "content": system})
        messages.append({"role": "user", "content": prompt})

        json_data = {
            "model": self.config.get("ollama_model", "llama3"),
            "messages": messages,
            "stream": False,
            "options": {
                "temperature": temperature,
                "num_predict": max_tokens
            }
        }
        if is_json:
            json_data["format"] = "json"

        response = requests.post(url, json=json_data, timeout=120)
        response.raise_for_status()
        return {"content": response.json()["message"]["content"]}


class AsyncHandler:
    """Handles running tasks in a separate thread to avoid blocking the GUI."""

    def __init__(self, app):
        self.app = app
        self.task_queue = queue.Queue()
        self.result_queue = queue.Queue()
        self.thread = threading.Thread(target=self._worker, daemon=True)
        self.thread.start()
        self.app.after(100, self._process_results)

    def _worker(self):
        while True:
            task, args, kwargs, callback = self.task_queue.get()
            try:
                result = task(*args, **kwargs)
                self.result_queue.put((callback, result))
            except Exception as e:
                logger.error(f"Async task failed: {e}", exc_info=True)
                self.result_queue.put((self.app.show_async_error, e))
            finally:
                self.task_queue.task_done()

    def run(self, task, args=(), kwargs=None, callback=None):
        if kwargs is None:
            kwargs = {}
        self.task_queue.put((task, args, kwargs, callback))

    def _process_results(self):
        try:
            while not self.result_queue.empty():
                callback, result = self.result_queue.get()
                if callback:
                    callback(result)
        finally:
            self.app.after(100, self._process_results)


# --- GUI COMPONENTS (TABS) ---


class BaseTab(ttk.Frame):
    def __init__(self, master, app_controller, **kwargs):
        super().__init__(master, padding=10, **kwargs)
        self.app = app_controller


class DashboardTab(BaseTab):
    def __init__(self, master, app_controller):
        super().__init__(master, app_controller)
        self.create_widgets()
        self.load_recent_projects()

    def create_widgets(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=BOTH, expand=True)

        header = ttk.Label(
            main_frame, text=i18n("welcome"), font=("", 24, "bold"), bootstyle=PRIMARY
        )
        header.pack(pady=(10, 20))

        # Quick Actions
        actions_frame = ttk.Labelframe(
            main_frame, text=i18n("quick_actions"), padding=15
        )
        actions_frame.pack(fill=X, pady=10)

        actions_frame.columnconfigure((0, 1, 2, 3), weight=1)

        btn_new_quote = ttk.Button(
            actions_frame,
            text=i18n("new_quote"),
            bootstyle=SUCCESS,
            command=lambda: self.app.notebook.select(self.app.tabs["quote_builder"]),
        )
        btn_new_quote.grid(row=0, column=0, padx=5, sticky="ew")

        btn_new_contract = ttk.Button(
            actions_frame,
            text=i18n("new_contract"),
            bootstyle=INFO,
            command=lambda: self.app.notebook.select(
                self.app.tabs["contract_builder"]
            ),
        )
        btn_new_contract.grid(row=0, column=1, padx=5, sticky="ew")

        btn_add_client = ttk.Button(
            actions_frame,
            text=i18n("add_client"),
            bootstyle=SECONDARY,
            command=lambda: self.app.notebook.select(self.app.tabs["client_manager"]),
        )
        btn_add_client.grid(row=0, column=2, padx=5, sticky="ew")

        btn_import = ttk.Button(
            actions_frame,
            text=i18n("import_soumission"),
            bootstyle=WARNING,
            command=lambda: self.app.tabs["soumission_reader"].load_file(),
        )
        btn_import.grid(row=0, column=3, padx=5, sticky="ew")

        # Recent Projects
        projects_frame = ttk.Labelframe(
            main_frame, text=i18n("recent_projects"), padding=15
        )
        projects_frame.pack(fill=BOTH, expand=True, pady=10)

        cols = [
            i18n("project_name"),
            i18n("client"),
            i18n("amount"),
            i18n("status"),
            i18n("created_date"),
        ]
        self.tree = ttk.Treeview(
            projects_frame, columns=cols, show="headings", bootstyle=PRIMARY
        )
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor=CENTER)
        self.tree.pack(fill=BOTH, expand=True)

    def load_recent_projects(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        query = """
            SELECT p.name, c.name as client_name, p.amount, p.status, p.created_at
            FROM projects p
            LEFT JOIN clients c ON p.client_id = c.id
            ORDER BY p.created_at DESC
            LIMIT 10
        """
        projects = self.app.db.execute_query(query, fetch="all")
        if projects:
            for proj in projects:
                amount_str = f"${proj['amount']:,.2f}" if proj['amount'] else "N/A"
                created_at_str = proj['created_at'].split(" ")[0] if proj['created_at'] else "N/A"
                self.tree.insert(
                    "",
                    END,
                    values=(
                        proj["name"],
                        proj["client_name"] or "N/A",
                        amount_str,
                        proj["status"],
                        created_at_str,
                    ),
                )
        self.app.after(30000, self.load_recent_projects) # Refresh every 30 seconds

class SoumissionReaderTab(BaseTab):
    def __init__(self, master, app_controller):
        super().__init__(master, app_controller)
        self.file_content = ""
        self.extracted_data = {}
        self.create_widgets()

    def create_widgets(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # Top controls
        top_frame = ttk.Frame(self)
        top_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        top_frame.columnconfigure(1, weight=1)

        btn_load = ttk.Button(top_frame, text=i18n("load_file"), command=self.load_file)
        btn_load.grid(row=0, column=0, sticky="w")

        self.status_label = ttk.Label(top_frame, text="")
        self.status_label.grid(row=0, column=1, padx=10, sticky="e")

        # Main content area
        notebook = ttk.Notebook(self)
        notebook.grid(row=1, column=0, sticky="nsew")

        # Review Tab
        review_frame = ttk.Frame(notebook, padding=10)
        notebook.add(review_frame, text=i18n("extraction_review"))
        review_frame.rowconfigure(0, weight=1)
        review_frame.columnconfigure(0, weight=1)

        scrolled_review = ScrolledFrame(review_frame, autohide=True)
        scrolled_review.grid(row=0, column=0, sticky="nsew")
        scrolled_review.columnconfigure(1, weight=1) # Make entry widgets expand

        self.review_widgets_frame = scrolled_review
        self.review_vars = {}

        # Line Items Tab
        items_frame = ttk.Frame(notebook, padding=10)
        notebook.add(items_frame, text=i18n("line_items_for_quote"))
        items_frame.rowconfigure(0, weight=1)
        items_frame.columnconfigure(0, weight=1)

        cols = [i18n("description"), i18n("quantity"), i18n("unit"), i18n("rate"), i18n("line_total")]
        self.items_tree = ttk.Treeview(items_frame, columns=cols, show="headings", bootstyle=INFO)
        for col in cols:
            self.items_tree.heading(col, text=col)
        self.items_tree.column(i18n("description"), width=400)
        self.items_tree.pack(fill=BOTH, expand=True)

        # Bottom controls
        bottom_frame = ttk.Frame(self)
        bottom_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))

        btn_push = ttk.Button(bottom_frame, text=i18n("push_to_quote_builder"), command=self.push_to_quote_builder, bootstyle=SUCCESS)
        btn_push.pack(side=RIGHT)

    def load_file(self):
        filepath = filedialog.askopenfilename(
            title="Select a Soumission File",
            filetypes=(
                ("All Supported", "*.pdf *.docx *.txt"),
                ("PDF Files", "*.pdf"),
                ("Word Documents", "*.docx"),
                ("Text Files", "*.txt"),
            ),
        )
        if not filepath:
            return

        self.status_label.config(text=i18n("parsing_in_progress"))
        self.update_idletasks()
        self.app.async_handler.run(self._extract_text, args=(filepath,), callback=self.on_text_extracted)

    def _extract_text(self, filepath):
        _, extension = os.path.splitext(filepath.lower())
        content = ""
        try:
            if extension == ".pdf":
                # First try with a robust text extractor
                try:
                    content = extract_text_from_pdf(filepath)
                except Exception as e:
                    logger.warning(f"pdfminer.six failed: {e}. Falling back to PyPDF2.")
                    reader = PdfReader(filepath)
                    content = "".join(page.extract_text() for page in reader.pages if page.extract_text())

                # If text quality is poor, try OCR
                if len(content.strip()) < 100:
                    logger.info("Low text content from PDF, attempting OCR.")
                    if pytesseract and convert_from_path and _optional_deps.get("pytesseract"):
                        try:
                            # This path might need to be configured by the user in settings
                            tesseract_path = self.app.config.get("tesseract_cmd_path", "")
                            if tesseract_path and os.path.exists(tesseract_path):
                                pytesseract.pytesseract.tesseract_cmd = tesseract_path
                            images = convert_from_path(filepath)
                            ocr_texts = [pytesseract.image_to_string(img) for img in images]
                            content = "\n".join(ocr_texts)
                            logger.info("OCR extraction successful.")
                        except Exception as ocr_error:
                            logger.warning(f"OCR failed: {ocr_error}. Using original extracted text.")
                            self.app.log_to_widget("OCR failed. Tesseract might not be installed or configured correctly.")
                    else:
                        logger.info("Tesseract not found, skipping OCR.")

            elif extension == ".docx":
                doc = Document(filepath)
                content = "\n".join([p.text for p in doc.paragraphs])
            elif extension == ".txt":
                with open(filepath, "rb") as f:
                    raw_data = f.read()
                    result = chardet.detect(raw_data)
                    encoding = result['encoding'] or 'utf-8'
                with open(filepath, "r", encoding=encoding, errors='ignore') as f:
                    content = f.read()
            return content
        except Exception as e:
            logger.error(f"Error reading file {filepath}: {e}")
            return f"Error: Could not read file. {e}"

    def on_text_extracted(self, content):
        if content.startswith("Error:"):
            self.status_label.config(text=content)
            messagebox.showerror("File Error", content)
            return

        self.file_content = content
        self.status_label.config(text=i18n("ai_extraction_in_progress"))
        self.app.log_to_widget("Starting AI extraction...")

        self.app.async_handler.run(self._run_ai_extraction, callback=self.on_ai_extraction_complete)

    def _run_ai_extraction(self):
        extraction_schema = """
        {
          "project": {"name": "", "location": "", "owner": "", "gc": "", "tender_number": "", "addenda": []},
          "timeline": {"start_date": "", "completion": "", "milestones": [], "penalties": ""},
          "items": [{"code": "", "description": "", "qty": 0, "unit": "", "unit_price": 0.0, "line_total": 0.0}],
          "pricing": {"subtotal": 0.0, "taxes": [{"name": "GST", "rate": 0.05, "amount": 0.0}, {"name": "QST", "rate": 0.09975, "amount": 0.0}], "total": 0.0},
          "allowances": [], "alternates": [], "exclusions": [],
          "bonding": "", "warranty": "", "holdback": "10%", "insurance": "", "permits": "", "safety": "",
          "change_orders": "", "signatures": [], "contacts": [],
          "risks": [{"note": "", "level": "low|medium|high"}],
          "confidence": 0.0
        }
        """
        system_prompt = f"""
        You are an expert assistant for construction project estimation in Canada.
        Your task is to extract structured information from the provided document text.
        Analyze the text and populate the fields in the following JSON schema.
        Be precise. If a value is not found, use null, an empty string "", 0 for numbers, or an empty list [].
        Do not invent information. For the 'items' array, extract every billable line item you can find.
        Calculate line_total if possible (qty * unit_price).
        Estimate your overall confidence in the extraction on a scale of 0.0 to 1.0.
        Your output MUST be a valid JSON object matching this schema, and nothing else.
        {extraction_schema}
        """

        prompt = f"Here is the document text to analyze:\n\n---\n\n{self.file_content[:24000]}\n\n---\n\nPlease extract the data into the specified JSON format."

        for attempt in range(2): # Retry logic
            result = self.app.ai_provider.ai_complete(prompt, system=system_prompt, is_json=True, temperature=0.1)
            if "content" in result:
                try:
                    json_str = result["content"].strip().replace("```json", "").replace("```", "")
                    self.extracted_data = json.loads(json_str)
                    return self.extracted_data
                except json.JSONDecodeError as e:
                    logger.error(f"AI extraction JSON decode error (attempt {attempt+1}): {e}")
                    self.app.log_to_widget(f"Warning: AI returned invalid JSON. Retrying...")
                    system_prompt += "\n\nIMPORTANT: Your previous response was not valid JSON. Please ensure your entire output is a single, valid JSON object and nothing else."
            else:
                return {"error": result.get("error", "Unknown AI error")}

        return {"error": "Failed to get valid JSON from AI after multiple attempts."}

    def on_ai_extraction_complete(self, data):
        self.status_label.config(text="")
        if "error" in data:
            messagebox.showerror("AI Error", f"Failed to extract data from document.\n{data['error']}")
            self.app.log_to_widget(f"AI Extraction Failed: {data['error']}")
            return

        self.app.log_to_widget("AI extraction complete. Populating review form.")
        self.populate_review_form()
        self.populate_items_tree()

    def populate_review_form(self):
        # Clear previous widgets
        for widget in self.review_widgets_frame.winfo_children():
            widget.destroy()
        self.review_vars.clear()

        row = 0
        confidence = self.extracted_data.get('confidence', 0.0)
        confidence_color = "success" if confidence > 0.7 else "warning" if confidence > 0.4 else "danger"

        ttk.Label(self.review_widgets_frame, text=f"Overall Confidence: {confidence:.1%}", bootstyle=confidence_color, font=("", 12, "bold")).grid(row=row, column=0, columnspan=3, pady=10, sticky='w')
        row += 1

        def create_entry(parent, key_path, value, r):
            ttk.Label(parent, text=key_path).grid(row=r, column=0, sticky='w', padx=5, pady=2)
            var = StringVar(value=str(value))
            entry = ttk.Entry(parent, textvariable=var, width=80)
            entry.grid(row=r, column=1, sticky='ew', padx=5, pady=2)
            self.review_vars[key_path] = var

        def flatten_dict(d, parent_key='', sep='.'):
            items = []
            for k, v in d.items():
                new_key = parent_key + sep + k if parent_key else k
                if isinstance(v, dict):
                    items.extend(flatten_dict(v, new_key, sep=sep).items())
                elif isinstance(v, list) and k != 'items':
                     items.append((new_key, json.dumps(v)))
                elif k != 'items':
                    items.append((new_key, v))
            return dict(items)

        flat_data = flatten_dict(self.extracted_data)
        for key, value in flat_data.items():
            create_entry(self.review_widgets_frame, key, value, row)
            row += 1

    def populate_items_tree(self):
        for i in self.items_tree.get_children():
            self.items_tree.delete(i)

        items = self.extracted_data.get("items", [])
        if not items or not isinstance(items, list): return

        for item in items:
            if not isinstance(item, dict): continue
            try:
                qty = float(item.get('qty', 0))
                rate = float(item.get('unit_price', 0.0))
                total = qty * rate
                self.items_tree.insert("", END, values=(
                    item.get('description', ''),
                    f"{qty:.2f}",
                    item.get('unit', ''),
                    f"{rate:,.2f}",
                    f"{total:,.2f}"
                ))
            except (ValueError, TypeError):
                self.items_tree.insert("", END, values=(
                    item.get('description', ''), item.get('qty', 'N/A'),
                    item.get('unit', ''), item.get('unit_price', 'N/A'),
                    item.get('line_total', 'N/A')
                ))

    def push_to_quote_builder(self):
        items_to_push = []
        for child in self.items_tree.get_children():
            values = self.items_tree.item(child)['values']
            try:
                item_data = {
                    'description': values[0],
                    'qty': float(values[1]),
                    'unit': values[2],
                    'rate': float(str(values[3]).replace(',', '')),
                }
                items_to_push.append(item_data)
            except (ValueError, IndexError):
                messagebox.showwarning("Data Error", "Could not parse one of the item lines. It will be skipped.")
                continue

        if not items_to_push:
            messagebox.showinfo("No Items", "No valid items to push to the Quote Builder.")
            return

        quote_builder_tab = self.app.tabs['quote_builder']
        quote_builder_tab.load_items(items_to_push)

        project_name = self.review_vars.get('project.name', StringVar()).get()
        if project_name:
            quote_builder_tab.project_name_var.set(project_name)

        self.app.notebook.select(quote_builder_tab)
        messagebox.showinfo("Success", f"{len(items_to_push)} items have been added to the Quote Builder.")

class QuoteBuilderTab(BaseTab):
    def __init__(self, master, app_controller):
        super().__init__(master, app_controller)
        self.create_widgets()
        self.update_totals()

    def create_widgets(self):
        self.columnconfigure(0, weight=3)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)

        top_frame = ttk.Frame(self)
        top_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 10))
        top_frame.columnconfigure(1, weight=1)

        ttk.Label(top_frame, text=i18n("project_name") + ":").grid(row=0, column=0, padx=(0,5), sticky='w')
        self.project_name_var = StringVar()
        ttk.Entry(top_frame, textvariable=self.project_name_var).grid(row=0, column=1, sticky='ew')

        ttk.Label(top_frame, text=i18n("client") + ":").grid(row=1, column=0, padx=(0,5), sticky='w')
        self.client_var = StringVar()
        self.client_selector = ttk.Combobox(top_frame, textvariable=self.client_var)
        self.client_selector.grid(row=1, column=1, sticky='ew', pady=(5,0))
        self.client_selector.bind("<FocusIn>", lambda e: self.refresh_client_list())

        items_frame = ttk.Labelframe(self, text="Items", padding=10)
        items_frame.grid(row=1, column=0, sticky='nsew')
        items_frame.columnconfigure(0, weight=1)
        items_frame.rowconfigure(0, weight=1)

        cols = [i18n("quantity"), i18n("description"), i18n("unit"), i18n("rate"), i18n("line_total")]
        self.items_tree = ttk.Treeview(items_frame, columns=cols, show="headings")
        self.items_tree.grid(row=0, column=0, sticky='nsew')

        vsb = ttk.Scrollbar(items_frame, orient="vertical", command=self.items_tree.yview)
        vsb.grid(row=0, column=1, sticky='ns')
        self.items_tree.configure(yscrollcommand=vsb.set)

        for col in cols:
            self.items_tree.heading(col, text=col)
        self.items_tree.column(i18n("description"), width=400, anchor=W)
        for col in [i18n("quantity"), i18n("unit"), i18n("rate"), i18n("line_total")]:
            self.items_tree.column(col, width=100, anchor=E)

        entry_frame = ttk.Frame(items_frame)
        entry_frame.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(10,0))
        entry_frame.columnconfigure(1, weight=1)

        self.qty_var = DoubleVar(value=1.0)
        self.desc_var = StringVar()
        self.unit_var = StringVar()
        self.rate_var = DoubleVar(value=0.0)

        ttk.Entry(entry_frame, textvariable=self.qty_var, width=8).grid(row=0, column=0, padx=(0,5))
        ttk.Entry(entry_frame, textvariable=self.desc_var).grid(row=0, column=1, padx=5, sticky='ew')
        ttk.Entry(entry_frame, textvariable=self.unit_var, width=10).grid(row=0, column=2, padx=5)
        ttk.Entry(entry_frame, textvariable=self.rate_var, width=12).grid(row=0, column=3, padx=5)

        btn_add = ttk.Button(entry_frame, text=i18n("add_item"), command=self.add_item, bootstyle=SUCCESS)
        btn_add.grid(row=0, column=4, padx=5)

        btn_del = ttk.Button(entry_frame, text=i18n("delete_item"), command=self.delete_item, bootstyle=DANGER)
        btn_del.grid(row=0, column=5, padx=5)

        sidebar = ttk.Frame(self)
        sidebar.grid(row=1, column=1, rowspan=2, sticky='nsew', padx=(10,0))
        sidebar.columnconfigure(0, weight=1)

        totals_frame = ttk.Labelframe(sidebar, text="Totals", padding=15)
        totals_frame.pack(fill=X, expand=False)
        totals_frame.columnconfigure(1, weight=1)

        self.totals_vars = {
            "subtotal": StringVar(), "taxes": StringVar(), "grand_total": StringVar(),
            "holdback_amount": StringVar(), "total_with_holdback": StringVar(), "total_without_holdback": StringVar()
        }

        row = 0
        for key, name in [("subtotal", i18n("subtotal")), ("taxes", i18n("taxes")), ("grand_total", i18n("grand_total"))]:
            ttk.Label(totals_frame, text=f"{name}:", font=("", 10, "bold")).grid(row=row, column=0, sticky='w', pady=2)
            ttk.Label(totals_frame, textvariable=self.totals_vars[key], font=("", 10, "bold"), anchor=E).grid(row=row, column=1, sticky='ew', pady=2)
            row += 1

        self.holdback_frame = ttk.Frame(totals_frame)
        self.holdback_frame.grid(row=row, column=0, columnspan=2, sticky='ew', pady=(10,0))
        self.holdback_frame.columnconfigure(1, weight=1)
        self.holdback_label = ttk.Label(self.holdback_frame, text="", font=("", 9, "italic"))
        self.holdback_label.grid(row=0, column=0, sticky='w')
        ttk.Label(self.holdback_frame, textvariable=self.totals_vars["holdback_amount"], font=("", 9, "italic"), anchor=E).grid(row=0, column=1, sticky='ew')

        tax_frame = ttk.Labelframe(sidebar, text="Tax Engine", padding=15)
        tax_frame.pack(fill=X, expand=False, pady=10)
        tax_frame.columnconfigure(1, weight=1)

        ttk.Label(tax_frame, text=i18n("province") + ":").grid(row=0, column=0, sticky='w')
        self.province_var = StringVar(value=self.app.config.get("defaults.province", "ON"))
        province_menu = ttk.Combobox(tax_frame, textvariable=self.province_var, values=list(PROVINCES.keys()))
        province_menu.grid(row=0, column=1, sticky='ew')
        province_menu.bind("<<ComboboxSelected>>", lambda e: self.update_totals())

        action_frame = ttk.Frame(sidebar)
        action_frame.pack(fill=X, pady=20)
        btn_export = ttk.Button(action_frame, text=i18n("export_pdf"), command=self.export_pdf, bootstyle=PRIMARY)
        btn_export.pack(fill=X, pady=(0, 5))

        btn_email = ttk.Button(action_frame, text=i18n("send_email"), command=self.send_email, bootstyle=INFO)
        btn_email.pack(fill=X)

    def add_item(self):
        try:
            qty = self.qty_var.get()
            desc = self.desc_var.get()
            unit = self.unit_var.get()
            rate = self.rate_var.get()
            if not desc:
                messagebox.showwarning("Input Error", "Description cannot be empty.")
                return

            total = qty * rate
            self.items_tree.insert("", END, values=(f"{qty:.2f}", desc, unit, f"{rate:,.2f}", f"{total:,.2f}"))
            self.update_totals()

            self.qty_var.set(1.0)
            self.desc_var.set("")
            self.unit_var.set("")
            self.rate_var.set(0.0)

        except Exception as e:
            messagebox.showerror("Input Error", f"Invalid input. Please check your numbers.\n{e}")

    def delete_item(self):
        selected = self.items_tree.selection()
        if not selected:
            messagebox.showwarning("Selection Error", "Please select an item to delete.")
            return
        for item in selected:
            self.items_tree.delete(item)
        self.update_totals()

    def load_items(self, items_list):
        for i in self.items_tree.get_children():
            self.items_tree.delete(i)

        for item in items_list:
             qty = item.get('qty', 1.0)
             rate = item.get('rate', 0.0)
             total = qty * rate
             self.items_tree.insert("", END, values=(
                 f"{qty:.2f}", item.get('description', ''),
                 item.get('unit', ''), f"{rate:,.2f}", f"{total:,.2f}"
             ))
        self.update_totals()

    def update_totals(self):
        subtotal = 0.0
        for child in self.items_tree.get_children():
            try:
                total_str = self.items_tree.item(child)['values'][4]
                subtotal += float(str(total_str).replace(',', ''))
            except (ValueError, IndexError):
                continue

        province = self.province_var.get()
        tax_info = TAX_STRUCTURE.get(province, {})

        tax_total = 0.0
        tax_breakdown = []
        for tax_name, rate in tax_info.items():
            tax_amount = subtotal * rate
            tax_total += tax_amount
            tax_breakdown.append(f"{tax_name}: ${tax_amount:,.2f}")

        grand_total = subtotal + tax_total

        self.totals_vars['subtotal'].set(f"${subtotal:,.2f}")
        self.totals_vars['taxes'].set(f"${tax_total:,.2f} ({', '.join(tax_breakdown)})")
        self.totals_vars['grand_total'].set(f"${grand_total:,.2f}")

        holdback_pct = self.app.config.get('defaults.holdback_percent', 10.0)
        if holdback_pct > 0:
            self.holdback_frame.grid()
            holdback_amount = grand_total * (holdback_pct / 100.0)
            self.holdback_label.config(text=i18n("construction_holdback", percent=holdback_pct))
            self.totals_vars["holdback_amount"].set(f"(${holdback_amount:,.2f})")
        else:
            self.holdback_frame.grid_remove()

    def refresh_client_list(self):
        clients = self.app.db.execute_query("SELECT id, name, company FROM clients ORDER BY name", fetch="all")
        if clients:
            self.client_list = {f"{c['name']} ({c['company'] or 'N/A'})": c['id'] for c in clients}
            self.client_selector['values'] = list(self.client_list.keys())
        else:
            self.client_list = {}
            self.client_selector['values'] = []

    def export_pdf(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Documents", "*.pdf")],
            title="Save Quote As PDF"
        )
        if not filename: return

        try:
            doc = SimpleDocTemplate(filename, pagesize=letter, rightMargin=inch/2, leftMargin=inch/2, topMargin=inch/2, bottomMargin=inch/2)
            styles = getSampleStyleSheet()

            styles.add(ParagraphStyle(name='CompanyHeader', fontName='Helvetica-Bold', fontSize=16, textColor=HexColor(self.style.colors.primary)))
            styles.add(ParagraphStyle(name='QuoteTitle', fontName='Helvetica-Bold', fontSize=20, alignment=TA_RIGHT, textColor=HexColor(self.style.colors.secondary)))
            styles.add(ParagraphStyle(name='SectionHeader', fontName='Helvetica-Bold', fontSize=12, spaceBefore=12, spaceAfter=6, textColor=HexColor(self.style.colors.primary)))
            styles.add(ParagraphStyle(name='RightAlign', alignment=TA_RIGHT))

            company_info = self.app.config.get('company_info')
            story = []

            header_data = [
                [Paragraph(company_info.get('name', ''), styles['CompanyHeader']), Paragraph('QUOTE', styles['QuoteTitle'])],
                [Paragraph(company_info.get('address', ''), styles['Normal']), ''],
                [Paragraph(company_info.get('phone', ''), styles['Normal']), ''],
                [Paragraph(company_info.get('email', ''), styles['Normal']), Paragraph(f"Date: {datetime.date.today().isoformat()}", styles['RightAlign'])],
            ]
            header_table = Table(header_data, colWidths=[4*inch, 3.5*inch])
            header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
            story.append(header_table)
            story.append(Spacer(1, 0.25*inch))

            story.append(Paragraph("Project: " + self.project_name_var.get(), styles['SectionHeader']))

            items_data = [['Qty', 'Description', 'Unit', 'Rate', 'Amount']]
            for child in self.items_tree.get_children():
                values = self.items_tree.item(child)['values']
                items_data.append([
                    Paragraph(str(v), styles['Normal']) if i !=1 else Paragraph(str(v), styles['BodyText']) for i, v in enumerate(values)
                ])

            items_table = Table(items_data, colWidths=[0.7*inch, 4.2*inch, 0.7*inch, 1*inch, 1*inch])
            items_table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), HexColor(self.style.colors.primary)),
                ('TEXTCOLOR', (0,0), (-1,0), white),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('ALIGN', (0,0), (0,-1), 'CENTER'),
                ('ALIGN', (-2,0), (-1,-1), 'RIGHT'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,0), 12),
                ('GRID', (0,0), (-1,-1), 1, black)
            ]))
            story.append(items_table)
            story.append(Spacer(1, 0.25*inch))

            subtotal = self.totals_vars['subtotal'].get()
            taxes = self.totals_vars['taxes'].get()
            total = self.totals_vars['grand_total'].get()
            totals_data = [['Subtotal:', subtotal], ['Taxes:', taxes], ['Total:', total]]
            totals_table = Table(totals_data, colWidths=[6*inch, 1.5*inch])
            totals_table.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
                ('FONTNAME', (0,2), (-1,2), 'Helvetica-Bold'),
            ]))
            story.append(totals_table)

            doc.build(story)
            messagebox.showinfo("Success", f"Quote successfully saved to {filename}")
        except Exception as e:
            logger.error(f"Failed to generate PDF: {e}", exc_info=True)
            messagebox.showerror("PDF Error", f"Could not create PDF file.\n{e}")

    def send_email(self):
        # First, generate the PDF to a temporary path
        import tempfile
        temp_dir = tempfile.gettempdir()
        # Use a more descriptive name for the temp file
        pdf_path = os.path.join(temp_dir, f"quote_{self.project_name_var.get().replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.pdf")

        # Call existing export_pdf logic but pass a filename
        # This avoids code duplication. Let's refactor export_pdf slightly.
        # For now, I will duplicate the code as refactoring export_pdf is more risky.
        try:
            doc = SimpleDocTemplate(pdf_path, pagesize=letter, rightMargin=inch/2, leftMargin=inch/2, topMargin=inch/2, bottomMargin=inch/2)
            styles = getSampleStyleSheet()

            styles.add(ParagraphStyle(name='CompanyHeader', fontName='Helvetica-Bold', fontSize=16, textColor=HexColor(self.style.colors.primary)))
            styles.add(ParagraphStyle(name='QuoteTitle', fontName='Helvetica-Bold', fontSize=20, alignment=TA_RIGHT, textColor=HexColor(self.style.colors.secondary)))
            styles.add(ParagraphStyle(name='SectionHeader', fontName='Helvetica-Bold', fontSize=12, spaceBefore=12, spaceAfter=6, textColor=HexColor(self.style.colors.primary)))
            styles.add(ParagraphStyle(name='RightAlign', alignment=TA_RIGHT))

            company_info = self.app.config.get('company_info')
            story = []

            header_data = [
                [Paragraph(company_info.get('name', ''), styles['CompanyHeader']), Paragraph('QUOTE', styles['QuoteTitle'])],
                [Paragraph(company_info.get('address', ''), styles['Normal']), ''],
                [Paragraph(company_info.get('phone', ''), styles['Normal']), ''],
                [Paragraph(company_info.get('email', ''), styles['Normal']), Paragraph(f"Date: {datetime.date.today().isoformat()}", styles['RightAlign'])],
            ]
            header_table = Table(header_data, colWidths=[4*inch, 3.5*inch])
            header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
            story.append(header_table)
            story.append(Spacer(1, 0.25*inch))

            story.append(Paragraph("Project: " + self.project_name_var.get(), styles['SectionHeader']))

            items_data = [['Qty', 'Description', 'Unit', 'Rate', 'Amount']]
            for child in self.items_tree.get_children():
                values = self.items_tree.item(child)['values']
                items_data.append([
                    Paragraph(str(v), styles['Normal']) if i !=1 else Paragraph(str(v), styles['BodyText']) for i, v in enumerate(values)
                ])

            items_table = Table(items_data, colWidths=[0.7*inch, 4.2*inch, 0.7*inch, 1*inch, 1*inch])
            items_table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), HexColor(self.style.colors.primary)),
                ('TEXTCOLOR', (0,0), (-1,0), white),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('ALIGN', (0,0), (0,-1), 'CENTER'),
                ('ALIGN', (-2,0), (-1,-1), 'RIGHT'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,0), 12),
                ('GRID', (0,0), (-1,-1), 1, black)
            ]))
            story.append(items_table)
            story.append(Spacer(1, 0.25*inch))

            subtotal = self.totals_vars['subtotal'].get()
            taxes = self.totals_vars['taxes'].get()
            total = self.totals_vars['grand_total'].get()
            totals_data = [['Subtotal:', subtotal], ['Taxes:', taxes], ['Total:', total]]
            totals_table = Table(totals_data, colWidths=[6*inch, 1.5*inch])
            totals_table.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
                ('FONTNAME', (0,2), (-1,2), 'Helvetica-Bold'),
            ]))
            story.append(totals_table)

            doc.build(story)
        except Exception as e:
            logger.error(f"Failed to generate temporary PDF for email: {e}", exc_info=True)
            messagebox.showerror("PDF Error", f"Could not create PDF for email attachment.\n{e}")
            return

        # Now, send the email
        smtp_config = self.app.config.get("smtp")
        client_email = simpledialog.askstring("Client Email", "Enter the recipient's email address:", parent=self)
        if not client_email:
            os.remove(pdf_path) # Clean up if user cancels
            return

        try:
            msg = MIMEMultipart()
            msg['From'] = smtp_config.get("username")
            msg['To'] = client_email
            msg['Subject'] = f"Quote for Project: {self.project_name_var.get()}"

            body = f"Dear Client,\n\nPlease find the attached quote for project '{self.project_name_var.get()}'.\n\nBest regards,\n{self.app.config.get('company_info.name')}"
            msg.attach(MIMEText(body, 'plain'))

            with open(pdf_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())

            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(pdf_path)}")
            msg.attach(part)

            server = smtplib.SMTP(smtp_config.get("server"), smtp_config.get("port"))
            server.starttls()
            server.login(smtp_config.get("username"), self.app.config.get("smtp.password"))
            text = msg.as_string()
            server.sendmail(smtp_config.get("username"), client_email, text)
            server.quit()

            messagebox.showinfo("Success", f"Email sent successfully to {client_email}")
            logger.info(f"Email sent to {client_email}")
        except Exception as e:
            logger.error(f"Failed to send email: {e}", exc_info=True)
            messagebox.showerror("Email Error", f"Could not send email.\nCheck SMTP settings and credentials.\n{e}")
        finally:
            os.remove(pdf_path) # Clean up the temporary file

class ContractBuilderTab(BaseTab):
    def __init__(self, master, app_controller):
        super().__init__(master, app_controller)
        self.create_widgets()

    def create_widgets(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.text_editor = ScrolledText(self, wrap=WORD, autohide=True, height=20)
        self.text_editor.grid(row=0, column=0, columnspan=2, sticky="nsew")

        # Insert some default text
        self.text_editor.insert(END, "<h1>Contract Title</h1>\n\n<p>This is a sample contract...</p>")

        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=1, column=1, sticky='e', pady=10)

        btn_export = ttk.Button(btn_frame, text="Export Contract to PDF", command=self.export_contract_pdf, bootstyle=SUCCESS)
        btn_export.pack()

    def export_contract_pdf(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Documents", "*.pdf")],
            title="Save Contract As PDF"
        )
        if not filename: return

        try:
            doc = SimpleDocTemplate(filename, pagesize=letter)
            styles = getSampleStyleSheet()
            # A simple parser for basic HTML tags, not a full-featured one
            story = []
            content = self.text_editor.get("1.0", END)
            for line in content.split('\n'):
                line = line.strip()
                if line.startswith("<h1>") and line.endswith("</h1>"):
                    story.append(Paragraph(line[4:-5], styles['h1']))
                elif line.startswith("<h2>") and line.endswith("</h2>"):
                    story.append(Paragraph(line[4:-5], styles['h2']))
                elif line.startswith("<p>") and line.endswith("</p>"):
                    story.append(Paragraph(line[3:-4], styles['Normal']))
                elif line:
                    story.append(Paragraph(line, styles['Normal']))
                story.append(Spacer(1, 12))

            doc.build(story)
            messagebox.showinfo("Success", f"Contract successfully saved to {filename}")
        except Exception as e:
            logger.error(f"Failed to generate contract PDF: {e}", exc_info=True)
            messagebox.showerror("PDF Error", f"Could not create PDF file.\n{e}")

class ClientManagerTab(BaseTab):
    def __init__(self, master, app_controller):
        super().__init__(master, app_controller)
        self.create_widgets()
        self.load_clients()

    def create_widgets(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        form_frame = ttk.Labelframe(self, text="Client Details", padding=10)
        form_frame.grid(row=0, column=0, sticky='ew', pady=(0, 10))
        form_frame.columnconfigure(1, weight=1)

        self.client_vars = {
            "id": IntVar(value=0), "name": StringVar(), "company": StringVar(),
            "email": StringVar(), "phone": StringVar(), "address": StringVar()
        }

        fields = ["name", "company", "email", "phone", "address"]
        for i, field in enumerate(fields):
            ttk.Label(form_frame, text=f"{field.title()}:").grid(row=i, column=0, sticky='w', pady=2)
            ttk.Entry(form_frame, textvariable=self.client_vars[field]).grid(row=i, column=1, columnspan=2, sticky='ew', pady=2)

        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=len(fields), column=1, sticky='e', pady=5)
        ttk.Button(btn_frame, text="Save", command=self.save_client, bootstyle=SUCCESS).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="New", command=self.new_client).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="Delete", command=self.delete_client, bootstyle=DANGER).pack(side=LEFT, padx=5)

        list_frame = ttk.Labelframe(self, text="Client List", padding=10)
        list_frame.grid(row=1, column=0, sticky='nsew')
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        cols = ["ID", "Name", "Company", "Email"]
        self.client_tree = ttk.Treeview(list_frame, columns=cols, show="headings")
        self.client_tree.grid(row=0, column=0, sticky='nsew')
        for col in cols: self.client_tree.heading(col, text=col)
        self.client_tree.column("ID", width=50)
        self.client_tree.bind("<<TreeviewSelect>>", self.on_client_select)

    def load_clients(self):
        for i in self.client_tree.get_children(): self.client_tree.delete(i)
        clients = self.app.db.execute_query("SELECT id, name, company, email FROM clients ORDER BY name", fetch="all")
        if clients:
            for client in clients:
                self.client_tree.insert("", END, values=list(client))

    def on_client_select(self, event):
        selected_item = self.client_tree.selection()
        if not selected_item: return
        client_id = self.client_tree.item(selected_item[0])['values'][0]
        client_data = self.app.db.execute_query("SELECT * FROM clients WHERE id = ?", (client_id,), fetch="one")
        if client_data:
            for key in self.client_vars:
                if key in client_data.keys():
                    self.client_vars[key].set(client_data[key] or "")

    def new_client(self):
        for key, var in self.client_vars.items():
            var.set(0 if isinstance(var, IntVar) else "")

    def save_client(self):
        client_id = self.client_vars["id"].get()
        data = {k: v.get() for k, v in self.client_vars.items() if k != 'id'}

        if client_id == 0: # New client
            query = "INSERT INTO clients (name, company, email, phone, address) VALUES (?, ?, ?, ?, ?)"
            params = (data['name'], data['company'], data['email'], data['phone'], data['address'])
        else: # Update existing
            query = "UPDATE clients SET name=?, company=?, email=?, phone=?, address=? WHERE id=?"
            params = (data['name'], data['company'], data['email'], data['phone'], data['address'], client_id)

        self.app.db.execute_query(query, params)
        self.load_clients()
        self.new_client()

    def delete_client(self):
        client_id = self.client_vars["id"].get()
        if client_id == 0:
            messagebox.showwarning("Warning", "No client selected to delete.")
            return
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this client?"):
            self.app.db.execute_query("DELETE FROM clients WHERE id=?", (client_id,))
            self.load_clients()
            self.new_client()

class TemplatesTab(BaseTab):
    def __init__(self, master, app_controller):
        super().__init__(master, app_controller)

        self.templates = {
            "Construction General": "<h1>General Construction Contract</h1>\n<p>This contract is between [Client Name] and [Your Company Name]...",
            "Electrical Services": "<h1>Electrical Services Agreement</h1>\n<p>Scope of work includes all electrical installations as per plan...",
            "Consulting Agreement": "<h1>Consulting Services Agreement</h1>\n<p>This agreement outlines the consulting services to be provided...",
            "Marketing Proposal": "<h1>Marketing Campaign Proposal</h1>\n<p>This document details the proposed marketing strategy...",
        }

        self.create_widgets()
        self.load_templates()

    def create_widgets(self):
        from tkinter import Listbox, SINGLE, END
        self.columnconfigure(1, weight=3)
        self.rowconfigure(0, weight=1)

        list_frame = ttk.Labelframe(self, text="Templates", padding=10)
        list_frame.grid(row=0, column=0, sticky='ns', padx=(0, 10))
        self.template_list = Listbox(list_frame, exportselection=False, selectmode=SINGLE)
        self.template_list.pack(fill='y', expand=True)
        self.template_list.bind('<<ListboxSelect>>', self.on_template_select)

        preview_frame = ttk.Labelframe(self, text="Preview", padding=10)
        preview_frame.grid(row=0, column=1, sticky='nsew')
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)

        self.preview_text = ScrolledText(preview_frame, wrap=WORD, autohide=True)
        self.preview_text.grid(row=0, column=0, sticky='nsew')
        self.preview_text.text['state'] = 'disabled'

        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=1, column=1, sticky='e', pady=10)
        btn_use = ttk.Button(btn_frame, text="Use This Template in Contract Builder", command=self.use_template)
        btn_use.pack()

    def load_templates(self):
        from tkinter import END
        self.template_list.delete(0, END)
        for name in self.templates.keys():
            self.template_list.insert(END, name)

    def on_template_select(self, event):
        from tkinter import NORMAL, DISABLED, END
        selection_indices = self.template_list.curselection()
        if not selection_indices: return

        selected_template_name = self.template_list.get(selection_indices[0])
        content = self.templates[selected_template_name]
        self.preview_text.text['state'] = 'normal'
        self.preview_text.delete('1.0', END)
        self.preview_text.insert(END, content)
        self.preview_text.text['state'] = 'disabled'

    def use_template(self):
        from tkinter import END
        selection_indices = self.template_list.curselection()
        if not selection_indices:
            messagebox.showwarning("No Selection", "Please select a template to use.")
            return

        selected_template_name = self.template_list.get(selection_indices[0])
        content = self.templates[selected_template_name]
        contract_tab = self.app.tabs['contract_builder']
        contract_tab.text_editor.delete('1.0', END)
        contract_tab.text_editor.insert(END, content)
        self.app.notebook.select(contract_tab)

class SettingsTab(BaseTab):
    def __init__(self, master, app_controller):
        super().__init__(master, app_controller)
        self.vars = {}
        self.create_widgets()
        self.load_settings()

    def create_widgets(self):
        scrolled_frame = ScrolledFrame(self, autohide=True)
        scrolled_frame.pack(fill=BOTH, expand=True)
        container = scrolled_frame

        # Company Info
        comp_frame = ttk.Labelframe(container, text="Company Information", padding=15)
        comp_frame.pack(fill=X, pady=10, padx=10)
        comp_frame.columnconfigure(1, weight=1)

        fields = ["name", "address", "phone", "email"]
        for i, field in enumerate(fields):
            ttk.Label(comp_frame, text=f"{field.title()}:").grid(row=i, column=0, sticky='w', pady=2)
            self.vars[f"company_info.{field}"] = StringVar()
            ttk.Entry(comp_frame, textvariable=self.vars[f"company_info.{field}"]).grid(row=i, column=1, sticky='ew', pady=2)

        # Defaults
        def_frame = ttk.Labelframe(container, text="Defaults & UI", padding=15)
        def_frame.pack(fill=X, pady=10, padx=10)
        def_frame.columnconfigure(1, weight=1)

        ttk.Label(def_frame, text="Default Province:").grid(row=0, column=0, sticky='w')
        self.vars["defaults.province"] = StringVar()
        ttk.Combobox(def_frame, textvariable=self.vars["defaults.province"], values=list(PROVINCES.keys())).grid(row=0, column=1, sticky='ew')

        ttk.Label(def_frame, text="Theme:").grid(row=1, column=0, sticky='w')
        self.vars["theme"] = StringVar()
        ttk.Combobox(def_frame, textvariable=self.vars["theme"], values=self.app.style.theme_names()).grid(row=1, column=1, sticky='ew')

        # AI Keys
        ai_frame = ttk.Labelframe(container, text="AI API Keys", padding=15)
        ai_frame.pack(fill=X, pady=10, padx=10)
        ai_frame.columnconfigure(1, weight=1)

        ai_keys = ["gemini_api_key", "openai_api_key", "anthropic_api_key", "ollama_base_url"]
        for i, key in enumerate(ai_keys):
            ttk.Label(ai_frame, text=f"{key}:").grid(row=i, column=0, sticky='w', pady=2)
            self.vars[f"ai_keys.{key}"] = StringVar()
            ttk.Entry(ai_frame, textvariable=self.vars[f"ai_keys.{key}"], show="*" if "key" in key else "").grid(row=i, column=1, sticky='ew', pady=2)

        # License
        lic_frame = ttk.Labelframe(container, text="License", padding=15)
        lic_frame.pack(fill=X, pady=10, padx=10)
        lic_frame.columnconfigure(1, weight=1)

        self.vars["license.key"] = StringVar()
        ttk.Label(lic_frame, text="License Key:").grid(row=0, column=0, sticky='w')
        ttk.Entry(lic_frame, textvariable=self.vars["license.key"]).grid(row=0, column=1, sticky='ew')

        self.license_status_label = ttk.Label(lic_frame, text="")
        self.license_status_label.grid(row=1, column=0, columnspan=2, pady=(5,0))

        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill=X, pady=20, padx=10)

        btn_save = ttk.Button(btn_frame, text="Save Settings", command=self.save_settings, bootstyle=SUCCESS)
        btn_save.pack(side=RIGHT)

    def load_settings(self):
        for key, var in self.vars.items():
            var.set(self.app.config.get(key, ''))

        if self.vars["license.key"].get():
            self.license_status_label.config(text="Status: QuoteForgeAI Pro, Lifetime License, $299 CAD (Active)", bootstyle=SUCCESS)
        else:
            self.license_status_label.config(text="Status: Unlicensed. Please enter a license key.", bootstyle=WARNING)

    def save_settings(self):
        for key, var in self.vars.items():
            self.app.config.set(key, var.get())

        self.app.config.save_config()
        self.load_settings()
        messagebox.showinfo("Success", "Settings have been saved.")
        self.app.apply_theme()

class LogsTab(BaseTab):
    def __init__(self, master, app_controller):
        super().__init__(master, app_controller)
        self.create_widgets()
        # The queue and handler are now created in the main app's __init__
        self.after(100, self.process_log_queue)

    def create_widgets(self):
        self.log_text = ScrolledText(self, wrap=WORD, state=DISABLED, autohide=True)
        self.log_text.pack(fill=BOTH, expand=True)

    def process_log_queue(self):
        while not self.app.log_queue.empty():
            message = self.app.log_queue.get()
            self.log_text.config(state=NORMAL)
            self.log_text.insert(END, message + '\n')
            self.log_text.config(state=DISABLED)
            self.log_text.yview(END)
        self.after(100, self.process_log_queue)

# --- MAIN APPLICATION ---

class QuoteForgeAIProApp(ttk.Window):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.title(i18n("app_title"))
        self.geometry("1280x800")
        center_window(self, 1280, 800)

        self.log_queue = queue.Queue()
        self.queue_handler = QueueHandler(self.log_queue)
        logger.addHandler(self.queue_handler)

        self.config = ConfigManager()
        self.db = DBManager(DB_FILE)
        self.async_handler = AsyncHandler(self)
        self.ai_provider = AIProvider(self.config, self.log_queue)

        global current_lang
        current_lang = self.config.get("defaults.language", "en")

        self.apply_theme()

        self.create_widgets()
        logger.info(f"{APP_NAME} v{APP_VERSION} started successfully.")

    def apply_theme(self):
        theme_name = self.config.get("theme", "darkly")
        self.style.theme_use(theme_name)

    def create_widgets(self):
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=BOTH, expand=True)

        self.tabs = {}
        tab_map = {
            "dashboard": (DashboardTab, i18n("dashboard")),
            "soumission_reader": (SoumissionReaderTab, i18n("soumission_reader")),
            "quote_builder": (QuoteBuilderTab, i18n("quote_builder")),
            "contract_builder": (ContractBuilderTab, i18n("contract_builder")),
            "client_manager": (ClientManagerTab, i18n("client_manager")),
            "templates": (TemplatesTab, i18n("templates")),
            "settings": (SettingsTab, i18n("settings")),
            "logs": (LogsTab, i18n("logs")),
        }

        for key, (tab_class, title) in tab_map.items():
            tab = tab_class(self.notebook, self)
            self.tabs[key] = tab
            self.notebook.add(tab, text=title)
            tab.winfo_toplevel().tk.call('rename', tab, f'!{key}tab')

    def log_to_widget(self, message):
        logger.info(message)

    def show_async_error(self, error):
        messagebox.showerror("Background Task Error", f"An error occurred in a background task:\n\n{error}")

    def on_closing(self):
        if messagebox.askokcancel("Quit", f"Do you want to exit {APP_NAME}?"):
            self.db.close()
            self.destroy()

def center_window(win, w, h):
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    win.geometry(f'{w}x{h}+{int(x)}+{int(y)}')


if __name__ == "__main__":
    if 'python' in sys.executable.lower(): # Basic check to avoid running in PyInstaller bundle analysis
        app = QuoteForgeAIProApp(themename="darkly")
        app.protocol("WM_DELETE_WINDOW", app.on_closing)
        app.mainloop()

# --- PACKAGING HELP ---
# To package this application into a single executable for Windows,
# use PyInstaller with the following command. Ensure you have an
# icon file named 'icon.ico' in an 'assets' subfolder, or remove
# the --icon flag.
#
# pyinstaller --noconfirm --name "QuoteForgeAI" --onefile --windowed --icon "assets\\icon.ico" QuoteForgeAI.py
#
