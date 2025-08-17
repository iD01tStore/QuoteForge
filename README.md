# QuoteForgeAI Pro

**Premium AI-Powered Quote, Contract, and Soumission Reader for Construction & Services**
**Version 1.0.0**
Created by **iD01t Productions**

---

## 🚀 Overview

**QuoteForgeAI Pro** is an all-in-one desktop application that transforms complex construction *soumissions* and multi-page contracts into clean, actionable quotes and professional agreements.

Built for contractors, estimators, and service providers in Quebec, Canada, and worldwide, it automates:

* **Parsing of big construction bids and contracts** (PDF, DOCX, TXT)
* **Automatic extraction of quantities, prices, timelines, legal clauses**
* **AI-generated professional quotes & contracts**
* **Export to branded PDFs with Quebec/Canada tax rules**
* **Direct email integration to clients**

Unlike subscription tools, **this is one single full version**: pay once, own it forever.

---

## ✨ Key Features

* **Soumission Reader for Construction**

  * AI extraction of scope, line items, taxes, delivery dates, bonding, warranty, exclusions, addenda, and more
  * Built for Quebec/Canadian compliance
  * Side-by-side review with confidence scoring

* **Quote Builder**

  * Add or AI-generate line items
  * Auto totals, GST/HST/QST
  * Export polished PDF with branding

* **Contract Builder**

  * Industry templates (Construction, Electrical, Consulting, Design, Marketing)
  * AI Enhance for clearer, more protective language
  * Compliance checker, clause generator, PDF export

* **Client Manager**

  * Local SQLite CRM with quick-add clients
  * Auto-link clients to quotes and contracts

* **Email Integration**

  * Send quotes and contracts directly with attachments via SMTP
  * Works with Gmail, Outlook, or custom mail servers

* **Templates & Branding**

  * Built-in templates with category filter
  * Customizable company details, logo, taxes, and terms

* **Local-first**

  * Your data stays local in `quoteforge_data.db`
  * No subscription cloud database required

---

## 🧠 AI Power: API Key Options

QuoteForgeAI Pro supports **all major AI providers**. Add your keys once in `Settings` and the app will route calls.

* **Google Gemini** → `GEMINI_API_KEY` (preferred for contract & soumission parsing)
* **OpenAI** → `OPENAI_API_KEY`
* **Anthropic Claude** → `ANTHROPIC_API_KEY`
* **Cohere** → `COHERE_API_KEY`
* **Mistral** → `MISTRAL_API_KEY`
* **AWS Bedrock** → `AWS_ACCESS_KEY_ID`, `AWS_SECRET_ACCESS_KEY`, `AWS_REGION`
* **Azure OpenAI** → `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_KEY`, `AZURE_OPENAI_DEPLOYMENT`
* **Groq** → `GROQ_API_KEY`
* **NVIDIA NIM** → `NVIDIA_API_KEY`
* **Ollama (local LLMs)** → `OLLAMA_HOST` (default `http://localhost:11434`)

Environment variables override config file values.

---

## 📦 Installation

**Windows (Python 3.11 recommended):**

```bash
# 1. Create environment
python -m venv .venv
.venv\Scripts\activate

# 2. Upgrade pip
python -m pip install --upgrade pip

# 3. Install requirements
pip install -r requirements.txt
```

**Build an .exe with icon:**

```bash
pyinstaller --noconfirm ^
  --name "QuoteForgeAI" ^
  --onefile ^
  --windowed ^
  --icon "assets\\icon.ico" ^
  main.py
```

---

## ▶️ Running

```bash
python main.py
```

The app opens with Dashboard + tabs for **Quote Builder, Contract Builder, Client Manager, Templates, Settings**.

---

## ⚙️ Configuration

The app auto-creates `config.json`.

Example:

```json
{
  "language": "en",
  "currency": "CAD",
  "tax_rate": 0.15,
  "company_name": "Your Company Inc.",
  "company_address": "123 Rue, Montreal, QC",
  "company_phone": "+1 514 000 0000",
  "company_email": "info@yourcompany.ca",

  "email_smtp_server": "smtp.gmail.com",
  "email_smtp_port": 587,
  "email_username": "you@gmail.com",
  "email_password": "app_password",

  "gemini_api_key": "",
  "openai_api_key": "",
  "anthropic_api_key": "",
  "cohere_api_key": "",
  "mistral_api_key": "",
  "aws_access_key_id": "",
  "aws_secret_access_key": "",
  "aws_region": "ca-central-1",
  "azure_openai_endpoint": "",
  "azure_openai_key": "",
  "azure_openai_deployment": "",
  "groq_api_key": "",
  "nvidia_api_key": "",
  "ollama_host": "http://localhost:11434"
}
```

---

## 📊 Workflow Example

1. **Import a soumission PDF**
   → Contract Builder → *Analyze Soumission* with Gemini
   → Extract items, totals, clauses

2. **Push items into Quote Builder**
   → Review and adjust
   → Auto-taxes calculated

3. **Enhance contract with AI**
   → Add clauses (Force Majeure, Indemnification, Governing Law)

4. **Export & Send**
   → PDF out with company branding
   → Email directly to client

---

## 🔐 Security & Privacy

* Local SQLite storage
* API keys stored in config or env vars
* Only text you choose is sent to AI providers
* Use Ollama for fully local private runs

---

## 💵 Licensing & Pricing

* **One-time fixed license**
* No subscription, no upsells
* Includes all features (quotes, contracts, soumission reader, AI integrations, PDF export, email)
* Future updates delivered as minor patches

---

## 📈 Roadmap

* Automatic GST/QST split for Quebec
* E-signature integration
* Multi-currency with FX lookup
* Industry-specific legal packs (construction, electrical, civil, HVAC)

---

## 📣 Final Pitch

QuoteForgeAI Pro is the **only fixed-price AI tool** designed for contractors and businesses who are tired of wasting hours on paperwork. Import your tender, extract the key numbers, build a professional quote, check legal compliance, export a branded PDF, and send—all in one place.

**Pay once, use forever.**
**Faster quotes. Cleaner contracts. More wins.**
**299$CAD**
