# ⚖️ Legal Document Template Generator (DOCX + Local LLM / Ollama)

A secure, production-grade Python web application built with **Streamlit**, **python-docx**, and an **OpenAI-compatible Local LLM endpoint (Ollama)** for automatically analyzing legal contracts in `.docx` format and generating reusable document templates with highlighted variable placeholders (e.g. `$$$pt_agreement_number$$$`).

---

## 🌟 Architecture & Key Features

| Feature | Description |
|---|---|
| **Local LLM / Ollama Integration** | Migrated from cloud Google Gemini to a self-hosted corporate Local LLM (Ollama) with OpenAI-compatible API (`http://brain.primocollect.ua/v1`). |
| **No External API Keys Required** | The internal corporate endpoint operates securely on the intranet without cloud tokens. Dummy auth (`api_key='ollama'`) is handled automatically. |
| **Dynamic Model Discovery** | Automatically fetches available models from the Ollama server using `client.models.list()`. Resilient to offline/remote environments with manual fallback. |
| **Strict JSON Response Format** | Calls `client.chat.completions.create` with `response_format={"type": "json_object"}` to enforce guaranteed structured JSON parsing. |
| **Full DOCX Architecture** | Built on `python-docx` for reliable handling of paragraphs, tables, runs, and document structure. |
| **Run-Spanning Replacement & Highlighting** | Resolves Word run-fragmentation issues at the paragraph/cell level and applies **yellow highlighting** (`WD_COLOR_INDEX.YELLOW`) to newly inserted `$$$variable_name$$$` tags. |
| **In-Memory Streaming** | Template is compiled directly into an `io.BytesIO()` memory buffer for one-click download of `Template.docx`. |
| **Human-in-the-Loop Review** | Interactive `st.data_editor` with `SelectboxColumn` for reviewing, editing, and reassigning extracted variables. |

---

## 🚀 Getting Started

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Settings

Create or check `config.json` (an example template is provided in `config.example.json`):

```json
{
  "system_prompt": "...",
  "variables": { ... },
  "base_url": "http://brain.primocollect.ua/v1",
  "selected_model": "qwen3.8:27b"
}
```

### 3. Run the Application

```bash
streamlit run app.py
```

Open `http://localhost:8501` in your browser.

---

## 📋 Usage Workflow

1. **⚙️ Settings Tab**:
   - Check the **Local API Base URL** (defaults to `http://brain.primocollect.ua/v1` or `http://10.0.1.17/v1`).
   - Select your model from the dynamically populated list (or type the model name, e.g. `qwen3.8:27b`).
   - Customize the **System Prompt** or **Variables Dictionary** if needed, and click **Save Settings**.
2. **📄 Generation Tab**:
   - Upload your contract or legal document (`.docx`).
   - Click **🤖 Аналізувати документ за допомогою Local LLM**.
   - Review and fine-tune variables in the **Human-in-the-Loop** table.
   - Click **🚀 Generate Template**.
   - Click **📥 Завантажити Template.docx** to download the finalized template with yellow-highlighted placeholders.

---

## 🧪 Testing & Quality Control (QC)

- **Unit Tests** (no live server needed):
  ```bash
  python test_app.py
  ```
- **End-to-End DOCX Pipeline Test**:
  ```bash
  python verify_pipeline.py
  ```
- **Full Quality Control Stage**:
  ```bash
  python qc_stage.py
  ```
