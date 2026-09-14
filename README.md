# ⚖️ Legal Document Template Generator (DOCX + Gemini)

A secure, production-grade Python web application built with **Streamlit**, **python-docx**, and **Google Gemini** for automatically analyzing legal contracts in `.docx` format and generating reusable document templates with highlighted variable placeholders (e.g. `$$$pt_agreement_number$$$`).

---

## 🌟 Key Architecture & Upgrades

| Feature | Description |
|---|---|
| **Full DOCX Architecture** | Complete migration from RTF to DOCX. Uses `python-docx` for reliable handling of paragraphs, tables, runs, and formatting. |
| **Dynamic Model Selection** | Uses `google.generativeai.list_models()` to dynamically discover all available models that support `generateContent`. Zero hardcoded deprecated model names. |
| **User Model Choice** | Interactive `st.selectbox` in Settings allows choosing the active model (e.g. `models/gemini-3.5-flash-lite`), saved to `config.json` as `selected_model`. |
| **Exact Character Matching** | System prompt enforces exact character-by-character snippet copying without date or number reformatting. |
| **Paragraph & Table Run Replacement** | Robust run-splitting engine that preserves styling, clears fragmented runs, and specifically applies **yellow highlighting** (`WD_COLOR_INDEX.YELLOW`) to variable tags. |
| **In-Memory Streaming** | DOCX template is rendered directly into an `io.BytesIO()` memory buffer for one-click downloading. |
| **Security & Privacy** | `config.json` is excluded via `.gitignore` to keep API keys private. A `config.example.json` template is provided. |

---

## 🚀 Getting Started

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Settings

Create `config.json` from the example (or let the app auto-create it):

```bash
cp config.example.json config.json
```

### 3. Run the Application

```bash
streamlit run app.py
```

The application will open in your browser at `http://localhost:8501`.

---

## 📋 Workflow

1. **⚙️ Settings Tab**:
   - Enter your **Gemini API Key**.
   - Select your preferred model from the dynamically populated **Робоча модель Gemini** dropdown (e.g., `models/gemini-3.5-flash-lite`).
   - Adjust the **System Prompt** or **Variables Dictionary** if desired, and click **Save Settings**.
2. **📄 Generation Tab**:
   - Upload your legal document (`.docx`).
   - Click **🤖 Аналізувати документ за допомогою Gemini**.
   - Review extracted variables in the **Human-in-the-Loop** editor: toggle `Keep`, adjust snippet text, or change assigned variables via the dropdown.
   - Click **🚀 Generate Template**.
   - Click **📥 Завантажити Template.docx** to download your template with highlighted variables.

---

## 🧪 Testing & Quality Control (QC)

- **Unit Tests**:
  ```bash
  python test_app.py
  ```
- **End-to-End DOCX Pipeline Test**:
  ```bash
  python verify_pipeline.py
  ```
- **Full Quality Control (QC) Stage**:
  ```bash
  python qc_stage.py
  ```
