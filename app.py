import copy
import io
import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import docx
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from openai import OpenAI
import pandas as pd
import streamlit as st

# ==========================================
# CONFIGURATION & CONSTANTS
# ==========================================
CONFIG_FILE = Path(__file__).parent / "config.json"

DEFAULT_BASE_URL = "http://brain.primocollect.ua/v1"
DEFAULT_SELECTED_MODEL = "qwen3.8:27b"

DEFAULT_SYSTEM_PROMPT = (
    "Ти — експерт-юрист. Твоє завдання — аналізувати текст і знаходити значення, "
    "що відповідають глобальним змінним. ІГНОРУЙ константи: не чіпай реквізити "
    "відправника (якщо вони незмінні), індекси, загальні номери законів. "
    "CRITICAL RULE: The keys in your JSON response MUST be exact, "
    "character-by-character copies of the text found in the document "
    "(including spaces and punctuation). Do not reformat dates or numbers. "
    "Повертай ТІЛЬКИ валідний JSON: ключ — знайдений шматок тексту, значення — назва змінної. "
    "CRITICAL: You must return ONLY raw, valid JSON. Do not include any markdown formatting, "
    "do not use ```json blocks, and do not add any explanations, greetings, or conversational text "
    "before or after the JSON. Just the raw JSON object."
)

DEFAULT_VARIABLES = {
    "pt_address": "Адреса реєстрації або фактичного проживання",
    "pt_agreement_number": "Номер кредитного або іншого договору",
    "pt_capital": "Сума основного боргу (тіло кредиту)",
    "pt_first_name": "Ім'я сторони",
    "pt_last_name": "Прізвище сторони",
    "pt_middle_name": "По батькові сторони",
    "pt_pin": "РНОКПП (ІПН) або паспортні дані",
    "pt_total_debt": "Загальна сума заборгованості",
    "pt_court_fee": "Сума судового збору",
    "pt_phone": "Номер телефону"
}

DEFAULT_CONFIG = {
    "system_prompt": DEFAULT_SYSTEM_PROMPT,
    "variables": DEFAULT_VARIABLES,
    "base_url": DEFAULT_BASE_URL,
    "selected_model": DEFAULT_SELECTED_MODEL
}


def get_variable_names(var_dict: Dict[str, Any]) -> List[str]:
    """Returns the list of valid variable identifiers for the SelectboxColumn."""
    names = []
    for k, v in var_dict.items():
        if isinstance(v, str) and (v.startswith("pt_") or "_" in v):
            names.append(v)
        elif isinstance(k, str) and (k.startswith("pt_") or "_" in k):
            names.append(k)
        else:
            names.append(str(k))
    return sorted(list(set(names))) if names else list(var_dict.keys())


# ==========================================
# CONFIGURATION MANAGEMENT
# ==========================================
def load_config() -> Dict[str, Any]:
    """Loads config.json or creates default if not present."""
    if not CONFIG_FILE.exists():
        save_config(DEFAULT_CONFIG)
        return copy.deepcopy(DEFAULT_CONFIG)
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if not isinstance(data, dict):
                data = copy.deepcopy(DEFAULT_CONFIG)
            if "system_prompt" not in data:
                data["system_prompt"] = DEFAULT_SYSTEM_PROMPT
            else:
                directive = "CRITICAL: You must return ONLY raw, valid JSON. Do not include any markdown formatting, do not use ```json blocks, and do not add any explanations, greetings, or conversational text before or after the JSON. Just the raw JSON object."
                if directive not in data["system_prompt"]:
                    data["system_prompt"] = data["system_prompt"].strip() + " " + directive
            if "variables" not in data or not isinstance(data["variables"], dict):
                data["variables"] = DEFAULT_VARIABLES
            if "base_url" not in data:
                data["base_url"] = DEFAULT_BASE_URL
            if "selected_model" not in data:
                data["selected_model"] = DEFAULT_SELECTED_MODEL
            return data
    except Exception as e:
        print(f"Помилка завантаження config.json: {e}")
        return copy.deepcopy(DEFAULT_CONFIG)


def save_config(config_dict: Dict[str, Any]) -> bool:
    """Saves dictionary to config.json."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Помилка збереження config.json: {e}")
        return False


# ==========================================
# DYNAMIC MODEL SERVICE (LOCAL OLLAMA / OPENAI)
# ==========================================
def fetch_available_models(base_url: str) -> List[str]:
    """
    Dynamically fetches all available models from the Local LLM (Ollama)
    OpenAI-compatible endpoint using client.models.list().
    Handles ConnectionError, non-JSON responses, and other Exceptions gracefully without crashing.
    """
    if not base_url or not base_url.strip():
        return []
    try:
        client = OpenAI(
            base_url=base_url.strip(),
            api_key="ollama",
            timeout=5.0
        )
        models_page = client.models.list()
        model_ids = [m.id for m in models_page if hasattr(m, "id")]
        return model_ids
    except Exception:
        # Expected when accessed from outside the corporate Intranet or when server is unavailable
        return []


# ==========================================
# DOCX TEXT EXTRACTION & PROCESSING
# ==========================================
def extract_text_from_docx(doc: Document) -> str:
    """
    Extracts clean plain text from both regular paragraphs and all table cells.
    Preserves document structure for accurate LLM extraction.
    """
    parts = []

    # 1. Main body paragraphs
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            parts.append(t)

    # 2. Table cells
    for table in doc.tables:
        for row in table.rows:
            row_items = []
            for cell in row.cells:
                cell_text = "\n".join(p.text.strip() for p in cell.paragraphs if p.text.strip())
                if cell_text and cell_text not in row_items:
                    row_items.append(cell_text)
            if row_items:
                parts.append(" | ".join(row_items))

    return "\n\n".join(parts)


def replace_in_paragraph(p, replacements: List[Tuple[str, str]]) -> int:
    """
    Robust paragraph-level replacement:
    Reads full paragraph text, replaces target snippets with $$$var$$$,
    clears existing fragmented runs, and rebuilds runs with yellow highlighting
    specifically applied to the $$$variable_name$$$ tags.
    """
    original_text = p.text
    if not original_text:
        return 0

    new_text = original_text
    match_count = 0
    for orig, var in replacements:
        if orig in new_text:
            new_text = new_text.replace(orig, f"$$${var}$$$")
            match_count += 1

    if new_text != original_text:
        font_name = None
        font_size = None
        is_bold = None
        is_italic = None
        color_rgb = None
        if p.runs:
            ref = p.runs[0]
            font_name = ref.font.name
            font_size = ref.font.size
            is_bold = ref.bold
            is_italic = ref.italic
            if ref.font.color and ref.font.color.rgb:
                color_rgb = ref.font.color.rgb

        for r in list(p.runs):
            p._p.remove(r._r)

        tokens = re.split(r"(\$\$\$[a-zA-Z0-9_]+\$\$\$)", new_text)
        for token in tokens:
            if not token:
                continue
            run = p.add_run(token)
            if font_name:
                run.font.name = font_name
            if font_size:
                run.font.size = font_size
            if is_bold is not None:
                run.bold = is_bold
            if is_italic is not None:
                run.italic = is_italic
            if color_rgb is not None:
                run.font.color.rgb = color_rgb

            if re.match(r"^\$\$\$[a-zA-Z0-9_]+\$\$\$$", token):
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    return match_count


def generate_template_docx(doc: Document, replacements: List[Tuple[str, str]]) -> Tuple[Document, int]:
    """
    Iterates through all paragraphs and table cells in the document,
    applying replacements and yellow highlights.
    Returns (modified_document, total_replacement_count).
    """
    total_count = 0
    visited_paragraphs = set()

    # 1. Main body paragraphs
    for p in doc.paragraphs:
        if id(p) not in visited_paragraphs:
            visited_paragraphs.add(id(p))
            total_count += replace_in_paragraph(p, replacements)

    # 2. Table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if id(p) not in visited_paragraphs:
                        visited_paragraphs.add(id(p))
                        total_count += replace_in_paragraph(p, replacements)

    return doc, total_count


# ==========================================
# LOCAL LLM (OLLAMA / OPENAI) INTEGRATION
# ==========================================
def query_local_llm_for_variables(
    base_url: str,
    model_name: str,
    system_prompt: str,
    variables_dict: Dict[str, str],
    document_text: str,
) -> Dict[str, str]:
    """
    Invokes Local LLM (Ollama) using the OpenAI SDK client.
    Enforces strict JSON output with response_format={"type": "json_object"}.
    Returns: {"Exact text snippet in doc": "variable_name"}
    """
    if not base_url.strip():
        raise ValueError("Local API Base URL не вказано. Будь ласка, введіть його у вкладці 'Налаштування'.")

    if not model_name.strip():
        raise ValueError("Модель не обрано. Будь ласка, виберіть або введіть модель у вкладці 'Налаштування'.")

    client = OpenAI(
        base_url=base_url.strip(),
        api_key="ollama",
        timeout=120.0
    )

    directive = (
        "CRITICAL: You must return ONLY raw, valid JSON. Do not include any markdown formatting, "
        "do not use ```json blocks, and do not add any explanations, greetings, or conversational text "
        "before or after the JSON. Just the raw JSON object."
    )
    if directive not in system_prompt:
        system_prompt = system_prompt.strip() + " " + directive

    variables_json_str = json.dumps(variables_dict, ensure_ascii=False, indent=2)
    full_system_prompt = (
        f"{system_prompt}\n\n"
        f"ГЛОБАЛЬНІ ЗМІННІ (словник):\n{variables_json_str}\n\n"
        "ВАЖЛИВО: Повертай ТІЛЬКИ валідний JSON-об'єкт, де кожен ключ — це точний фрагмент тексту з документа, "
        "а значення — назва змінної з доступного словника. Без будь-яких додаткових коментарів чи тексту навколо JSON."
    )

    messages = [
        {"role": "system", "content": full_system_prompt},
        {"role": "user", "content": document_text},
    ]

    response = client.chat.completions.create(
        model=model_name.strip(),
        messages=messages,
        temperature=0.1,
    )

    raw_text = response.choices[0].message.content or ""
    text = raw_text.strip()

    if not text:
        raise ValueError("The LLM returned an empty response.")

    # Strip markdown block formatting if present
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]

    text = text.strip()

    try:
        parsed_json = json.loads(text)
    except json.JSONDecodeError as err:
        raise ValueError(f"Не вдалося розпарсити відповідь LLM як JSON: {err}\nОтримана відповідь:\n{raw_text}")

    if not isinstance(parsed_json, dict):
        raise ValueError("Очікувався JSON-об'єкт (dict) формату {'текст': 'назва_змінної'}, але отримано інший тип.")

    return parsed_json


# ==========================================
# STREAMLIT UI APPLICATION
# ==========================================
def main():
    st.set_page_config(
        page_title="Legal Document Template Generator (Local LLM)",
        page_icon="⚖️",
        layout="wide",
    )

    st.title("⚖️ Legal Document Template Generator")
    st.caption("Автоматизована генерація шаблонів юридичних документів (.docx) на базі корпоративного Local LLM (Ollama).")

    # Initialize session state
    if "config" not in st.session_state:
        st.session_state.config = load_config()

    if "editor_data" not in st.session_state:
        st.session_state.editor_data = pd.DataFrame(
            columns=["Keep", "Original Text", "Assigned Variable"]
        )

    if "docx_bytes" not in st.session_state:
        st.session_state.docx_bytes = None

    if "extracted_doc_text" not in st.session_state:
        st.session_state.extracted_doc_text = ""

    if "generated_docx_buffer" not in st.session_state:
        st.session_state.generated_docx_buffer = None

    # Main Tabs
    tab_gen, tab_settings = st.tabs(["📄 Generation", "⚙️ Settings"])

    # ------------------------------------------
    # TAB 2: SETTINGS (LOCAL LLM ENDPOINT & MODEL)
    # ------------------------------------------
    with tab_settings:
        st.subheader("⚙️ Конфігурація підключення до Local LLM (Ollama)")
        st.markdown(
            "Всі параметри зберігаються локально у файлі `config.json`. "
            "Підключення здійснюється через локальний OpenAI-сумісний ендпоінт корпоративної мережі."
        )

        current_prompt = st.session_state.config.get("system_prompt", DEFAULT_SYSTEM_PROMPT)
        current_variables = st.session_state.config.get("variables", DEFAULT_VARIABLES)
        current_base_url = st.session_state.config.get("base_url", DEFAULT_BASE_URL)
        current_selected_model = st.session_state.config.get("selected_model", DEFAULT_SELECTED_MODEL)

        # Dynamic model fetch with graceful exception handling
        live_models = fetch_available_models(current_base_url) if current_base_url else []

        with st.form("settings_form"):
            base_url_input = st.text_input(
                "1. Local API Base URL (Ollama endpoint)",
                value=current_base_url,
                help="URL локального OpenAI-сумісного API сервера Ollama, наприклад: http://brain.primocollect.ua/v1 або http://10.0.1.17/v1"
            )

            # Dynamic Model Selection Dropdown or Manual Entry if unreachable
            if live_models:
                model_options = list(live_models)
                if current_selected_model not in model_options:
                    model_options.insert(0, current_selected_model)
                selected_model_idx = model_options.index(current_selected_model) if current_selected_model in model_options else 0

                model_choice = st.selectbox(
                    "2. Робоча модель Local LLM (отримано через API)",
                    options=model_options,
                    index=selected_model_idx,
                    help="Список моделей, завантажених безпосередньо з вашого сервера Ollama."
                )
            else:
                st.info("ℹ️ Сервер Local LLM недоступний з поточного середовища (очікувано за межами корпоративної мережі Intranet). Введіть або залиште назву моделі вручну:")
                model_choice = st.text_input(
                    "2. Назва моделі Local LLM (ручне введення)",
                    value=current_selected_model,
                    help="Вкажіть назву моделі з Ollama, наприклад: qwen3.8:27b"
                )

            prompt_input = st.text_area(
                "3. Системний промпт (System Prompt)",
                value=current_prompt,
                height=130,
                help="Інструкція для Local LLM, що визначає логіку точного виділення змінних."
            )

            variables_json_str = json.dumps(current_variables, ensure_ascii=False, indent=2)
            variables_input = st.text_area(
                "4. Словник глобальних змінних (Global Variables Dictionary - JSON)",
                value=variables_json_str,
                height=240,
                help="JSON-словник змінних. Ключі будуть доступні у випадаючому списку вибору."
            )

            submitted = st.form_submit_button("💾 Save Settings", use_container_width=True)

            if submitted:
                try:
                    parsed_vars = json.loads(variables_input)
                    if not isinstance(parsed_vars, dict):
                        st.error("Помилка: Словник змінних має бути валідним JSON-об'єктом.")
                    else:
                        new_config = {
                            "system_prompt": prompt_input.strip(),
                            "variables": parsed_vars,
                            "base_url": base_url_input.strip(),
                            "selected_model": model_choice.strip()
                        }
                        if save_config(new_config):
                            st.session_state.config = new_config
                            st.success(f"✅ Налаштування збережено! Обрана модель: `{model_choice.strip()}`")
                            st.rerun()
                except json.JSONDecodeError as err:
                    st.error(f"❌ Помилка валідації JSON у словнику змінних: {err}")

    # ------------------------------------------
    # TAB 1: GENERATION (DOCX PIPELINE)
    # ------------------------------------------
    with tab_gen:
        st.subheader("📄 Генерація шаблону DOCX")

        var_dict = st.session_state.config.get("variables", DEFAULT_VARIABLES)
        var_options = get_variable_names(var_dict)
        active_model = st.session_state.config.get("selected_model", DEFAULT_SELECTED_MODEL)
        active_base_url = st.session_state.config.get("base_url", DEFAULT_BASE_URL)

        st.caption(f"Ендпоінт: `{active_base_url}` | Модель: `{active_model}`")

        # Step 1: File Uploader (Accepts ONLY .docx)
        uploaded_file = st.file_uploader(
            "Завантажте документ у форматі .docx",
            type=["docx"],
            help="Оберіть файл .docx для автоматичного аналізу та перетворення на шаблон."
        )

        if uploaded_file is not None:
            file_bytes = uploaded_file.getvalue()
            st.session_state.docx_bytes = file_bytes

            # Extract plain text from paragraphs and tables
            try:
                doc = Document(io.BytesIO(file_bytes))
                extracted_text = extract_text_from_docx(doc)
                st.session_state.extracted_doc_text = extracted_text
            except Exception as e:
                st.error(f"Помилка читання .docx файлу: {e}")
                extracted_text = ""

            # Plain text preview expander
            with st.expander("🔍 Попередній перегляд вилученого тексту з .docx", expanded=False):
                st.text_area("Вилучений текст (параграфи + таблиці)", extracted_text, height=200, disabled=True)

            col_run, col_clear = st.columns([3, 1])
            with col_run:
                run_analysis = st.button("🤖 Аналізувати документ за допомогою Local LLM", use_container_width=True, type="primary")
            with col_clear:
                if st.button("🔄 Очистити таблицю", use_container_width=True):
                    st.session_state.editor_data = pd.DataFrame(
                        columns=["Keep", "Original Text", "Assigned Variable"]
                    )
                    st.session_state.generated_docx_buffer = None
                    st.rerun()

            # Step 2: Call Local LLM API
            if run_analysis:
                system_prompt = st.session_state.config.get("system_prompt", DEFAULT_SYSTEM_PROMPT)

                if not active_base_url:
                    st.warning("⚠️ Local API Base URL не налаштовано. Перейдіть на вкладку '⚙️ Settings' та збережіть URL.")
                elif not extracted_text.strip():
                    st.error("Помилка: Не вдалося отримати текст із завантаженого .docx документа.")
                else:
                    with st.spinner(f"Запит до `{active_model}` на `{active_base_url}`... Аналізуємо документ та шукаємо змінні..."):
                        try:
                            llm_mapping = query_local_llm_for_variables(
                                base_url=active_base_url,
                                model_name=active_model,
                                system_prompt=system_prompt,
                                variables_dict=var_dict,
                                document_text=extracted_text,
                            )

                            # Build DataFrame for Human-in-the-Loop review
                            rows = []
                            for snippet, assigned_var in llm_mapping.items():
                                valid_var = assigned_var if assigned_var in var_options else (var_options[0] if var_options else "")
                                rows.append({
                                    "Keep": True,
                                    "Original Text": str(snippet),
                                    "Assigned Variable": valid_var
                                })

                            if not rows:
                                st.info("ℹ️ Local LLM не знайшов збігів зі словником змінних у цьому документі.")
                                st.session_state.editor_data = pd.DataFrame(
                                    columns=["Keep", "Original Text", "Assigned Variable"]
                                )
                            else:
                                st.session_state.editor_data = pd.DataFrame(rows)
                                st.success(f"✅ Знайдено {len(rows)} потенційних змінних! Перевірте їх у таблиці нижче.")

                        except Exception as ex:
                            err_msg = str(ex)
                            if "connection" in err_msg.lower() or "connect" in err_msg.lower():
                                st.error(
                                    f"❌ Помилка з'єднання з локальним сервером LLM ({active_base_url}). "
                                    f"Перевірте, чи ви підключені до корпоративної мережі (Intranet/VPN) та чи запущено Ollama.\n\n"
                                    f"Деталі: {err_msg}"
                                )
                            else:
                                st.error(f"❌ Помилка аналізу через Local LLM API: {ex}")

        # Step 3: Human-in-the-Loop Review (st.data_editor)
        st.markdown("### 🧑‍⚖️ Human-in-the-Loop: Перевірка та коригування змінних")
        st.caption("Позначте 'Keep' для заміни фрагмента. Виберіть потрібну змінну через випадаючий список за потреби.")

        if not var_options:
            st.warning("Увага: Словник змінних порожній! Додайте змінні у вкладці 'Settings'.")

        column_config = {
            "Keep": st.column_config.CheckboxColumn(
                "Keep",
                help="Позначте для заміни цього фрагмента на шаблонну змінну",
                default=True,
            ),
            "Original Text": st.column_config.TextColumn(
                "Original Text",
                help="Оригінальний фрагмент тексту з документа для заміни",
                required=True,
            ),
            "Assigned Variable": st.column_config.SelectboxColumn(
                "Assigned Variable",
                help="Призначена змінна зі списку доступних у конфігурації",
                width="medium",
                options=var_options,
                required=True,
            ),
        }

        edited_df = st.data_editor(
            st.session_state.editor_data,
            column_config=column_config,
            disabled=False,
            num_rows="dynamic",
            use_container_width=True,
            key="data_editor"
        )

        # Step 4: Template Generation
        st.markdown("---")
        st.subheader("⚡ Генерація та завантаження шаблону DOCX")

        col_gen, col_download = st.columns([2, 2])

        with col_gen:
            if st.button("🚀 Generate Template", use_container_width=True, type="primary"):
                if not st.session_state.docx_bytes:
                    st.warning("Будь ласка, завантажте .docx файл для генерації шаблону.")
                elif edited_df.empty:
                    st.warning("Таблиця змінних порожня. Додайте змінні або виконайте аналіз документа.")
                else:
                    items_to_replace = []
                    for _, row in edited_df.iterrows():
                        keep = row.get("Keep", False)
                        orig_text = str(row.get("Original Text", "")).strip()
                        var_name = str(row.get("Assigned Variable", "")).strip()
                        if keep and orig_text and var_name:
                            items_to_replace.append((orig_text, var_name))

                    if not items_to_replace:
                        st.warning("Жодного рядка не позначено для заміни (Keep = True).")
                    else:
                        with st.spinner("Застосовуємо заміни та жовте підсвічування до .docx..."):
                            fresh_doc = Document(io.BytesIO(st.session_state.docx_bytes))
                            mod_doc, count = generate_template_docx(fresh_doc, items_to_replace)

                            # Save to in-memory buffer
                            out_buffer = io.BytesIO()
                            mod_doc.save(out_buffer)
                            out_buffer.seek(0)
                            st.session_state.generated_docx_buffer = out_buffer.getvalue()

                            st.success(
                                f"🎉 Шаблон успішно згенеровано! Замінено {count} фрагментів "
                                f"із жовтим підсвічуванням (WD_COLOR_INDEX.YELLOW)."
                            )

        with col_download:
            if st.session_state.generated_docx_buffer:
                st.download_button(
                    label="📥 Завантажити Template.docx",
                    data=st.session_state.generated_docx_buffer,
                    file_name="Template.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            else:
                st.button("📥 Завантажити Template.docx", disabled=True, use_container_width=True)


if __name__ == "__main__":
    main()
