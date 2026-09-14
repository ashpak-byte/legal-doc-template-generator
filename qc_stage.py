import io
import json
import sys
from pathlib import Path
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import app


def print_qc_row(check_name: str, passed: bool, details: str):
    badge = "[PASS]" if passed else "[FAIL]"
    print(f"{badge} {check_name:38} : {details}")


def run_full_qc() -> bool:
    print("=" * 70)
    print("      LEGAL TEMPLATE GENERATOR - QUALITY CONTROL (QC) STAGE")
    print("=" * 70)

    all_passed = True

    # 1. Config Check
    config = app.load_config()
    cfg_ok = bool(config.get("variables") and config.get("system_prompt"))
    print_qc_row("1. Configuration file (config.json)", cfg_ok, f"Loaded {len(config.get('variables', {}))} variables")
    if not cfg_ok:
        all_passed = False

    # 2. DOCX Engine Check
    sample_path = Path("sample_contract.docx")
    if sample_path.exists():
        doc = Document(str(sample_path))
        plain = app.extract_text_from_docx(doc)
        mod_doc, count = app.generate_template_docx(doc, [("4820-2024", "pt_agreement_number")])
        buf = io.BytesIO()
        mod_doc.save(buf)
        buf.seek(0)
        reloaded = Document(buf)
        docx_ok = count >= 1 and "$$$pt_agreement_number$$$" in app.extract_text_from_docx(reloaded)
        print_qc_row("2. DOCX Extraction & Highlight Engine", docx_ok, f"python-docx text + table replacement OK (count={count})")
        if not docx_ok:
            all_passed = False
    else:
        print_qc_row("2. DOCX Extraction & Highlight Engine", False, "sample_contract.docx not found")
        all_passed = False

    # 3. Gemini API Key Configuration
    api_key = config.get("gemini_api_key", "").strip()
    if not api_key:
        print_qc_row("3. Gemini API Key Configuration", False, "API Key is empty in config.json")
        all_passed = False
        return False
    else:
        print_qc_row("3. Gemini API Key Configuration", True, f"Key present (len={len(api_key)})")

    # 4. Discovery of Supported Models via ModelService.ListModels
    try:
        supported_models = app.fetch_available_models(api_key)
        has_models = len(supported_models) > 0
        top_models = ", ".join(supported_models[:3])
        print_qc_row("4. Model Discovery (ListModels)", has_models, f"Found {len(supported_models)} models ({top_models})")
        if not has_models:
            all_passed = False
    except Exception as e:
        print_qc_row("4. Model Discovery (ListModels)", False, str(e))
        all_passed = False
        return False

    # 5. Live Structured Extraction using selected_model
    chosen_model = config.get("selected_model", supported_models[0] if supported_models else "models/gemini-3.5-flash-lite")
    try:
        sample_doc = "Договір № 4820-2024. Позичальник: Коваленко Петро. Сума: 45000 грн."
        mapping = app.query_gemini_for_variables(
            api_key=api_key,
            model_name=chosen_model,
            system_prompt=config["system_prompt"],
            variables_dict=config["variables"],
            document_text=sample_doc
        )
        has_keys = len(mapping) > 0
        print_qc_row(f"5. End-to-End Extraction ({chosen_model})", has_keys, f"Extracted {len(mapping)} variables successfully")
        if not has_keys:
            all_passed = False
    except Exception as e:
        print_qc_row(f"5. End-to-End Extraction ({chosen_model})", False, str(e))
        all_passed = False

    print("=" * 70)
    if all_passed:
        print("OVERALL QC STATUS: PASSED [READY FOR PRODUCTION]")
    else:
        print("OVERALL QC STATUS: FAILED")
    print("=" * 70)
    return all_passed


if __name__ == "__main__":
    success = run_full_qc()
    sys.exit(0 if success else 1)
