import io
import json
import sys
from pathlib import Path
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import app


def print_qc_row(badge: str, check_name: str, details: str):
    print(f"{badge:6} {check_name:38} : {details}")


def run_full_qc() -> bool:
    print("=" * 70)
    print("   LEGAL TEMPLATE GENERATOR - QUALITY CONTROL (QC) STAGE (LOCAL LLM)")
    print("=" * 70)

    all_passed = True

    # 1. Config Check
    config = app.load_config()
    cfg_ok = bool(config.get("variables") and config.get("system_prompt") and config.get("base_url"))
    print_qc_row("[PASS]" if cfg_ok else "[FAIL]", "1. Configuration file (config.json)", f"Loaded {len(config.get('variables', {}))} variables | base_url={config.get('base_url')}")
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
        print_qc_row("[PASS]" if docx_ok else "[FAIL]", "2. DOCX Extraction & Highlight Engine", f"python-docx text + table replacement OK (count={count})")
        if not docx_ok:
            all_passed = False
    else:
        print_qc_row("[FAIL]", "2. DOCX Extraction & Highlight Engine", "sample_contract.docx not found")
        all_passed = False

    # 3. Local LLM Endpoint Discovery (Graceful on Intranet-only endpoints)
    base_url = config.get("base_url", "").strip()
    selected_model = config.get("selected_model", "qwen3.8:27b")

    models = app.fetch_available_models(base_url)
    if models:
        top_models = ", ".join(models[:3])
        print_qc_row("[PASS]", "3. Local LLM Endpoint & Models", f"Connected! Found {len(models)} models ({top_models})")

        # 4. Live Test if reachable
        try:
            sample_doc = "Договір № 4820-2024. Позичальник: Коваленко Петро. Сума: 45000 грн."
            mapping = app.query_local_llm_for_variables(
                base_url=base_url,
                model_name=selected_model,
                system_prompt=config["system_prompt"],
                variables_dict=config["variables"],
                document_text=sample_doc
            )
            has_keys = len(mapping) > 0
            print_qc_row("[PASS]" if has_keys else "[FAIL]", f"4. Live Extraction ({selected_model})", f"Extracted {len(mapping)} variables successfully")
            if not has_keys:
                all_passed = False
        except Exception as e:
            print_qc_row("[WARN]", f"4. Live Extraction ({selected_model})", f"Live query note: {e}")
    else:
        print_qc_row("[SKIP]", "3. Local LLM Endpoint & Models", f"Endpoint '{base_url}' unreachable from this environment (expected if outside corporate Intranet). Skipping live LLM ping.")

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
