import io
import json
import unittest
from pathlib import Path

import docx
from docx import Document
from docx.enum.text import WD_COLOR_INDEX

import app


class TestLegalTemplateGeneratorOllama(unittest.TestCase):

    def setUp(self):
        self.config = app.load_config()

    def test_config_structure(self):
        self.assertIn("system_prompt", self.config)
        self.assertIn("variables", self.config)
        self.assertIn("base_url", self.config)
        self.assertIn("selected_model", self.config)
        self.assertIn("provider", self.config)
        self.assertIn("gemini_api_key", self.config)
        self.assertIn("CRITICAL RULE", self.config["system_prompt"])
        var_names = app.get_variable_names(self.config["variables"])
        self.assertIn("pt_agreement_number", var_names)
        self.assertIn("pt_first_name", var_names)
        self.assertIn("pt_total_debt", var_names)

    def test_docx_text_extraction(self):
        doc = Document()
        doc.add_paragraph("Договір позики № 987654.")
        table = doc.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "Позичальник:"
        table.cell(0, 1).text = "Шевченко Тарас"

        extracted = app.extract_text_from_docx(doc)
        self.assertIn("987654", extracted)
        self.assertIn("Позичальник:", extracted)
        self.assertIn("Шевченко Тарас", extracted)

    def test_docx_replacement_and_highlight(self):
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("Договір позики № ")
        p.add_run("4820")
        p.add_run("-2024")

        table = doc.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "Сума:"
        table.cell(0, 1).text = "52300 грн."

        replacements = [
            ("4820-2024", "pt_agreement_number"),
            ("52300", "pt_total_debt")
        ]

        mod_doc, count = app.generate_template_docx(doc, replacements)
        self.assertEqual(count, 2)

        # Verify paragraph replacement & highlight
        p_reloaded = mod_doc.paragraphs[0]
        self.assertIn("$$$pt_agreement_number$$$", p_reloaded.text)
        highlighted_runs = [r for r in p_reloaded.runs if r.font.highlight_color == WD_COLOR_INDEX.YELLOW]
        self.assertEqual(len(highlighted_runs), 1)
        self.assertEqual(highlighted_runs[0].text, "$$$pt_agreement_number$$$")

        # Verify table cell replacement & highlight
        cell_p = mod_doc.tables[0].cell(0, 1).paragraphs[0]
        self.assertIn("$$$pt_total_debt$$$", cell_p.text)
        cell_highlighted = [r for r in cell_p.runs if r.font.highlight_color == WD_COLOR_INDEX.YELLOW]
        self.assertEqual(len(cell_highlighted), 1)
        self.assertEqual(cell_highlighted[0].text, "$$$pt_total_debt$$$")

    def test_in_memory_buffer_save(self):
        doc = Document()
        doc.add_paragraph("Тестовий документ")
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        self.assertGreater(len(buf.getvalue()), 0)
        reloaded = Document(buf)
        self.assertEqual(reloaded.paragraphs[0].text, "Тестовий документ")

    def test_json_cleaning(self):
        # 1. Clean markdown code fence with ```json
        raw_response = "```json\n{\n  \"4820-2024\": \"pt_agreement_number\"\n}\n```"
        text = raw_response.strip()
        if text.startswith("```json"):
            text = text[7:]
        elif text.startswith("```"):
            text = text[3:]
        if text.endswith("```"):
            text = text[:-3]
        text = text.strip()
        data = json.loads(text)
        self.assertEqual(data.get("4820-2024"), "pt_agreement_number")

        # 2. Markdown code fence with plain ```
        plain_fence = "```\n{\n  \"52300\": \"pt_total_debt\"\n}\n```"
        text2 = plain_fence.strip()
        if text2.startswith("```json"):
            text2 = text2[7:]
        elif text2.startswith("```"):
            text2 = text2[3:]
        if text2.endswith("```"):
            text2 = text2[:-3]
        text2 = text2.strip()
        data2 = json.loads(text2)
        self.assertEqual(data2.get("52300"), "pt_total_debt")

        # 3. Empty string check
        empty_text = "   "
        self.assertEqual(len(empty_text.strip()), 0)

    def test_strict_json_directive_in_config(self):
        self.assertIn("CRITICAL: You must return ONLY raw, valid JSON", self.config["system_prompt"])

    def test_parse_llm_json_response_function(self):
        # Valid JSON
        res1 = app.parse_llm_json_response('{"test": "pt_pin"}')
        self.assertEqual(res1, {"test": "pt_pin"})

        # Markdown wrapped
        res2 = app.parse_llm_json_response('```json\n{"123": "pt_capital"}\n```')
        self.assertEqual(res2, {"123": "pt_capital"})

        # Empty response error
        with self.assertRaises(ValueError):
            app.parse_llm_json_response("")

        # Malformed response error
        with self.assertRaises(ValueError):
            app.parse_llm_json_response("This is not JSON")

    def test_fetch_available_models_graceful_on_unreachable_endpoint(self):
        # Must return empty list gracefully and not raise uncaught exception
        models_local = app.fetch_local_models("http://127.0.0.1:9999/v1")
        self.assertIsInstance(models_local, list)
        self.assertEqual(len(models_local), 0)

        models_compat = app.fetch_available_models("http://127.0.0.1:9999/v1")
        self.assertIsInstance(models_compat, list)
        self.assertEqual(len(models_compat), 0)

        # Gemini empty key
        models_gemini = app.fetch_gemini_models("")
        self.assertEqual(models_gemini, [])


if __name__ == "__main__":
    unittest.main()
