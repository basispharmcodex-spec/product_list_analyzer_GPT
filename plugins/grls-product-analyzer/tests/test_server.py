import importlib.util
import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


PLUGIN_ROOT = Path(__file__).resolve().parents[1]
SERVER_PATH = PLUGIN_ROOT / "mcp-server" / "server.py"

spec = importlib.util.spec_from_file_location("grls_mcp_server", SERVER_PATH)
server = importlib.util.module_from_spec(spec)
assert spec and spec.loader
spec.loader.exec_module(server)


def make_xlsx(path: Path, rows: list[list[str]], sheet_name: str = "Действующий") -> None:
    shared = []
    shared_index = {}

    def shared_id(value: str) -> int:
        if value not in shared_index:
            shared_index[value] = len(shared)
            shared.append(value)
        return shared_index[value]

    xml_rows = []
    for row_number, values in enumerate(rows, start=1):
        cells = []
        for column_number, value in enumerate(values, start=1):
            if value == "":
                continue
            number = column_number
            letters = ""
            while number:
                number, remainder = divmod(number - 1, 26)
                letters = chr(65 + remainder) + letters
            cells.append(f'<c r="{letters}{row_number}" t="s"><v>{shared_id(value)}</v></c>')
        xml_rows.append(f'<row r="{row_number}">{"".join(cells)}</row>')

    shared_xml = "".join(f"<si><t>{value}</t></si>" for value in shared)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>""",
        )
        archive.writestr(
            "xl/workbook.xml",
            f"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""",
        )
        archive.writestr(
            "xl/sharedStrings.xml",
            f"""<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared)}" uniqueCount="{len(shared)}">{shared_xml}</sst>""",
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            f"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:Q{len(rows)}"/>
  <sheetData>{''.join(xml_rows)}</sheetData>
</worksheet>""",
        )


def make_text_pdf(path: Path, lines: list[str]) -> None:
    from pypdf import PdfWriter
    from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject

    writer = PdfWriter()
    page = writer.add_blank_page(width=612, height=792)
    font = DictionaryObject(
        {
            NameObject("/Type"): NameObject("/Font"),
            NameObject("/Subtype"): NameObject("/Type1"),
            NameObject("/BaseFont"): NameObject("/Helvetica"),
        }
    )
    font_reference = writer._add_object(font)
    page[NameObject("/Resources")] = DictionaryObject(
        {NameObject("/Font"): DictionaryObject({NameObject("/F1"): font_reference})}
    )
    escaped = [
        line.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
        for line in lines
    ]
    commands = ["BT", "/F1 12 Tf", "72 720 Td"]
    for index, line in enumerate(escaped):
        if index:
            commands.append("0 -20 Td")
        commands.append(f"({line}) Tj")
    commands.append("ET")
    stream = DecodedStreamObject()
    stream.set_data("\n".join(commands).encode("latin-1"))
    page[NameObject("/Contents")] = writer._add_object(stream)
    with path.open("wb") as output:
        writer.write(output)


def grls_fixture_rows() -> list[list[str]]:
    rows = [[""] * 17 for _ in range(10)]
    rows[4][2:] = [
        "Номер регистрационного удостоверения",
        "Дата регистрации",
        "Дата окончания действия регистрационного удостоверения",
        "Дата аннулирования",
        "Держатель РУ",
        "Страна",
        "Торговое наименование",
        "МНН",
        "Формы выпуска",
        "Стадии производства",
        "Нормативная документация",
        "Фармакотерапевтическая группа",
        "ЖНВЛП",
        "Контролируемые вещества",
        "Орфанный",
    ]
    rows[5][2] = "Действующий"
    rows[6][2:] = [
        "ЛП-TEST-001",
        "01.01.2026",
        "",
        "",
        "ООО Тест",
        "Россия",
        "Тестовир",
        "Ацикловир",
        "таблетки, 200 мг",
        "Все стадии",
        "НД-001",
        "противовирусное средство",
        "Да",
        "Нет",
        "Нет",
    ]
    rows[7][2:] = [
        "ЛП-TEST-002",
        "02.01.2026",
        "",
        "",
        "ООО Другое",
        "Россия",
        "Валатест",
        "Валацикловир",
        "таблетки, 500 мг",
        "Все стадии",
        "НД-002",
        "противовирусное средство",
        "Нет",
        "Нет",
        "Нет",
    ]
    rows[8][2:] = [
        "ЛП-OLD-001",
        "01.01.2020",
        "",
        "",
        "ООО Тест",
        "Россия",
        "Тестовир",
        "Ацикловир",
        "таблетки, 200 мг",
        "Все стадии",
        "НД-OLD",
        "противовирусное средство",
        "Да",
        "Нет",
        "Нет",
    ]
    rows[9][2:] = [
        "ФС-TEST-003",
        "03.01.2026",
        "",
        "",
        "ООО Субстанция",
        "Россия",
        "Ацикловир",
        "Ацикловир",
        "субстанция-порошок, 25 кг",
        "Все стадии",
        "ФС-003",
        "",
        "Нет",
        "Нет",
        "Нет",
    ]
    return rows


class ServerTests(unittest.TestCase):
    def test_bundled_grls_status(self):
        status = server.get_grls_status({})
        self.assertTrue(Path(status["active_archive_path"]).is_file())
        self.assertEqual(status["archive_as_of"], "2026-09-21")
        self.assertEqual(status["stale_after_days"], 30)

    def test_csv_product_list_preview(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "products.csv"
            path.write_text("id,product\n1,Acyclovir Tablets\n", encoding="utf-8")
            result = server.inspect_product_list({"file_path": str(path), "max_rows": 10})
            self.assertEqual(result["last_row"], 2)
            self.assertEqual(result["rows"][1]["values"][1], "Acyclovir Tablets")

    def test_pdf_product_list_preview(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "products.pdf"
            make_text_pdf(path, ["Product list", "Paracetamol Tablets 500 mg"])
            result = server.inspect_product_list(
                {"file_path": str(path), "max_rows": 10, "max_pages": 5}
            )
            self.assertEqual(result["format"], "pdf")
            self.assertEqual(result["total_pages"], 1)
            self.assertEqual(result["rows"][1]["page"], 1)
            self.assertIn("Paracetamol Tablets 500 mg", result["rows"][1]["values"][0])

    def test_literal_search_and_report(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            xlsx_path = directory_path / "grls-Действующий.xlsx"
            make_xlsx(xlsx_path, grls_fixture_rows())
            eaeu_path = directory_path / "grls-Выдано_по_правилам_ЕАЭС.xlsx"
            eaeu_rows = grls_fixture_rows()
            eaeu_rows[5][2] = "Выдано по правилам ЕАЭС"
            make_xlsx(eaeu_path, eaeu_rows)
            archive_path = directory_path / "grls.zip"
            with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.write(xlsx_path, xlsx_path.name)
                archive.write(eaeu_path, eaeu_path.name)

            exact = server.search_grls(
                {
                    "archive_path": str(archive_path),
                    "queries": ["Ацикловир"],
                    "fields": ["inn"],
                    "match_mode": "exact",
                }
            )
            self.assertEqual(exact["returned"], 3)
            self.assertEqual(exact["records"][0]["registration_number"], "ЛП-TEST-001")
            self.assertEqual({record["source_status"] for record in exact["records"]}, {"Действующий"})

            contains = server.search_grls(
                {
                    "archive_path": str(archive_path),
                    "queries": ["ацикловир"],
                    "fields": ["inn"],
                    "match_mode": "contains",
                }
            )
            self.assertEqual(contains["returned"], 4)

            report_path = directory_path / "report.html"
            build = server.build_html_report(
                {
                    "archive_path": str(archive_path),
                    "output_path": str(report_path),
                    "decisions": [
                        {
                            "source_row": 1,
                            "serial_number": "1",
                            "product": "Acyclovir Tablets",
                            "dosage": "200 mg",
                            "packing": "10 tablets",
                            "segment": "Antiviral",
                            "normalized_inn": "Ацикловир",
                            "accepted_registry_inns": ["Ацикловир"],
                            "desired_form": "таблетки",
                            "form_terms": ["таблет"],
                            "match_note": "Проверенное прямое совпадение.",
                            "confidence": "Высокая",
                            "decision_status": "matched",
                        },
                        {
                            "source_row": 2,
                            "serial_number": "2",
                            "product": "Acyclovir",
                            "dosage": "",
                            "packing": "",
                            "segment": "Antiviral",
                            "normalized_inn": "Ацикловир",
                            "accepted_registry_inns": ["Ацикловир"],
                            "desired_form": "",
                            "form_terms": [],
                            "match_note": "Форма в Product List не указана.",
                            "confidence": "Высокая",
                            "decision_status": "matched",
                        }
                    ],
                    "title": "Тестовый отчёт",
                    "grls_as_of": "01.01.2026",
                }
            )
            self.assertEqual(build["unique_registration_numbers"], 3)
            validation = server.validate_html_report({"report_path": str(report_path)})
            self.assertTrue(validation["valid"], validation["errors"])
            report_html = report_path.read_text(encoding="utf-8")
            self.assertIn('id="helpTooltip"', report_html)
            self.assertIn("позиций Product List", report_html)
            self.assertIn(
                "Основные показатели — «Количество конкурентных препаратов на рынке» и «Производители на рынке».",
                report_html,
            )
            self.assertIn("ЛП-TEST-001", report_html)
            self.assertIn("ЛП-OLD-001", report_html)
            self.assertIn("ФС-TEST-003", report_html)
            self.assertNotIn('id="manufacturerFilter"', report_html)
            self.assertIn('id="formFilter"', report_html)
            self.assertIn('id="typeFilter"', report_html)
            self.assertIn('id="vedFilter"', report_html)
            self.assertIn('data-key="market_manufacturer_count"', report_html)
            self.assertIn('id="summaryTopScroll"', report_html)
            self.assertIn('id="summaryTableScroll"', report_html)
            self.assertIn("setupSummaryScrollSync", report_html)
            self.assertNotIn("Прокрутка таблицы влево / вправо", report_html)
            self.assertIn("row.addEventListener('pointerdown'", report_html)
            self.assertIn("if (dragged) return;", report_html)
            self.assertIn('data-primary-metric="medicine"', report_html)
            self.assertIn("const hasDesiredForm", report_html)
            self.assertIn('data-key="expiration_date"', report_html)
            self.assertNotIn('data-key="trade_name_count"', report_html)
            self.assertNotIn('data-key="confidence"', report_html)
            self.assertNotIn('data-key="reregistration"', report_html)
            self.assertIn("['reregistration','Перерегистрация']", report_html)
            self.assertIn("'таблетки'", report_html)
            self.assertIn("'мазь'", report_html)
            self.assertNotIn('id="statusFilter"', report_html)
            self.assertNotIn('data-key="source_status"', report_html)
            self.assertNotIn("Что включено", report_html)
            self.assertNotIn("Контроль сопоставления", report_html)

            product_data = re.search(
                r'<script id="productsData" type="application/json">(.*?)</script>',
                report_html,
                re.DOTALL,
            )
            self.assertIsNotNone(product_data)
            summaries = json.loads(product_data.group(1))
            summary = summaries[0]
            self.assertEqual(summary["registration_count"], 3)
            self.assertEqual(summary["medicine_count"], 2)
            self.assertEqual(summary["substance_count"], 1)
            self.assertEqual(summary["form_registration_count"], 2)
            self.assertEqual(summary["market_manufacturer_count"], 1)
            fallback_summary = summaries[1]
            self.assertEqual(fallback_summary["desired_form"], "")
            self.assertEqual(fallback_summary["medicine_count"], 2)
            self.assertEqual(fallback_summary["form_registration_count"], 0)

    def test_update_archive_and_use_it_by_default(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            data_path = directory_path / "plugin-data"
            xlsx_path = directory_path / "grls-Действующий.xlsx"
            make_xlsx(xlsx_path, grls_fixture_rows())
            archive_path = directory_path / "grls2026-09-22.zip"
            with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.write(xlsx_path, xlsx_path.name)

            previous = os.environ.get("GRLS_ANALYZER_DATA_DIR")
            os.environ["GRLS_ANALYZER_DATA_DIR"] = str(data_path)
            try:
                update = server.update_grls_archive({"archive_path": str(archive_path)})
                self.assertTrue(update["updated"])
                self.assertEqual(update["source"], "user_update")
                self.assertEqual(update["archive_as_of"], "2026-09-22")

                exact = server.search_grls(
                    {
                        "queries": ["Ацикловир"],
                        "fields": ["inn"],
                        "match_mode": "exact",
                    }
                )
                self.assertEqual(exact["returned"], 3)
                self.assertTrue(Path(exact["archive_path"]).samefile(data_path / "grls-current.zip"))
            finally:
                if previous is None:
                    os.environ.pop("GRLS_ANALYZER_DATA_DIR", None)
                else:
                    os.environ["GRLS_ANALYZER_DATA_DIR"] = previous

    def test_report_deduplicates_same_registration_number_across_inns(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            rows = grls_fixture_rows()
            rows[7][2] = "ЛП-TEST-001"
            xlsx_path = directory_path / "grls-Действующий.xlsx"
            make_xlsx(xlsx_path, rows)
            archive_path = directory_path / "grls.zip"
            with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.write(xlsx_path, xlsx_path.name)

            report_path = directory_path / "deduplicated.html"
            decisions = []
            for source_row, inn in ((1, "Ацикловир"), (2, "Валацикловир")):
                decisions.append(
                    {
                        "source_row": source_row,
                        "product": inn,
                        "dosage": "",
                        "packing": "",
                        "segment": "",
                        "normalized_inn": inn,
                        "accepted_registry_inns": [inn],
                        "desired_form": "таблетки",
                        "form_terms": ["таблет"],
                        "match_note": "Проверенное точное совпадение.",
                        "confidence": "Высокая",
                        "decision_status": "matched",
                    }
                )
            build = server.build_html_report(
                {
                    "archive_path": str(archive_path),
                    "output_path": str(report_path),
                    "decisions": decisions,
                }
            )
            self.assertEqual(build["source_rows"], 3)
            self.assertEqual(build["unique_registration_numbers"], 3)
            validation = server.validate_html_report({"report_path": str(report_path)})
            self.assertTrue(validation["valid"], validation["errors"])

    def test_monthly_update_reminder(self):
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            data_path = directory_path / "plugin-data"
            xlsx_path = directory_path / "grls-Действующий.xlsx"
            make_xlsx(xlsx_path, grls_fixture_rows())
            archive_path = directory_path / "grls2000-01-01.zip"
            with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.write(xlsx_path, xlsx_path.name)

            previous = os.environ.get("GRLS_ANALYZER_DATA_DIR")
            os.environ["GRLS_ANALYZER_DATA_DIR"] = str(data_path)
            try:
                server.update_grls_archive({"archive_path": str(archive_path)})
                status = server.get_grls_status({})
                self.assertTrue(status["requires_update"])
                self.assertGreaterEqual(status["age_days"], 30)
                self.assertIn("прикрепить", status["reminder"])
            finally:
                if previous is None:
                    os.environ.pop("GRLS_ANALYZER_DATA_DIR", None)
                else:
                    os.environ["GRLS_ANALYZER_DATA_DIR"] = previous

    def test_stdio_protocol_lists_tools(self):
        request = json.dumps(
            {"jsonrpc": "2.0", "id": 1, "method": "tools/list", "params": {}}
        )
        completed = subprocess.run(
            [sys.executable, str(SERVER_PATH)],
            input=request + "\n",
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=30,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        response = json.loads(completed.stdout.strip())
        names = {tool["name"] for tool in response["result"]["tools"]}
        self.assertIn("get_grls_status", names)
        self.assertIn("update_grls_archive", names)
        self.assertIn("inspect_product_list", names)
        self.assertIn("build_html_report", names)


if __name__ == "__main__":
    unittest.main()
