import io
import unittest
import openpyxl
from backend.excel.builder import build_workbook


class TestExcelBuilder(unittest.TestCase):
    def setUp(self):
        self.sample_file_results = [
            {
                "fileName": "Setor21.pdf",
                "tabName": "Principal",
                "tables": [
                    {
                        "tableName": "Setor 21",
                        "headers": ["Processo", "Tipo de Solicitação", "Situação", "Ano", "Valor"],
                        "rows": [
                            ["0001", "Alvará", "Deferido", 2026, 150.0],
                            ["0002", "Licença", "Em Análise", 2026, 250.0],
                            ["0003", "Alvará", "Indeferido", 2026, 100.0],
                            ["0004", "Certidão", "Deferido", 2026, 50.0],
                        ],
                    }
                ],
                "extractionMethod": "digital",
                "warning": None,
            }
        ]

    def test_dashboard_is_first_sheet(self):
        xlsx_bytes, summary = build_workbook(self.sample_file_results)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        
        # O nome da primeira aba deve ser Dashboards
        self.assertEqual(wb.sheetnames[0], "Dashboards")
        self.assertIn("Principal", wb.sheetnames)

    def test_dashboard_sheet_contains_kpi_and_distributions(self):
        xlsx_bytes, summary = build_workbook(self.sample_file_results)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb["Dashboards"]

        # Verifica presenca do Banner/Titulo na coluna B (com margem esquerda na coluna A)
        self.assertIn("Painel", str(ws.cell(row=1, column=2).value) + str(ws.cell(row=2, column=2).value))
        
        # Procura por rotulos de KPI e tabelas de distribuicao na aba Dashboards
        dashboard_texts = [str(cell.value) for row in ws.iter_rows() for cell in row if cell.value is not None]
        all_text = " ".join(dashboard_texts)
        self.assertIn("REGISTROS", all_text.upper())
        self.assertIn("TIPO", all_text.upper())
        self.assertIn("SITUAÇÃO", all_text.upper())

    def test_data_sheet_does_not_contain_bottom_panels(self):
        xlsx_bytes, summary = build_workbook(self.sample_file_results)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb["Principal"]

        # Na aba de dados, nao deve mais existir o card de contabilizacao no rodape
        data_sheet_texts = [str(cell.value) for row in ws.iter_rows() for cell in row if cell.value is not None]
        for text in data_sheet_texts:
            self.assertNotIn("Contabilização por", text)

    def test_dashboard_with_responsavel(self):
        # Arquivo principal + arquivo secundario (responsavel)
        results = [
            {
                "fileName": "Principal.pdf",
                "tabName": "Principal",
                "tables": [
                    {
                        "tableName": "Setor 21",
                        "headers": ["Processo", "Tipo de Solicitação", "Situação"],
                        "rows": [
                            ["0001", "Alvará", "Deferido"],
                            ["0002", "Licença", "Em Análise"],
                        ],
                    }
                ],
                "extractionMethod": "digital",
                "warning": None,
            },
            {
                "fileName": "Joao.pdf",
                "tabName": "João Silva",
                "tables": [
                    {
                        "tableName": "João",
                        "headers": ["Processo", "Tipo de Solicitação"],
                        "rows": [["0001", "Alvará"]],
                    }
                ],
                "extractionMethod": "digital",
                "warning": None,
            },
        ]
        xlsx_bytes, summary = build_workbook(results)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb["Dashboards"]
        all_text = " ".join([str(c.value) for r in ws.iter_rows() for c in r if c.value is not None])
        self.assertIn("RESPONSÁVE", all_text.upper())
        self.assertIn("João Silva", all_text)

    def test_dashboard_formulas_bound_to_autofilter(self):
        xlsx_bytes, summary = build_workbook(self.sample_file_results)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb["Dashboards"]

        # KPI 1 deve usar SUBTOTAL(3, ...) para reagir ao autofiltro da aba de dados
        kpi1_formula = str(ws.cell(row=4, column=2).value)
        self.assertTrue(kpi1_formula.startswith("=SUBTOTAL(3,"))

        # Tabela 1 deve usar SUMPRODUCT com OFFSET/SUBTOTAL para contagem sensivel ao filtro
        qtd_formula = str(ws.cell(row=9, column=3).value)
        self.assertIn("SUMPRODUCT", qtd_formula)
        self.assertIn("SUBTOTAL(3, OFFSET", qtd_formula)


if __name__ == "__main__":
    unittest.main()
