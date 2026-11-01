"""Проверки импорта, расчётов, отказов и сохранения. python -m unittest -v"""
import shutil
import tempfile
import unittest
import warnings
from decimal import Decimal
from pathlib import Path
import openpyxl
import analyze

BASE = Path(__file__).resolve().parent


class AnalysisTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.reports = self.root/'reports'
        shutil.copytree(BASE/'reports', self.reports)

    def tearDown(self):
        self.tmp.cleanup()

    def edit(self, year, sheet, cell, value):
        path = next(self.reports.glob(f'*_{year}_*.xlsx'))
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', UserWarning)
            book = openpyxl.load_workbook(path)
        book[sheet][cell] = value
        book.save(path)
        book.close()

    def test_all_27_source_values(self):
        years, data, sources, company = analyze.load_reports(self.reports)
        self.assertEqual(years, ['2023', '2024', '2025'])
        expected = {
            '1100': [2913239,2565939,2480094], '1200': [7395989,6882274,6989243],
            '1300': [7522643,7995006,7951739], '1400': [172396,130308,104885],
            '1500': [2614189,1322899,1412713], '1600': [10309228,9448213,9469337],
            '1700': [10309228,9448213,9469337], '2110': [16268221,18925717,21714136],
            '2400': [2285641,2852109,3357237]}
        for code, values in expected.items():
            self.assertEqual([data[code][y] for y in years], list(map(Decimal, values)))
        self.assertEqual(len(sources), 27)
        self.assertIn('АККОНД', company.upper())

    def test_12_indicators(self):
        years, data, _, _ = analyze.load_reports(self.reports)
        expected = [['71.74','2.83','0.73','14.05'], ['72.84','5.20','0.85','15.07'],
                    ['73.81','4.95','0.84','15.46']]
        for actual, want in zip(analyze.indicators(data, years), expected):
            self.assertEqual([str(v.quantize(Decimal('.01'))) for v in actual[1:]], want)
        self.assertEqual(len(analyze.balance_checks(data, years)), 9)

    def test_result_numeric_and_traceable(self):
        path = self.root/'result.xlsx'
        self.assertEqual(analyze.analyze(self.reports, path), (3,9))
        book = openpyxl.load_workbook(path, data_only=True)
        self.assertEqual(len(book.sheetnames), 7)
        for row in book['Коэффициенты'].iter_rows(min_row=2, max_row=4, min_col=2, max_col=5):
            self.assertTrue(all(c.data_type == 'n' for c in row))
        self.assertEqual(book['Источники'].max_row, 28)
        self.assertAlmostEqual(book['Коэффициенты']['C2'].value, 7395989/2614189, places=12)
        book.close()

    def test_comparative_conflict(self):
        self.edit(2025, 'Бухгалтерский баланс', 'M41', '6 882 275')
        with self.assertRaisesRegex(ValueError, 'Расхождение за 2024, код 1200'):
            analyze.load_reports(self.reports)

    def test_empty_required_cell(self):
        self.edit(2025, 'Бухгалтерский баланс', 'K41', None)
        with self.assertRaisesRegex(ValueError, 'Не заполнено значение'):
            analyze.load_reports(self.reports)

    def test_wrong_company(self):
        self.edit(2025, 'Сведения об организации', 'H10', '1234567890')
        with self.assertRaisesRegex(ValueError, 'ИНН отчёта'):
            analyze.load_reports(self.reports)

    def test_wrong_unit(self):
        self.edit(2025, 'Сведения об организации', 'H16', 'Млн. руб.')
        with self.assertRaisesRegex(ValueError, 'тыс. руб.'):
            analyze.load_reports(self.reports)

    def test_duplicate_year(self):
        f = next(self.reports.glob('*_2025_*.xlsx'))
        shutil.copyfile(f, self.reports/'duplicate.xlsx')
        with self.assertRaisesRegex(ValueError, 'Два отчёта за 2025'):
            analyze.load_reports(self.reports)

    def test_wrong_balance_keeps_previous_output(self):
        output = self.root/'result.xlsx'
        output.write_bytes(b'previous report sentinel')
        self.edit(2025, 'Бухгалтерский баланс', 'K71', '9 469 000')
        with self.assertRaisesRegex(ValueError, 'Баланс не сходится'):
            analyze.analyze(self.reports, output)
        self.assertEqual(output.read_bytes(), b'previous report sentinel')

    def test_zero_and_negative_profit(self):
        self.assertIsNone(analyze.divide(Decimal(1), Decimal(0)))
        self.assertEqual(analyze.number('(1 234,50)'), Decimal('-1234.50'))
        self.assertEqual(analyze.number('(-)'), Decimal(0))
        with self.assertRaises(ValueError):
            analyze.number('NaN')
        self.edit(2025, 'Отчет о финансовых результатах', 'L24', '(1 000)')
        years, data, _, _ = analyze.load_reports(self.reports)
        self.assertLess(analyze.indicators(data, years)[-1][-1], 0)

    def test_missing_report_folder(self):
        with self.assertRaisesRegex(ValueError, 'нет отчётов'):
            analyze.load_reports(self.root/'missing')

    def test_removed_required_code(self):
        self.edit(2025, 'Бухгалтерский баланс', 'I41', None)
        with self.assertRaisesRegex(ValueError, 'нет строк'):
            analyze.load_reports(self.reports)


if __name__ == '__main__':
    unittest.main()
