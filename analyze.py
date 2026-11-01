"""Импорт годовых XLSX ФНС и анализ отчётности. Python 3.10+, openpyxl."""
import argparse
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import json
import os
import re
import tempfile
import time
import warnings

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrRef, StrData, StrVal, NumData, NumVal
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

VERSION = '2.0.2'
NAMES = {
    '1100': 'Внеоборотные активы', '1200': 'Оборотные активы',
    '1300': 'Собственный капитал', '1400': 'Долгосрочные обязательства',
    '1500': 'Краткосрочные обязательства', '1600': 'Активы, всего',
    '1700': 'Пассивы, всего', '2110': 'Выручка', '2400': 'Чистая прибыль',
}
HEADERS = ['Год', 'Доля оборотных активов, %', 'Текущая ликвидность',
           'Финансовая автономия', 'Чистая рентабельность продаж, %']


def number(value, context=''):
    """Явный прочерк ФНС означает ноль; пустая ячейка — ошибку."""
    if value is None or str(value).strip() == '':
        raise ValueError('Не заполнено значение: ' + context)
    text = re.sub(r'[\s\u00a0\u202f]', '', str(value)).replace(',', '.')
    if text in ('-', '—', '–', '(-)', '(—)', '(–)'):
        return Decimal(0)
    if text.startswith('(') and text.endswith(')'):
        text = '-' + text[1:-1]
    try:
        result = Decimal(text)
    except InvalidOperation:
        raise ValueError(f'Неверное число {value!r}: {context}') from None
    if not result.is_finite():
        raise ValueError('Недопустимое число: ' + context)
    return result


def norm(text):
    return re.sub(r'\s+', ' ', str(text or '').lower().replace('ё', 'е')).strip()


def metadata(sheet):
    info = {}
    for row in sheet.iter_rows():
        cells = [c.value for c in row if c.value is not None and str(c.value).strip()]
        if len(cells) < 2:
            continue
        label = norm(cells[0])
        if label == 'инн' and 'inn' not in info:
            info['inn'] = str(cells[1]).strip()
        elif label == 'единица измерения':
            info['unit'] = norm(cells[1])
        elif label == 'полное наименование юридического лица':
            info['name'] = str(cells[1])
    if not re.fullmatch(r'\d{10}|\d{12}', info.get('inn', '')):
        raise ValueError('Не найден ИНН организации')
    if info.get('unit') not in ('тыс. руб.', 'тыс. руб', 'тыс руб', 'тыс руб.'):
        raise ValueError('Поддерживаются только отчёты в тыс. руб.')
    return info


def parse_sheet(sheet, wanted, filename):
    headers = []
    for row in sheet.iter_rows(min_row=1, max_row=min(20, sheet.max_row)):
        for cell in row:
            if norm(cell.value) == 'код строки':
                headers.append(cell)
    if len(headers) != 1:
        raise ValueError(f'{filename}: не найден однозначный заголовок «Код строки» на {sheet.title}')
    header = headers[0]
    columns = {}
    for cell in sheet[header.row]:
        if cell.column > header.column:
            match = re.search(r'\b(20\d{2})\b', str(cell.value or ''))
            if match:
                year = match.group(1)
                if year in columns:
                    raise ValueError('Повтор года в заголовке: ' + year)
                columns[year] = cell.column
    if not columns:
        raise ValueError('Не найдены годы в заголовках: ' + sheet.title)
    result = {y: {} for y in columns}
    seen = set()
    for row in sheet.iter_rows(min_row=header.row + 1):
        raw_code = row[header.column - 1].value
        code = str(raw_code).strip()
        if code not in wanted:
            continue
        if code in seen:
            raise ValueError(f'Повтор кода {code}: {filename}, {sheet.title}')
        seen.add(code)
        for year, column in columns.items():
            cell = row[column - 1]
            context = f'{filename}, {sheet.title}!{cell.coordinate}, {year}, код {code}'
            value = number(cell.value, context)
            if value < 0 and code != '2400':
                raise ValueError('Отрицательное значение укрупнённой статьи: ' + context)
            result[year][code] = (value, sheet.title, cell.coordinate, str(cell.value))
    if seen != set(wanted):
        raise ValueError(f'{filename}: нет строк {sorted(set(wanted) - seen)}')
    return result


def read_report(path):
    if path.stat().st_size > 20 * 1024 * 1024:
        raise ValueError('Отчёт превышает ограничение 20 МБ: ' + path.name)
    with warnings.catch_warnings():
        warnings.filterwarnings('ignore', message='Workbook contains no default style')
        book = openpyxl.load_workbook(path, data_only=True)
    try:
        sheets = {norm(s.title): s for s in book}
        needed = ['сведения об организации', 'бухгалтерский баланс', 'отчет о финансовых результатах']
        if any(name not in sheets for name in needed):
            raise ValueError(f'{path.name}: нужны сведения, баланс и отчёт о финансовых результатах')
        info = metadata(sheets[needed[0]])
        balance = parse_sheet(sheets[needed[1]], set(NAMES) - {'2110', '2400'}, path.name)
        financial = parse_sheet(sheets[needed[2]], {'2110', '2400'}, path.name)
        current = max(balance, key=int)
        if current != max(financial, key=int):
            raise ValueError(path.name + ': годы баланса и финансовых результатов различаются')
        periods = {}
        for year in balance.keys() | financial.keys():
            periods[year] = {**balance.get(year, {}), **financial.get(year, {})}
        return {**info, 'year': current, 'periods': periods, 'file': path.name}
    finally:
        book.close()


def load_reports(folder, expected_inn='2126000147'):
    paths = sorted(p for p in Path(folder).glob('*.xlsx') if not p.name.startswith('~$'))
    if not paths:
        raise ValueError('В папке reports нет отчётов .xlsx. Распакуйте архивы ФНС.')
    reports = [read_report(path) for path in paths]
    if any(r['inn'] != expected_inn for r in reports):
        raise ValueError('ИНН отчёта не совпадает с ожидаемым ' + expected_inn)
    owners = {}
    for report in reports:
        year = report['year']
        if year in owners:
            raise ValueError(f'Два отчёта за {year}. Оставьте одну выбранную версию отчёта.')
        owners[year] = report
    years = sorted(owners, key=int)
    if len(years) < 2:
        raise ValueError('Нужно минимум два годовых отчёта; каждый год берётся из своего отчёта')
    data, sources = {c: {} for c in NAMES}, []
    for year in years:
        owner = owners[year]
        for code in NAMES:
            value, sheet, cell, raw = owner['periods'][year][code]
            for report in reports:
                other = report['periods'].get(year, {}).get(code)
                if other is not None and other[0] != value:
                    raise ValueError(f'Расхождение за {year}, код {code}: {owner["file"]} = {value}; '
                                     f'{report["file"]} = {other[0]}. Требуется выбор сопоставимой версии.')
            data[code][year] = value
            sources.append([int(year), code, NAMES[code], float(value), owner['file'], sheet, cell, raw])
    return years, data, sources, owners[years[-1]]['name']


def divide(a, b, multiplier=1):
    return None if b == 0 else a / b * multiplier


def indicators(data, years):
    return [[int(y), divide(data['1200'][y], data['1600'][y], 100),
             divide(data['1200'][y], data['1500'][y]),
             divide(data['1300'][y], data['1600'][y]),
             divide(data['2400'][y], data['2110'][y], 100)] for y in years]


def balance_checks(data, years):
    rows = []
    for y in years:
        asset = data['1600'][y]
        differences = [asset - data['1700'][y], asset - data['1100'][y] - data['1200'][y],
                       asset - sum(data[c][y] for c in ('1300', '1400', '1500'))]
        for label, difference in zip(('Актив = пассив', 'Сумма активов', 'Сумма пассивов'), differences):
            rows.append([int(y), label, difference, 'OK' if abs(difference) <= 2 else 'Ошибка'])
    if any(r[-1] != 'OK' for r in rows):
        raise ValueError('Баланс не сходится: ' + '; '.join(str(r) for r in rows if r[-1] != 'OK'))
    return rows


def excel_value(value):
    if isinstance(value, Decimal):
        return float(value)
    if value is None:
        return 'н/р'
    # Имена/исходный текст не должны превращаться в формулы Excel.
    if isinstance(value, str) and value.startswith(('=', '+', '-', '@')):
        return "'" + value
    return value


def table(book, title, headers, rows, widths):
    sheet = book.create_sheet(title)
    sheet.append(headers)
    for row in rows:
        sheet.append([excel_value(v) for v in row])
    sheet.freeze_panes = 'B2'
    sheet.auto_filter.ref = sheet.dimensions
    sheet.sheet_view.showGridLines = False
    for cell in sheet[1]:
        cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='17365D')
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    sheet.row_dimensions[1].height = 42
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name='Calibri', size=11)
            cell.alignment = Alignment(vertical='top', wrap_text=isinstance(cell.value, str))
            if isinstance(cell.value, float):
                cell.number_format = '#,##0.00;[Red](#,##0.00);0.00'
            if cell.row % 2 == 0:
                cell.fill = PatternFill('solid', fgColor='EEF3F8')
    for i, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(i)].width = width
    return sheet


def write_excel(path, years, data, sources, company, checks):
    book = openpyxl.Workbook()
    book.remove(book.active)
    book.properties.title = 'Анализ бухгалтерской отчётности: ' + company
    book.properties.creator = 'AKKOND Excel Analyzer ' + VERSION
    result = table(book, 'Коэффициенты', HEADERS, indicators(data, years), [12, 26, 25, 25, 30])
    result.append([])
    start_method = len(years) + 3
    result.merge_cells(start_row=start_method, start_column=1, end_row=start_method, end_column=5)
    result.cell(start_method, 1, 'Формулы по кодам строк бухгалтерской отчётности')
    for index, (name, formula) in enumerate(zip(HEADERS[1:], ['1200 / 1600 × 100', '1200 / 1500', '1300 / 1600', '2400 / 2110 × 100']), start_method+1):
        result.merge_cells(start_row=index, start_column=1, end_row=index, end_column=2)
        result.merge_cells(start_row=index, start_column=3, end_row=index, end_column=5)
        result.cell(index, 1, name)
        result.cell(index, 3, formula)
        result.row_dimensions[index].height = 32
    note_row = start_method + 5
    result.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=5)
    result.cell(note_row, 1, 'При нулевом знаменателе — н/р. Расчёты выполняются Decimal. В XLSX сохранены числовые результаты; отображение — два знака после запятой.')
    result.row_dimensions[note_row].height = 42
    for row in result.iter_rows(min_row=start_method):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='center')
            cell.font = Font(name='Calibri', size=11, bold=cell.row == start_method)
    result.auto_filter.ref = f'A1:E{len(years)+1}'
    chart = LineChart()
    chart.title = 'Текущая ликвидность (коэффициент)'
    chart.title.overlay = False
    chart.layout = Layout(manualLayout=ManualLayout(
        x=0.13, y=0.18, w=0.80, h=0.67, xMode='edge', yMode='edge'))
    chart.add_data(Reference(result, min_col=3, min_row=1, max_row=len(years)+1), titles_from_data=True)
    chart.set_categories(Reference(result, min_col=1, min_row=2, max_row=len(years)+1))
    chart.style = 13
    chart.varyColors = False
    chart.legend = None
    chart.width, chart.height = 19, 11
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.tickLblPos = chart.y_axis.tickLblPos = 'nextTo'
    chart.x_axis.axPos = 'b'
    chart.y_axis.axPos = 'l'
    chart.y_axis.scaling.min = 0
    chart.y_axis.numFmt = '0.00'
    chart.y_axis.majorTickMark = 'out'
    series = chart.series[0]
    series.smooth = False
    series.marker.symbol = 'circle'
    series.marker.size = 7
    series.marker.graphicalProperties.solidFill = '2364AA'
    series.marker.graphicalProperties.line.solidFill = '2364AA'
    series.graphicalProperties.line.solidFill = '2364AA'
    series.cat = AxDataSource(strRef=StrRef(
        f=f"'Коэффициенты'!$A$2:$A${len(years)+1}",
        strCache=StrData(ptCount=len(years), pt=[StrVal(idx=i, v=y) for i,y in enumerate(years)])))
    values = [row[2] for row in indicators(data, years)]
    numeric_values = [float(v) for v in values if v is not None]
    chart.y_axis.scaling.max = max(1, max(numeric_values, default=0) * 1.30)
    series.val.numRef.numCache = NumData(formatCode='0.00', ptCount=len(years),
        pt=[NumVal(idx=i, v=float(v)) for i,v in enumerate(values) if v is not None])
    chart.dataLabels = DataLabelList(showVal=True, showLegendKey=False,
        showSerName=False, showCatName=False, showPercent=False,
        numFmt='0.00', dLblPos='t')
    result.add_chart(chart, 'G2')
    dynamics = []
    for code in ('1600', '1200', '1300', '1400', '1500', '2110', '2400'):
        for old, new in zip(years, years[1:]):
            before, after = data[code][old], data[code][new]
            dynamics.append([code, NAMES[code], f'{old}–{new}', before, after, after-before,
                             divide(after-before, before, 100) if before > 0 else None,
                             '' if before > 0 else 'База нулевая или отрицательная'])
    table(book, 'Динамика', ['Код', 'Показатель', 'Период', 'База, тыс. руб.', 'Отчёт, тыс. руб.',
                           'Изменение, тыс. руб.', 'Прирост, %', 'Примечание'], dynamics,
          [12, 34, 17, 23, 23, 25, 18, 35])
    structure = [[int(y), c, NAMES[c], data[c][y], data['1600'][y], divide(data[c][y], data['1600'][y], 100)]
                 for y in years for c in ('1100', '1200', '1300', '1400', '1500')]
    table(book, 'Структура', ['Год', 'Код', 'Показатель', 'Сумма, тыс. руб.', 'Баланс, тыс. руб.', 'Доля, %'],
          structure, [12, 12, 36, 24, 24, 18])
    raw = table(book, 'Исходные данные', ['Код', 'Показатель', 'Единица'] + [int(y) for y in years],
                [[c, NAMES[c], 'тыс. руб.'] + [data[c][y] for y in years] for c in NAMES],
                [12, 36, 18] + [21]*len(years))
    finance_chart = BarChart()
    finance_chart.title = 'Выручка и чистая прибыль (тыс. руб.)'
    finance_chart.title.overlay = False
    finance_chart.layout = Layout(manualLayout=ManualLayout(
        x=0.20, y=0.18, w=0.73, h=0.63, xMode='edge', yMode='edge'))
    finance_chart.add_data(Reference(raw, min_row=9, max_row=10,
                                    min_col=3, max_col=3+len(years)),
                           titles_from_data=True, from_rows=True)
    finance_chart.set_categories(Reference(raw, min_row=1, min_col=4, max_col=3+len(years)))
    finance_chart.style = 13
    finance_chart.varyColors = False
    finance_chart.x_axis.delete = finance_chart.y_axis.delete = False
    finance_chart.x_axis.tickLblPos = finance_chart.y_axis.tickLblPos = 'nextTo'
    finance_chart.x_axis.axPos = 'b'
    finance_chart.y_axis.axPos = 'l'
    finance_chart.y_axis.numFmt = '#,##0'
    finance_chart.legend.position = 'b'
    finance_chart.legend.overlay = False
    if all(data[c][y] >= 0 for c in ('2110', '2400') for y in years):
        finance_chart.y_axis.scaling.min = 0
    finance_chart.width, finance_chart.height = 20, 11
    # Названия серий берём из колонки показателей, а не из единицы измерения.
    from openpyxl.chart.series import SeriesLabel
    for series, label, color in zip(finance_chart.series,
            ('Выручка', 'Чистая прибыль'), ('2364AA', 'E69F00')):
        series.tx = SeriesLabel(v=label)
        series.graphicalProperties.solidFill = color
        series.graphicalProperties.line.solidFill = color
    raw.add_chart(finance_chart, 'H2')
    table(book, 'Проверки', ['Год', 'Проверка', 'Разница, тыс. руб.', 'Статус'], checks, [12, 30, 24, 17])
    table(book, 'Источники', ['Год', 'Код', 'Показатель', 'Значение, тыс. руб.', 'Файл', 'Лист', 'Ячейка', 'Исходный текст'],
          sources, [12, 12, 34, 25, 58, 40, 14, 25])
    notes = [
        ['Организация', company], ['Версия программы', VERSION],
        ['Создано', datetime.now().astimezone().isoformat(timespec='seconds')],
        ['Выбор периодов', 'Каждый год — из текущего столбца отдельного годового отчёта. Сравнительные столбцы проверяются на совпадение по девяти используемым кодам.'],
        ['Область проверки', 'Сопоставление не проверяет все статьи отчётности и не подтверждает её достоверность. При изменении формы XLSX может понадобиться обновление импортера.'],
        ['Баланс', 'Допуск 2 тыс. руб. на округление укрупнённых статей.'],
        ['Динамика', 'Абсолютное изменение = отчёт − база; прирост = изменение / база × 100. При базе ≤ 0 прирост не рассчитывается.'],
        ['Структура', 'Внеоборотные + оборотные активы составляют отдельную группу 100%; капитал + долгосрочные + краткосрочные обязательства — другую. Не суммировать все пять долей.'],
        ['Прочерки', 'Явные прочерки ФНС приняты за ноль; пустые обязательные ячейки вызывают ошибку.'],
        ['Точность', 'Совпадение с Excel проверяется при отображении двух знаков; это не доказательство отсутствия ошибок на любых данных.'],
        ['Время', 'В консоли — время чтения, вычислений и записи. Действия пользователя и проверка человеком не включены.'],
        ['Предыдущий эксперимент', 'Сокращение времени 57,8% получено для версии CSV; на эту версию XLSX не переносится.'],
    ]
    table(book, 'Методика', ['Параметр', 'Описание'], notes, [32, 110])
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(suffix='.xlsx', dir=path.parent)
    os.close(handle)
    try:
        book.save(temporary)
        # Контроль готового файла до замены предыдущего результата.
        verify = openpyxl.load_workbook(temporary, data_only=True)
        assert verify['Коэффициенты'].max_row >= len(years)+1
        verify.close()
        os.replace(temporary, path)
    finally:
        book.close()
        if os.path.exists(temporary):
            os.unlink(temporary)


def analyze(folder, output, inn='2126000147'):
    years, data, sources, company = load_reports(folder, inn)
    checks = balance_checks(data, years)
    write_excel(output, years, data, sources, company, checks)
    return len(years), len(checks)


def main():
    base = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--reports', type=Path, default=base/'reports')
    parser.add_argument('--output', type=Path, default=base/'results'/'Анализ_АККОНД.xlsx')
    parser.add_argument('--inn', default='2126000147')
    args = parser.parse_args()
    started = time.perf_counter()
    try:
        count, checks = analyze(args.reports, args.output, args.inn)
    except Exception as error:
        # Ошибка должна быть видна пользователю Windows; старый файл не выдаётся за новый.
        print('ОШИБКА:', error)
        print('Новый анализ не создан. Если старый файл существует, он НЕ обновлён.')
        return 1
    elapsed = time.perf_counter() - started
    print(f'Готово. Лет: {count}; проверок баланса: {checks}.')
    print('Результат:', args.output.resolve())
    print(f'Чтение, расчёты и запись: {elapsed:.3f} с.')
    print('Это время программы, без действий пользователя и проверки человеком.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
