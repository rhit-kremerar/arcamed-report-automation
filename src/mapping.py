"""Monthly workflow: read files -> validate -> map amounts -> report -> save a copy.

Client mapping rules live in CSV/Excel, not in Python. Keep the checks here so a
mistyped mapping cannot silently write an amount to the wrong reporting line.
"""
import argparse
import csv
import json
import re
import sys
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

TOLERANCE = Decimal('0.01')
MAPPING_COLUMNS = (
    '50001 account', '50001 description', 'TB RE account',
    'TB RE description', 'TB RE row', 'Amount to use',
)


def identifier(value):
    """Treat account codes as text, including numbers Excel reads as 100106.0."""
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return re.sub(r'^(\d+)\.0$', r'\1', text)


def normalized(value):
    """Ignore capitalization and extra spaces when comparing descriptions."""
    return ' '.join(str(value or '').split()).casefold()


def amount(value):
    """Read money without float rounding; Excel parentheses mean negative."""
    text = str(value if value is not None else '').strip().replace(',', '').replace('$', '')
    if text in ('', '-', '—'):
        return Decimal(0)
    if text.startswith('(') and text.endswith(')'):
        text = '-' + text[1:-1]
    result = Decimal(text)
    if not result.is_finite():
        raise ValueError(f'Invalid amount: {value!r}')
    return result


def period(value):
    """Return (year, month) for the date/header formats used in these files."""
    if isinstance(value, (date, datetime)):
        return value.year, value.month
    text = str(value or '').strip()
    for fmt in ('%b %Y', '%B %Y', "%b '%y", '%b %y', '%Y-%m', '%Y-%m-%d', '%m/%d/%Y'):
        try:
            parsed = datetime.strptime(text, fmt)
            return parsed.year, parsed.month
        except ValueError:
            pass
    return None


def read_rows(path, sheet):
    """Read values for validation; formula cells use Excel's last saved result."""
    if path.suffix.lower() == '.csv':
        with path.open(newline='', encoding='utf-8-sig') as handle:
            return list(csv.reader(handle))
    from openpyxl import load_workbook
    book = load_workbook(path, data_only=True, read_only=True)
    try:
        if sheet not in book.sheetnames:
            raise ValueError(f'Cannot find the worksheet "{sheet}" in {path.name}. '
                             f'Available tabs: {", ".join(book.sheetnames)}. '
                             'Select the correct workbook or change the sheet name in the app. '
                             'The RE Only template must contain a TB RE tab.')
        return list(book[sheet].values)
    finally:
        book.close()


def read_source(rows, errors):
    """Read the six trial-balance columns and check each account and total."""
    periods = set()
    for row in rows[:6]:
        for cell in row:
            match = re.search(r'\bPeriod\s*=\s*(.+)', str(cell or ''), re.I)
            if match:
                found = period(match.group(1))
                if found:
                    periods.add(found)
    if len(periods) != 1:
        raise ValueError('Cannot identify the reporting month. The first six rows of the trial balance '
                         'must contain one header such as "Period = Aug 2026". Check the input file and tab.')
    accounts = {}
    totals = None
    for index, row in enumerate(rows[6:], 7):
        values = list(row[:6]) + [None] * max(0, 6 - len(row))
        account = identifier(values[0])
        if not account and not any(v not in (None, '') for v in values):
            continue
        if normalized(values[0]).startswith('total') or normalized(values[1]).startswith('total'):
            totals = [amount(v) for v in values[2:6]]
            break
        if not re.fullmatch(r'\d{6}', account):
            errors.append(f'Trial balance row {index}: account {account!r} must be a six-digit account number.')
            continue
        if account in accounts:
            errors.append(f'Duplicate source account {account} at row {index}.')
            continue
        try:
            beginning, debit, credit, ending = [amount(v) for v in values[2:6]]
        except (InvalidOperation, ValueError) as exc:
            errors.append(f'Source row {index}: invalid balance: {exc}.')
            continue
        # Finance confirmed different rules for balance sheet and P&L accounts.
        number = int(account)
        if 100106 <= number <= 300555:
            rule, selected = 'ending balance', ending
        elif 400007 <= number <= 900900:
            rule, selected = 'debit minus credit', debit - credit
        else:
            errors.append(f'Account {account}: no approved calculation rule exists for this account number. Ask finance to confirm its rule.')
            continue
        if abs(beginning + debit - credit - ending) > TOLERANCE:
            errors.append(f'Account {account}: beginning + debit - credit differs from ending.')
        accounts[account] = dict(description=str(values[1] or ''), beginning=beginning,
                                 debit=debit, credit=credit, ending=ending, amount=selected, rule=rule)
    if not accounts:
        errors.append('No source accounts found.')
    sums = {key: sum((a[key] for a in accounts.values()), Decimal(0))
            for key in ('beginning', 'debit', 'credit', 'ending')}
    if abs(sums['debit'] - sums['credit']) > TOLERANCE:
        errors.append('Source debit and credit totals do not balance.')
    if abs(sums['ending']) > TOLERANCE:
        errors.append('Signed source ending balances do not sum to zero within one cent.')
    if totals is None:
        errors.append('Source Total row is missing.')
    else:
        for key, stated in zip(sums, totals):
            if abs(sums[key] - stated) > TOLERANCE:
                errors.append(f'Source Total row {key} does not match account sum.')
    return next(iter(periods)), accounts, sums


def read_mapping(path, errors, sheet='Mappings'):
    """Read the same six columns from either a CSV or an editable Excel sheet."""
    if path.suffix.lower() == '.csv':
        rows = read_rows(path, '')
    elif path.suffix.lower() == '.xlsx':
        from openpyxl import load_workbook
        # Read formulas as formulas so they can be rejected instead of using old caches.
        book = load_workbook(path, read_only=True, data_only=False)
        try:
            if sheet not in book.sheetnames:
                raise ValueError(f'Mapping workbook needs a {sheet!r} tab. Available tabs: {", ".join(book.sheetnames)}. '
                                 'Change the Excel mapping tab name in the app to match your file.')
            rows = list(book[sheet].values)
        finally:
            book.close()
    else:
        raise ValueError('Mapping file must be .csv or .xlsx.')

    headers = [str(value or '').strip() for value in rows[0]] if rows else []
    missing = [name for name in MAPPING_COLUMNS if name not in headers]
    if missing:
        raise ValueError(f'Mapping file is missing columns: {", ".join(missing)}.')
    if any(headers.count(name) != 1 for name in MAPPING_COLUMNS):
        raise ValueError('Mapping column headings must appear only once.')

    mappings = {}
    for row_number, values in enumerate(rows[1:], 2):
        # Ignore empty Excel rows and allow extra client notes columns.
        row = {name: values[headers.index(name)] if headers.index(name) < len(values) else None
               for name in MAPPING_COLUMNS}
        if all(value is None or str(value).strip() == '' for value in row.values()):
            continue
        if any(isinstance(value, str) and value.startswith('=') for value in row.values()):
            errors.append(f'Mapping row {row_number}: use typed values, not formulas.')
            continue
        account = identifier(row['50001 account'])
        if not re.fullmatch(r'\d{6}', account):
            errors.append(f'Mapping row {row_number}: invalid source ID {account!r}.')
            continue
        if account in mappings:
            errors.append(f'Duplicate mapping for source {account} at row {row_number}.')
            continue
        mappings[account] = row
    return mappings


def find_month_column(target_rows, source_period):
    """Find the monthly column, excluding repeated headers in the YTD section.

    This helper returns a zero-based list index; Excel columns are index + 1.
    """
    if len(target_rows) < 3:
        raise ValueError('Destination must contain row 3 monthly headers.')
    # Only the monthly section is eligible; hidden YTD columns repeat the headers.
    ytd = [i for i, value in enumerate(target_rows[1]) if normalized(value) == 'ytd']
    monthly_end = min(ytd) if ytd else len(target_rows[2])
    columns = [i for i, value in enumerate(target_rows[2][:monthly_end])
               if period(value) == source_period]
    if len(columns) != 1:
        raise ValueError(f'Cannot select the reporting month {source_period[0]}-{source_period[1]:02d}: '
                         f'found {len(columns)} matching headers in TB RE row 3. '
                         'The template needs exactly one matching monthly column before the YTD section.')
    return columns[0]


def index_destinations(target_rows):
    """Index reporting lines by both code and description; codes can repeat."""
    identities = defaultdict(list)
    eligible = set()
    for index, row in enumerate(target_rows[3:], 4):
        if len(row) >= 3 and normalized(row[0]) and (normalized(row[2]) or (len(row) >= 6 and normalized(row[5]))):
            identities[(identifier(row[1]), normalized(row[2]))].append(index)
            eligible.add(index)
    return identities, eligible


def aggregate_mappings(accounts, mappings, target_rows, errors):
    """Validate each destination, then sum source amounts into reporting rows."""
    identities, eligible = index_destinations(target_rows)
    aggregated = defaultdict(lambda: Decimal(0))
    sources = defaultdict(list)
    unmapped, moved, missing_sources = [], [], []
    for account, mapping in mappings.items():
        if account not in accounts:
            missing_sources.append(account)
        code = identifier(mapping['TB RE account'])
        description = normalized(mapping['TB RE description'])
        row_hint = identifier(mapping['TB RE row'])
        parts = [bool(code), bool(description), bool(row_hint)]
        if any(parts) and not all(parts):
            errors.append(f'Account {account}: fill in all three destination fields (TB RE account, description and row), or leave all three blank.')
            continue
        if not any(parts):
            if account in accounts:
                unmapped.append(dict(account=account, description=accounts[account]['description'],
                                     amount=accounts[account]['amount']))
            continue
        if code.casefold() == 'no account code':
            code = ''
        matches = identities.get((code, description), [])
        if not code:
            # Explicit no-code lines can have their identity only in E/F.
            matches = [index for index in eligible if len(target_rows[index - 1]) >= 6
                       and not identifier(target_rows[index - 1][1])
                       and not normalized(target_rows[index - 1][2])
                       and identifier(target_rows[index - 1][4]) == account
                       and normalized(target_rows[index - 1][5]) == description]
        if not row_hint.isdigit() or int(row_hint) < 4:
            errors.append(f'Account {account}: TB RE row {row_hint!r} must be a whole row number of 4 or higher.')
            continue
        hinted = int(row_hint)
        if hinted in matches:
            destination = hinted
        elif len(matches) == 1:
            destination = matches[0]
            moved.append(dict(account=account, old_row=hinted, new_row=destination))
        else:
            errors.append(f'Account {account}: cannot uniquely match reporting line "{code} - {description}". '
                          'Copy its account, description and row from the current TB RE template into the mapping.')
            continue
        if account not in accounts:
            continue
        source = accounts[account]
        if normalized(mapping['50001 description']) != normalized(source['description']):
            errors.append(f'Account {account}: the mapping description does not match the trial balance. '
                          f'Use "{source["description"]}" after checking that it is the correct account.')
        if normalized(mapping['Amount to use']) != source['rule']:
            errors.append(f'Account {account}: Amount to use must be "{source["rule"]}". Correct this field in the mapping.')
        aggregated[destination] += source['amount']
        sources[destination].append(account)
    for account in accounts.keys() - mappings.keys():
        unmapped.append(dict(account=account, description=accounts[account]['description'],
                             amount=accounts[account]['amount']))
    return aggregated, sources, unmapped, moved, missing_sources, eligible


def compare_destination(target_rows, column, aggregated, sources, eligible):
    """List missing reporting lines and compare mapped amounts to saved values."""
    unresolved = []
    for row in sorted(eligible - aggregated.keys()):
        values = target_rows[row - 1]
        unresolved.append(dict(row=row, account=identifier(values[1]), description=values[2],
                               existing=values[column] if column < len(values) else None))
    comparisons = []
    for row, selected in sorted(aggregated.items()):
        values = target_rows[row - 1]
        existing = values[column] if column < len(values) else None
        try:
            prior = None if existing in (None, '') else amount(existing)
        except (InvalidOperation, ValueError):
            prior = None
        comparisons.append(dict(row=row, account=identifier(values[1]), description=values[2],
                                sources=sources[row], amount=selected, existing=existing,
                                difference=selected - prior if prior is not None else None))
    return unresolved, comparisons


def build_report(source_rows, target_rows, mapping_path, requested_period=None, mapping_sheet='Mappings'):
    """Coordinate the read/validate/map steps without changing any workbook."""
    errors, warnings = [], []
    source_period, accounts, totals = read_source(source_rows, errors)
    if requested_period and requested_period != source_period:
        errors.append('Requested month differs from input reporting period.')
    column = find_month_column(target_rows, source_period)
    mappings = read_mapping(mapping_path, errors, mapping_sheet)
    aggregated, sources, unmapped, moved, missing_sources, eligible = aggregate_mappings(
        accounts, mappings, target_rows, errors)
    unresolved, comparisons = compare_destination(target_rows, column, aggregated, sources, eligible)
    warnings.extend([
        'Draft mapping requires finance review; 406010 candidate is 6706-0000 / Bill Back - Utilities / row 91.',
        'Finance should confirm 610400 Fire/Life/Safety and the new mortgage account 210030.',
        'Only mapped cells in the selected month are replaced; report formulas elsewhere are retained.',
    ])
    if source_period == (2026, 8):
        for key in ('debit', 'credit'):
            if abs(totals[key] - Decimal('2319336.99')) > TOLERANCE:
                errors.append(f'August 2026 {key} total differs from expected 2,319,336.99.')
    report = dict(period=f'{source_period[0]:04d}-{source_period[1]:02d}', month_column=column + 1,
                  source_count=len(accounts), totals=totals, errors=errors, warnings=warnings,
                  unmapped_accounts=unmapped, mapping_accounts_absent_from_source=missing_sources,
                  moved_destinations=moved, unresolved_destination_lines=unresolved,
                  comparisons=comparisons)
    return report, aggregated, column + 1


def parse_arguments(argv=None):
    """Define the options shown by: python src/mapping.py --help."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', type=Path, required=True)
    parser.add_argument('--template', type=Path, required=True)
    parser.add_argument('--mapping', type=Path, required=True, help='Mapping .xlsx or .csv file')
    parser.add_argument('--mapping-sheet', default='Mappings', help='Excel mapping tab name (default: Mappings)')
    parser.add_argument('--output', type=Path)
    parser.add_argument('--report', type=Path, required=True)
    parser.add_argument('--month', help='Optional expected YYYY-MM period')
    parser.add_argument('--source-sheet', default='Report1 (6)')
    parser.add_argument('--audit-only', action='store_true')
    parser.add_argument('--allow-unmapped', action='store_true', help='Explicitly permit a partial draft output')
    parser.add_argument('--missing-lines', choices=('error', 'preserve', 'zero', 'blank'), default='error')
    return parser.parse_args(argv)


def validate_paths(args):
    """Protect input files and require a new output/report path for every run."""
    if not args.audit_only and (not args.output or args.output.suffix.lower() != '.xlsx'):
        raise ValueError('Workbook conversion requires --output ending in .xlsx.')
    if not args.audit_only and (args.input.suffix.lower() != '.xlsx' or args.template.suffix.lower() != '.xlsx'):
        raise ValueError('Conversion requires original .xlsx files; CSV exports support --audit-only.')
    protected = {p.resolve() for p in (args.input, args.template, args.mapping)}
    for path in [args.report] + ([args.output] if args.output else []):
        if path.resolve() in protected or path.exists():
            raise ValueError(f'Refusing to overwrite input or existing file: {path}')
    if args.output and args.output.resolve() == args.report.resolve():
        raise ValueError('Output workbook and audit report must have different paths.')
    requested = period(args.month) if args.month else None
    if args.month and (not re.fullmatch(r'\d{4}-\d{2}', args.month) or requested is None):
        raise ValueError('--month must be YYYY-MM.')
    return requested


def write_workbook(template, output, report, aggregated, column, missing_lines):
    """Change only selected-month values in a copy; retain other formulas."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.views import Selection
    from openpyxl.workbook.properties import CalcProperties

    # data_only=False is essential: loading cached values would discard formulas.
    book = load_workbook(template, data_only=False)
    try:
        sheet = book['TB RE']
        for row, selected in aggregated.items():
            sheet.cell(row, column).value = float(selected)
        if missing_lines in ('zero', 'blank'):
            for line in report['unresolved_destination_lines']:
                cell = sheet.cell(line['row'], column)
                if cell.data_type != 'f':
                    cell.value = 0 if missing_lines == 'zero' else None
        book['MASTER DEFINITION']['B1'] = int(report['period'][5:])
        # The template opens on an old lookup tab. Show the updated report instead.
        sheet.sheet_state = 'visible'
        book.active = book.index(sheet)
        for tab in book:
            tab.sheet_view.tabSelected = tab is sheet
        for view in book.views:
            view.showSheetTabs = True
            view.firstSheet = 0
            view.activeTab = book.index(sheet)
        month_cell = f'{get_column_letter(column)}4'
        sheet.sheet_view.topLeftCell = f'{get_column_letter(column)}1'
        pane = sheet.sheet_view.pane
        if pane:
            pane.topLeftCell = month_cell
        sheet.sheet_view.selection = [Selection(
            pane=pane.activePane if pane else None, activeCell=month_cell, sqref=month_cell)]
        # openpyxl cannot calculate formulas; Excel refreshes them on opening.
        book.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True, forceFullCalc=True, calcMode='auto')
        book.save(output)
    finally:
        book.close()


def main(argv=None):
    """Run the monthly workflow and return a command-line exit code."""
    args = parse_arguments(argv)
    try:
        # 1. Read and validate before deciding whether a workbook can be written.
        requested = validate_paths(args)
        report, aggregated, column = build_report(read_rows(args.input, args.source_sheet),
                                                  read_rows(args.template, 'TB RE'), args.mapping,
                                                  requested, args.mapping_sheet)
        # 2. Unresolved decisions need explicit client choices, never guesses.
        blocks = list(report['errors'])
        if report['unmapped_accounts'] and not args.allow_unmapped:
            blocks.append('Unmapped source accounts remain. Review mappings or explicitly use --allow-unmapped.')
        if report['unresolved_destination_lines'] and args.missing_lines == 'error':
            blocks.append('Destination lines without source amounts remain; choose --missing-lines blank, preserve or zero.')
        report['missing_lines_policy'] = args.missing_lines
        report['write_blockers'] = blocks
        report['output_written'] = False
        # 3. Save a new workbook only when validation and mapping decisions pass.
        if not args.audit_only and not blocks:
            write_workbook(args.template, args.output, report, aggregated, column, args.missing_lines)
            report['output_written'] = True
        # 4. Save the audit so finance can review amounts and unresolved items.
        with args.report.open('x', encoding='utf-8') as handle:
            json.dump(report, handle, indent=2, default=str)
        print(f"Period {report['period']}: {report['source_count']} accounts, {len(aggregated)} mapped lines, "
              f"{len(report['unmapped_accounts'])} unmapped accounts. Report: {args.report}")
        if blocks:
            print('\n'.join(blocks), file=sys.stderr)
        if report['output_written']:
            print(f'Created {args.output}. Open in Excel and save to recalculate formulas.')
        return 1 if report['errors'] or (blocks and not args.audit_only) else 0
    except (OSError, ValueError, KeyError, ImportError, InvalidOperation) as exc:
        print(f'Error: {exc}', file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
