"""Readable PDF findings for finance; JSON remains a developer option."""
from decimal import Decimal
from xml.sax.saxutils import escape


def write_pdf(path, report):
    """Write every finding with wrapping and automatic page breaks."""
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

    styles = getSampleStyleSheet()
    content = []

    def line(text, style='BodyText'):
        content.append(Paragraph(escape(str(text)), styles[style]))
        content.append(Spacer(1, 6))

    def section(title, findings):
        line(title, 'Heading2')
        for finding in findings or ['None found.']:
            line(finding)

    line('Arcamed trial balance audit', 'Title')
    line(f"Reporting month: {report['period']} | Source accounts: {report['source_count']}")
    line('This report checks the trial balance and mapping. It does not create or change a reporting workbook.')
    section('Balance totals', [f'{key.title()}: {Decimal(str(value)):,.2f}'
                               for key, value in report['totals'].items()])
    section('Items preventing workbook generation', report['write_blockers'])
    section('Review notes', report['warnings'])
    section('Accounts without a mapping', [f"{item['account']} - {item['description']}: {item['amount']}"
                                           for item in report['unmapped_accounts']])
    section('Reporting lines without a source amount',
            [f"Row {item['row']} - {item['account']} {item['description']} | Existing: {item['existing']}"
             for item in report['unresolved_destination_lines']])
    differences = [item for item in report['comparisons']
                   if item['difference'] is not None and Decimal(str(item['difference'])) != 0]
    section('Amounts different from the existing reporting month',
            [f"Row {item['row']} - {item['description']} | Calculated: {item['amount']} | "
             f"Existing: {item['existing']} | Difference: {item['difference']}" for item in differences])
    section('Amounts that could not be compared',
            [f"Row {item['row']} - {item['description']}: no saved numeric value to compare."
             for item in report['comparisons'] if item['difference'] is None])
    section('Mappings whose source account is absent this month', report['mapping_accounts_absent_from_source'])
    section('Reporting rows that moved', [f"Account {item['account']}: row {item['old_row']} moved to {item['new_row']}."
                                          for item in report['moved_destinations']])
    line(f"Mapped reporting lines: {len(report['comparisons'])}. Missing-amount policy: {report['missing_lines_policy']}.")
    with path.open('xb') as handle:
        SimpleDocTemplate(handle).build(content)
