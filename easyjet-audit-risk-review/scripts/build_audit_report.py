"""Rebuild the memo and workbook from the packaged easyJet audit project sources."""

import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
DOCX = BASE / 'docs' / 'easyjet-audit-risk-review.docx'
XLSX = BASE / 'data' / 'easyjet-audit-risk-review.xlsx'

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ET.register_namespace('w', W_NS)


def get_docx_sectpr(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    ns = {'w': W_NS}
    body = root.find('w:body', ns)
    sect = body.find('w:sectPr', ns)
    return ET.tostring(sect, encoding='unicode') if sect is not None else ''


def w_run(text: str, bold: bool = False, size: int | None = None, italic: bool = False) -> str:
    rpr = []
    if bold:
        rpr.append('<w:b/>')
    if italic:
        rpr.append('<w:i/>')
    if size is not None:
        rpr.append(f'<w:sz w:val="{size}"/>')
        rpr.append(f'<w:szCs w:val="{size}"/>')
    rpr_xml = f"<w:rPr>{''.join(rpr)}</w:rPr>" if rpr else ''
    return f'<w:r>{rpr_xml}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def w_para(text: str = '', bold: bool = False, size: int | None = None, italic: bool = False, before: int = 80, after: int = 80) -> str:
    return (
        f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}"/></w:pPr>'
        f'{w_run(text, bold=bold, size=size, italic=italic)}</w:p>'
    )


def w_multi_para(parts, before: int = 60, after: int = 60) -> str:
    runs = ''.join(w_run(**part) for part in parts)
    return f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}"/></w:pPr>{runs}</w:p>'


def w_table(headers, rows, widths=None) -> str:
    if widths is None:
        widths = [1800] * len(headers)
    borders = (
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="6" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="6" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
    )

    def cell(text, width, header=False):
        return (
            f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>'
            f'{w_run(text, bold=header, size=18 if header else 17)}</w:p></w:tc>'
        )

    xml = ['<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' + borders + '</w:tblPr>']
    xml.append('<w:tr>' + ''.join(cell(h, widths[i], header=True) for i, h in enumerate(headers)) + '</w:tr>')
    for row in rows:
        xml.append('<w:tr>' + ''.join(cell(str(val), widths[i], header=False) for i, val in enumerate(row)) + '</w:tr>')
    xml.append('</w:tbl>')
    return ''.join(xml)


sectpr_xml = get_docx_sectpr(DOCX)

risk_headers = [
    'Risk area',
    'Why the team would care',
    'Assertion / audit angle',
    'What could break',
    'Planning response',
]

risk_rows = [
    [
        'Revenue and customer contract liabilities',
        'IFRS 15 is applied across different triggers: passenger seats and most ancillaries at flown date, cancellation fees when processed, easyJet Plus over membership term, partner commissions net as agent, and Holidays non-flight elements over the stay. The balance sheet closed with GBP1.95bn of unearned revenue and GBP19m of other customer contract liabilities.',
        'Cut-off, completeness, accuracy, valuation and presentation; presumed fraud risk around revenue timing.',
        'Unearned revenue may be released too early or too late; cancelled flights may stay in unearned revenue instead of moving to refund or voucher liabilities; breakage may be recognised too aggressively; partner revenue may be grossed up incorrectly; Holidays revenue may be recognised at departure rather than over the holiday.',
        'Walk through booking, flown, cancellation and voucher flows; test key automated rules; reperform the contract-liability roll-forward; select flights and holidays spanning 30 September; inspect partner and insurance contracts for principal-agent; challenge breakage and compensation assumptions using post year-end data and disruption trends.',
    ],
    [
        'Leased aircraft maintenance provision',
        'PwC treated this as the FY25 Group key audit matter. The year-end provision of GBP939m is sensitive to utilisation, contract terms, uncontracted maintenance costs, inflation, discount rates and USD exposure.',
        'Valuation, completeness, accuracy.',
        'The provision could be understated if assumptions on heavy-maintenance timing, escalation or lease-return outcomes are too optimistic. Current versus non-current classification can also move materially.',
        'Obtain the model early; map major leased-aircraft populations; test flying hours and cycles to engineering records; compare cost assumptions to recent shop visits and third-party contracts; review lease terms for restoration versus penalty outcomes; challenge discounting and sensitivity analysis.',
    ],
    [
        'Customer claims, cancellations, ETS and other judgemental accruals',
        'The Audit Committee highlighted customer claims, disruption-related liabilities, ETS obligations and legal exposures. These balances can change quickly when operational disruption rises.',
        'Completeness, valuation, cut-off, presentation.',
        'Compensation, refunds, vouchers, ETS accruals or litigation provisions may be incomplete or based on stale assumptions. Operational events may not be reflected consistently across accruals, payables and disclosures.',
        'Use disruption and cancellation data to challenge accrual completeness; test post year-end settlements; inspect legal and claims correspondence; recalculate ETS accruals to emissions and allowance data; assess whether the same operational events have been captured consistently across all affected balances.',
    ],
    [
        'Carrying value of Airline assets',
        'Goodwill of GBP387m, landing rights of GBP155m and the wider Airline asset base depend on value-in-use assumptions in a commercially sensitive industry exposed to yields, fuel, FX, carbon and capacity.',
        'Valuation and disclosure.',
        'Management may rely on optimistic assumptions for route profitability, yields, cost pass-through or capacity growth, and downside scenarios may not be severe enough for a cyclical airline.',
        'Tie the model to Board-approved plans; compare revenue, yield and capacity assumptions to external market evidence and internal trading; involve valuation support on WACC; challenge whether carbon, ETS and SAF costs are reflected consistently; focus on downside scenarios that genuinely erode headroom.',
    ],
    [
        'Going concern and liquidity planning',
        'Going concern is not a formality in aviation. easyJet ended FY25 with GBP602m net cash and GBP4.8bn of available liquidity including a committed USD1.7bn RCF, but the resilience case still depends on downside assumptions and executable management actions.',
        'Disclosure, completeness, presentation.',
        'Downside assumptions may be too mild, or mitigations such as delaying aircraft deliveries or reducing capex may be assumed without enough evidence they are controllable in the required timeframe.',
        'Review the base case and severe downside model as a planning workstream, not a late-stage memo; test consistency with principal risks, fleet commitments and financing maturities; inspect facility agreements; challenge the practicality and timing of management actions; reassess whether post year-end events change the credibility of the downside case.',
    ],
]

analytical_headers = ['Metric', 'FY2024', 'FY2025', 'Why it matters in planning']
analytical_rows = [
    ['Group revenue growth', 'n/a', '8.6%', 'Supportive growth, but the audit read-across is whether it is supported by capacity, flown sectors and holidays expansion, with no strain showing up in year-end cut-off.'],
    ['Airline revenue growth', 'n/a', '6.0%', 'Core airline revenue grew more slowly than Holidays. That mix shift matters because the revenue pattern and customer remedies are different across the two streams.'],
    ['easyJet holidays revenue growth', 'n/a', '26.0%', 'Fast growth in a newer revenue stream usually justifies extra planning attention, even if the absolute balance is smaller than the airline ledger.'],
    ['Operating margin', '6.3%', '6.9%', 'Improving margin is positive, but the team still needs to understand whether the movement is driven by trading, mix, released provisions or estimation changes.'],
    ['Receivables days', '5.3 days', '4.7 days', 'Low days are expected because cash is usually collected before travel. A sharp movement would push the team back into accrued income, B2B balances and classification.'],
    ['Cash conversion', '2.5x', '2.7x', 'Strong cash conversion is normal for an advance-booking airline, but it also means revenue cannot be audited in isolation from contract liabilities.'],
    ['Closing customer contract liabilities as % of revenue', '19.1%', '19.5%', 'This is a better airline planning metric than inventory days or gross margin. It shows how much of the revenue cycle still sits on the balance sheet at year end.'],
    ['Revenue recognised from opening customer liabilities', 'GBP1.446bn', 'GBP1.703bn', 'A large part of current-year revenue was already carried on the balance sheet at the prior year end, which is why the roll-forward is a core audit bridge.'],
]

body_parts = [
    w_para('Audit Risk Review - easyJet plc FY2025', bold=True, size=28, before=80, after=120),
    w_para('Planning-style portfolio project for graduate roles in audit, accounting advisory, and transaction services.', italic=True, size=18, before=0, after=40),
    w_para('Based on easyJet plc Annual Report and Accounts 2025 and published FY2025 disclosures.', italic=True, size=18, before=0, after=120),
    w_para('1. Executive Summary', bold=True, size=24, before=160, after=80),
    w_para('- easyJet\'s revenue cycle is cash-before-service. Customers usually pay at booking, while revenue is recognised later when the flight takes place or, for easyJet holidays non-flight elements, over the holiday period. For planning, that means revenue risk is really a combined income statement and balance-sheet workstream.', size=19),
    w_para('- At 30 September 2025, easyJet reported GBP1.95bn of unearned revenue and GBP19m of other customer contract liabilities. The revenue workstream should therefore be planned jointly with IT audit and data analytics rather than treated as a standard substantive-only area.', size=19),
    w_para('- The three issues most likely to drive senior attention are: (1) revenue and customer liabilities, because IFRS 15 is applied across multiple triggers and a high-volume booking platform; (2) the leased aircraft maintenance provision, because the estimate is large and judgement heavy; and (3) asset carrying value and going concern, because the airline remains commercially exposed to yield, fuel, FX, disruption and capacity assumptions.', size=19),
    w_para('2. Key Audit Risks', bold=True, size=24, before=160, after=80),
    w_table(risk_headers, risk_rows, widths=[1400, 2600, 1600, 2500, 2900]),
    w_para('3. Analytical Review', bold=True, size=24, before=160, after=80),
    w_para('- For an airline, the most useful planning metrics are the ones that explain revenue timing, balance-sheet build-up and cash conversion. Retail-style ratios such as inventory days add little to the planning story here.', size=19),
    w_table(analytical_headers, analytical_rows, widths=[2200, 1200, 1200, 4400]),
    w_para('4. Revenue Recognition Under IFRS 15 - Planning Issues', bold=True, size=24, before=160, after=80),
    w_para('- Revenue is not one policy at easyJet. Passenger seats and most airline ancillary revenue are recognised at a point in time when the flight takes place; cancellation fees are recognised when the cancellation request is processed; easyJet Plus is recognised evenly over the membership term; partner revenue is recognised net as commission because easyJet acts as agent; travel insurance commission is recognised at booking; and easyJet holidays non-flight elements are recognised over time across the holiday.', size=19),
    w_para('- Timing is therefore the first real IFRS 15 issue. The event date that drives recognition is usually not the booking date or cash-receipt date. For passenger and most ancillary revenue it is the flown date; for holidays it is the passage of the holiday period; for membership income it is the membership term. The planning response needs to focus on whether system logic is aligned to those triggers.', size=19),
    w_para('- Variable consideration is also real here rather than theoretical. easyJet offsets delay and cancellation compensation against revenue up to the value of the related flight, with any excess taken to other costs. The outstanding liability is based on known eligible events, passengers impacted and expected claim rates. In IFRS 15 terms, management should only reduce liabilities or recognise breakage when it is highly probable there will not be a significant reversal later.', size=19),
    w_para('- Vouchers, unrequested refunds and flight-transfer options add a second layer of judgement. Once a flight is cancelled, the cash may need to move out of unearned revenue into other customer contract liabilities. Revenue can only be recognised from breakage when redemption or refund is remote. This is commercially sensitive because disruption levels, claim behaviour and local regulation can change quickly.', size=19),
    w_para('- Cut-off risk is heightened by scale and by the fact that the balance-sheet bridge is large. In FY25, easyJet deferred GBP11.019bn of revenue during the year and closed with GBP1.95bn of unearned revenue. It recognised GBP1.678bn of opening unearned revenue and GBP25m of opening other customer contract liabilities through the year. A mis-dated flight status, failed interface or wrong liability classification can therefore move a material amount of revenue between periods.', size=19),
    w_para('- There is also an important bridge issue in the disclosure: the contract-liability roll-forward is presented inclusive of airline passenger duty, whereas passenger revenue is recognised excluding APD. A realistic audit team would build that bridge explicitly rather than force a false tie between operational bookings, contract liabilities and the revenue note.', size=19),
    w_para('- For accounting advisory, the most credible review questions are: does each product have the right performance obligation and trigger; are partner and insurance arrangements correctly presented net rather than gross; is holidays revenue really being recognised over the stay; and is breakage constrained appropriately? For audit, the equivalent response is to test system configuration, reperform the contract-liability roll-forward, use population analytics around 30 September flights and holidays, inspect manual journals, and use post year-end claims and voucher redemption data to challenge estimates.', size=19),
    w_para('5. Audit Approach Summary', bold=True, size=24, before=160, after=80),
    w_para('- Front-load systems and data-flow understanding. For easyJet, planning quality depends heavily on how booking, departure-control, cancellation, voucher and finance systems interact.', size=19),
    w_para('- Treat the revenue workstream as an IT-enabled audit area. The team should identify the rules that release unearned revenue, create refund or voucher liabilities and recognise Holidays revenue over the stay, then decide where control testing is efficient and where substantive data work is still needed.', size=19),
    w_para('- Use commercial information in the challenge process. Revenue and estimates should be challenged against disruption experience, customer behaviour, Holidays growth, yield trends, fuel and FX pressure, and route maturity rather than in isolation.', size=19),
    w_para('- Keep professional scepticism on the balances management can influence through judgement: contract liabilities, breakage, compensation accruals, provisions, valuation models and year-end classification decisions.', size=19),
    w_para('6. Conclusion', bold=True, size=24, before=160, after=80),
    w_para('- The highest-risk area remains revenue and customer liabilities. It is the most pervasive issue on the file because it combines multiple IFRS 15 recognition triggers, large contract-liability balances, variable consideration, cut-off exposure and dependence on complex systems.', size=19),
    w_para('- The leased aircraft maintenance provision remains the most judgemental estimate, but revenue is the wider planning risk because it affects both reported performance and the credibility of the balance-sheet liabilities that support it.', size=19),
]

new_document_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:document xmlns:w="{W_NS}"><w:body>' + ''.join(body_parts) + sectpr_xml + '</w:body></w:document>'
)

# Write updated docx by replacing document.xml only.
with zipfile.ZipFile(DOCX, 'r') as zin:
    items = {name: zin.read(name) for name in zin.namelist() if name != 'word/document.xml'}
with zipfile.ZipFile(DOCX, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for name, data in items.items():
        zout.writestr(name, data)
    zout.writestr('word/document.xml', new_document_xml)

# Workbook helpers.
MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', MAIN_NS)


def col_letter(idx: int) -> str:
    result = ''
    while idx:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def inline_cell(ref: str, text: str, style: int = 0) -> str:
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'


def num_cell(ref: str, value, style: int = 0) -> str:
    return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'


def formula_cell(ref: str, formula: str, value, style: int = 0) -> str:
    return f'<c r="{ref}" s="{style}"><f>{escape(formula)}</f><v>{value}</v></c>'


def make_sheet(rows, widths=None):
    all_cells = []
    max_col = 1
    max_row = len(rows)
    for r_idx, row in enumerate(rows, start=1):
        cell_xml = []
        for c_idx, cell in enumerate(row, start=1):
            if cell is None:
                continue
            ref = f'{col_letter(c_idx)}{r_idx}'
            max_col = max(max_col, c_idx)
            ctype = cell.get('type', 'str')
            style = cell.get('style', 0)
            if ctype == 'str':
                cell_xml.append(inline_cell(ref, str(cell.get('value', '')), style))
            elif ctype == 'num':
                cell_xml.append(num_cell(ref, cell.get('value', 0), style))
            elif ctype == 'formula':
                cell_xml.append(formula_cell(ref, cell.get('formula', ''), cell.get('value', 0), style))
        all_cells.append(f'<row r="{r_idx}">{"".join(cell_xml)}</row>')
    dim = f'A1:{col_letter(max_col)}{max_row}'
    cols_xml = ''
    if widths:
        cols_xml = '<cols>' + ''.join(
            f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths, start=1)
        ) + '</cols>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{MAIN_NS}">'
        f'<dimension ref="{dim}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="18"/>'
        f'{cols_xml}<sheetData>{"".join(all_cells)}</sheetData></worksheet>'
    )

rows_overview = [
    [ {'value':'Audit Risk Review - easyJet plc FY2025', 'style':1}, None, None, None, None ],
    [ {'value':'Workbook purpose', 'style':2}, {'value':'Support schedules for the planning memo, with emphasis on IFRS 15 timing, variable consideration, cut-off and contract-liability movements.', 'style':0} ],
    [ {'value':'How a Big 4 team would use this workbook', 'style':2}, {'value':'(1) scope planning analytics; (2) map revenue streams to recognition triggers; (3) bridge contract liabilities to revenue release; and (4) tie narrative points back to the annual report.', 'style':0} ],
    [ {'value':'Key technical points added in this revision', 'style':2}, {'value':'Flown-date recognition, over-time holiday recognition, principal-versus-agent judgement, variable consideration from compensation and breakage, APD bridge, and year-end cut-off around customer liabilities.', 'style':0} ],
    [ {'value':'Source base', 'style':2}, {'value':'easyJet plc Annual Report and Accounts 2025 and the FY2025 investor reporting page.', 'style':0} ],
]

rows_analytics = [
    [ {'value':'Planning Analytics', 'style':1}, None, None, None, None ],
    [ {'value':'Metric', 'style':2}, {'value':'FY2024', 'style':2}, {'value':'FY2025', 'style':2}, {'value':'Movement / calc', 'style':2}, {'value':'Planning read-across', 'style':2} ],
    [ {'value':'Group revenue (GBPm)', 'style':0}, {'type':'num', 'value':9309, 'style':0}, {'type':'num', 'value':10106, 'style':0}, {'type':'num', 'value':797, 'style':0}, {'value':'Supportive growth, but planning still needs to confirm it is consistent with capacity and flown sectors.', 'style':0} ],
    [ {'value':'Group revenue growth', 'style':0}, None, {'type':'formula', 'formula':'(C3/B3)-1', 'value':0.0856160704694382, 'style':3}, None, {'value':'Useful high-level sense check; does not replace cut-off testing.', 'style':0} ],
    [ {'value':'Airline revenue (GBPm)', 'style':0}, {'type':'num', 'value':8172, 'style':0}, {'type':'num', 'value':8666, 'style':0}, {'type':'num', 'value':494, 'style':0}, {'value':'Core airline growth is slower than Holidays, which matters because the recognition patterns differ.', 'style':0} ],
    [ {'value':'Airline revenue growth', 'style':0}, None, {'type':'formula', 'formula':'(C5/B5)-1', 'value':0.06045031815956934, 'style':3}, None, {'value':'Mix shift means the team should not over-rely on prior-year revenue testing approaches.', 'style':0} ],
    [ {'value':'easyJet holidays revenue (GBPm)', 'style':0}, {'type':'num', 'value':1521, 'style':0}, {'type':'num', 'value':1917, 'style':0}, {'type':'num', 'value':396, 'style':0}, {'value':'Fast growth means disproportionate planning attention is justified.', 'style':0} ],
    [ {'value':'easyJet holidays growth', 'style':0}, None, {'type':'formula', 'formula':'(C7/B7)-1', 'value':0.2603546386587771, 'style':3}, None, {'value':'Over-time revenue recognition across the stay is a real technical issue here.', 'style':0} ],
    [ {'value':'Operating margin', 'style':0}, {'type':'formula', 'formula':'589/9309', 'value':0.06327210226662369, 'style':3}, {'type':'formula', 'formula':'696/10106', 'value':0.06886997823075401, 'style':3}, None, {'value':'Margin improvement should be challenged against trading, mix and estimation changes.', 'style':0} ],
    [ {'value':'Receivables days', 'style':0}, {'type':'formula', 'formula':'135/9309*365', 'value':5.293264582661941, 'style':0}, {'type':'formula', 'formula':'130/10106*365', 'value':4.695230556105284, 'style':0}, None, {'value':'Low days are expected because cash is generally collected before travel.', 'style':0} ],
    [ {'value':'Cash conversion (cash generated from operations / operating profit)', 'style':0}, {'type':'formula', 'formula':'1483/589', 'value':2.5178268251273344, 'style':0}, {'type':'formula', 'formula':'1875/696', 'value':2.6939655172413794, 'style':0}, None, {'value':'Revenue cannot be audited in isolation from contract liabilities.', 'style':0} ],
    [ {'value':'Closing customer contract liabilities (GBPm)', 'style':0}, {'type':'formula', 'formula':'1741+35', 'value':1776, 'style':0}, {'type':'formula', 'formula':'1950+19', 'value':1969, 'style':0}, {'type':'num', 'value':193, 'style':0}, {'value':'Large closing liabilities explain why balance-sheet testing is central to the revenue workstream.', 'style':0} ],
    [ {'value':'Closing customer contract liabilities as % of revenue', 'style':0}, {'type':'formula', 'formula':'(1741+35)/9309', 'value':0.19078203888709852, 'style':3}, {'type':'formula', 'formula':'(1950+19)/10106', 'value':0.19483475163269346, 'style':3}, None, {'value':'A better planning metric for an airline than inventory days or gross margin.', 'style':0} ],
    [ {'value':'Revenue recognised from opening customer liabilities (GBPm)', 'style':0}, {'type':'formula', 'formula':'1399+47', 'value':1446, 'style':0}, {'type':'formula', 'formula':'1678+25', 'value':1703, 'style':0}, {'type':'num', 'value':257, 'style':0}, {'value':'Shows how much of current-year revenue was already sitting on the prior-year balance sheet.', 'style':0} ],
    [ {'value':'Maintenance provision as % of revenue', 'style':0}, {'type':'formula', 'formula':'894/9309', 'value':0.09603609410248146, 'style':3}, {'type':'formula', 'formula':'939/10106', 'value':0.09291509994062933, 'style':3}, None, {'value':'Quick sense check of scale for the key judgemental estimate on the file.', 'style':0} ],
]

rows_risks = [
    [ {'value':'Risk Register', 'style':1}, None, None, None, None ],
    [ {'value':'Risk area', 'style':2}, {'value':'Why it matters', 'style':2}, {'value':'Assertion / angle', 'style':2}, {'value':'Failure mode', 'style':2}, {'value':'Planning response', 'style':2} ],
]
for row in risk_rows:
    rows_risks.append([{'value':row[i], 'style':0} for i in range(5)])

rows_revenue = [
    [ {'value':'Revenue Workstream - IFRS 15 technical focus', 'style':1}, None, None, None, None, None ],
    [ {'value':'Technical note', 'style':4}, {'value':'The contract-liability roll-forward is presented inclusive of airline passenger duty, while passenger revenue is recognised excluding APD. The planning file should therefore build an explicit APD bridge rather than forcing a direct tie.', 'style':4}, None, None, None, None ],
    [ {'value':'Revenue stream', 'style':2}, {'value':'FY2025 disclosed amount (GBPm)', 'style':2}, {'value':'IFRS 15 timing', 'style':2}, {'value':'Variable consideration / judgement', 'style':2}, {'value':'Main cut-off risk', 'style':2}, {'value':'Audit / advisory focus', 'style':2} ],
    [ {'value':'Passenger revenue', 'style':0}, {'type':'num', 'value':6072, 'style':0}, {'value':'Point in time when the flight takes place; no-shows recognised once the flight has departed.', 'style':0}, {'value':'Delay and cancellation compensation offsets revenue up to the related fare; liability uses known events, passengers impacted and claim-rate estimates.', 'style':0}, {'value':'Flights sold before 30 September but flown after it; cancelled flights moved to refund or voucher liabilities.', 'style':0}, {'value':'Test the flown-date trigger, contract-liability unwind, post year-end flights and compensation model inputs; inspect manual journals around period end.', 'style':0} ],
    [ {'value':'Airline ancillary revenue', 'style':0}, {'value':'Not separately disclosed', 'style':0}, {'value':'Generally point in time when the flight takes place; cancellation fees recognised when the cancellation request is processed.', 'style':0}, {'value':'Different ancillary products may have different recognition triggers, especially change and cancellation fees.', 'style':0}, {'value':'Ancillaries linked to year-end flights can be released in the wrong period if product mapping is wrong.', 'style':0}, {'value':'Walk through product mapping from booking system to ledger; test event-date logic by ancillary type.', 'style':0} ],
    [ {'value':'Partner revenue and in-flight sales', 'style':0}, {'value':'Not separately disclosed', 'style':0}, {'value':'Recognised net as commission because easyJet acts as agent; travel insurance commission recognised at booking.', 'style':0}, {'value':'Principal-versus-agent judgement and correct trigger for travel-insurance commission.', 'style':0}, {'value':'Gross presentation or wrong booking/service trigger could overstate revenue.', 'style':0}, {'value':'Inspect partner agreements, confirm who controls the service and who bears inventory risk, and compare gross cash with net revenue recognition.', 'style':0} ],
    [ {'value':'easyJet Plus', 'style':0}, {'value':'Not separately disclosed', 'style':0}, {'value':'Recognised evenly over the annual membership period.', 'style':0}, {'value':'Deferred income period, renewals and start-date logic.', 'style':0}, {'value':'Membership income released too quickly around year end.', 'style':0}, {'value':'Test member start and end dates and reperform straight-line release on a sample basis.', 'style':0} ],
    [ {'value':'easyJet holidays non-flight revenue', 'style':0}, {'type':'num', 'value':1917, 'style':0}, {'value':'Recognised over time across the holiday period; flight revenue is netted out of the holidays revenue line.', 'style':0}, {'value':'Refunds, vouchers and holidays spanning year end require judgement over the amount earned to date.', 'style':0}, {'value':'Revenue recognised at departure instead of over the stay, or cut-off based on booking date rather than days consumed.', 'style':0}, {'value':'Select bookings spanning 30 September and reperform apportionment over the stay; challenge policy application in system or manual schedules.', 'style':0} ],
    [ {'value':'Contract liability roll-forward', 'style':2}, None, None, None, None, None ],
    [ {'value':'Item', 'style':2}, {'value':'FY2024', 'style':2}, {'value':'FY2025', 'style':2}, {'value':'Technical note', 'style':2}, None, None ],
    [ {'value':'Opening unearned revenue', 'style':0}, {'type':'num', 'value':1501, 'style':0}, {'type':'num', 'value':1741, 'style':0}, {'value':'Opening balance released as flights and holidays are provided.', 'style':0} ],
    [ {'value':'Opening other customer contract liabilities', 'style':0}, {'type':'num', 'value':79, 'style':0}, {'type':'num', 'value':35, 'style':0}, {'value':'Mainly vouchers and amounts awaiting customer instruction after cancellations.', 'style':0} ],
    [ {'value':'Revenue deferred during the year', 'style':0}, {'type':'num', 'value':10170, 'style':0}, {'type':'num', 'value':11019, 'style':0}, {'value':'Inclusive of APD and other charges.', 'style':0} ],
    [ {'value':'Revenue recognised during the year - unearned revenue', 'style':0}, {'type':'num', 'value':-9219, 'style':0}, {'type':'num', 'value':-10045, 'style':0}, {'value':'Release of revenue as performance obligations are satisfied.', 'style':0} ],
    [ {'value':'Airline passenger duty on revenue recognised', 'style':0}, {'type':'num', 'value':-711, 'style':0}, {'type':'num', 'value':-765, 'style':0}, {'value':'Explains why the contract-liability roll-forward does not tie directly to the passenger revenue line.', 'style':0} ],
    [ {'value':'Additional other customer contract liability', 'style':0}, {'type':'num', 'value':187, 'style':0}, {'type':'num', 'value':186, 'style':0}, {'value':'Arises as flights are cancelled and customers elect vouchers or delay instruction.', 'style':0} ],
    [ {'value':'Reduction in other customer contract liability', 'style':0}, {'type':'num', 'value':-184, 'style':0}, {'type':'num', 'value':-184, 'style':0}, {'value':'Redemptions, refunds and other resolution of customer options.', 'style':0} ],
    [ {'value':'Foreign exchange impact', 'style':0}, {'type':'num', 'value':0, 'style':0}, {'type':'num', 'value':1, 'style':0}, {'value':'Minor movement in FY2025.', 'style':0} ],
    [ {'value':'Closing unearned revenue', 'style':0}, {'type':'num', 'value':1741, 'style':0}, {'type':'num', 'value':1950, 'style':0}, {'value':'Key year-end balance for cut-off and completeness testing.', 'style':0} ],
    [ {'value':'Closing other customer contract liabilities', 'style':0}, {'type':'num', 'value':35, 'style':0}, {'type':'num', 'value':19, 'style':0}, {'value':'Monitor classification between unearned revenue and other payables after cancellations.', 'style':0} ],
    [ {'value':'Revenue recognised from opening unearned revenue', 'style':0}, {'type':'num', 'value':1399, 'style':0}, {'type':'num', 'value':1678, 'style':0}, {'value':'Useful bridge for the revenue workstream.', 'style':0} ],
    [ {'value':'Revenue recognised from opening other customer contract liabilities', 'style':0}, {'type':'num', 'value':47, 'style':0}, {'type':'num', 'value':25, 'style':0}, {'value':'Shows how much opening voucher/refund liability cleared into revenue.', 'style':0} ],
]

rows_sources = [
    [ {'value':'Source Log', 'style':1}, None, None ],
    [ {'value':'Source', 'style':2}, {'value':'Why used', 'style':2}, {'value':'Relevant pages / lines', 'style':2} ],
    [ {'value':'easyJet Annual Report and Accounts 2025', 'style':0}, {'value':'Primary source for revenue policies, contract liabilities, provisions, liquidity and the auditor\'s key audit matter.', 'style':0}, {'value':'Revenue policy pages 155-156; contract liabilities pages 167-168; auditor report page 139.', 'style':0} ],
    [ {'value':'easyJet FY2025 annual report investor page', 'style':0}, {'value':'Official investor source confirming the FY2025 filing and supporting public presentation context.', 'style':0}, {'value':'Investor reports and presentations page for 2025 annual report.', 'style':0} ],
]

sheet_xmls = {
    'xl/worksheets/sheet1.xml': make_sheet(rows_overview, widths=[32, 110]),
    'xl/worksheets/sheet2.xml': make_sheet(rows_analytics, widths=[34, 14, 14, 18, 78]),
    'xl/worksheets/sheet3.xml': make_sheet(rows_risks, widths=[24, 40, 20, 40, 46]),
    'xl/worksheets/sheet4.xml': make_sheet(rows_revenue, widths=[28, 18, 28, 36, 28, 42]),
    'xl/worksheets/sheet5.xml': make_sheet(rows_sources, widths=[40, 56, 32]),
}

with zipfile.ZipFile(XLSX, 'r') as zin:
    items = {name: zin.read(name) for name in zin.namelist() if name not in sheet_xmls}
with zipfile.ZipFile(XLSX, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for name, data in items.items():
        zout.writestr(name, data)
    for name, xml in sheet_xmls.items():
        zout.writestr(name, xml)

print('Updated', DOCX.name, 'and', XLSX.name)
