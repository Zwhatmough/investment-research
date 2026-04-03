"""Rebuild the packaged easyJet audit report workbook and memo deliverables."""

from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

BASE = Path(__file__).resolve().parents[1]
DOCX = BASE / 'docs' / 'easyjet-audit-risk-review.docx'
XLSX = BASE / 'data' / 'easyjet-audit-risk-review.xlsx'

# -----------------------------
# Source data
# -----------------------------
raw_rows = [
    ('Revenue', 10106, 9309, 8171, 'AR25 p.144 / Note 8 pp.163-164; FY2023 comp AR24 corresponding Note 8', 'Statutory group revenue.'),
    ('Passenger revenue', 6072, 5715, 5221, 'AR25 p.144 / Note 8 pp.163-164; FY2023 comp AR24 corresponding Note 8', 'Primary airline revenue stream; useful for flown-date cut-off analysis.'),
    ('Ancillary revenue', 2594, 2457, 2174, 'AR25 p.144 / Note 8 pp.163-164; FY2023 comp AR24 corresponding Note 8', 'Includes airline ancillary revenue recognised with the flight or booking event as applicable.'),
    ('easyJet holidays revenue (net of flight revenue)', 1440, 1137, 776, 'AR25 p.144 / Note 8 pp.163-164; FY2023 comp AR24 corresponding Note 8', 'Group Holidays revenue after eliminating intercompany flight revenue.'),
    ('Airline revenue (passenger + ancillary)', 8666, 8172, 7395, 'Derived from passenger revenue + ancillary revenue (AR25 Note 8)', 'Used for mix commentary and airline-only growth analysis.'),
    ('Statutory EBITDA (derived)', 1439, 1359, 1126, 'Derived from operating profit plus depreciation and amortisation; AR25 p.174 Note 23 / AR24 p.187 Note 24', 'Derived subtotal used only as an operating cash proxy, not as a primary IFRS subtotal.'),
    ('Depreciation and amortisation', 743, 770, 673, 'AR25 p.174 Note 23; FY2023 comp AR24 p.187 Note 24', 'Important bridge between EBITDA and operating profit in an asset-intensive airline.'),
    ('Total operating costs before depreciation and amortisation', 8667, 7950, 7045, 'Derived from revenue less statutory EBITDA', 'More meaningful for easyJet than a retailer-style cost of sales subtotal.'),
    ('Fuel expense', 2253, 2223, 2033, 'AR25 p.144 income statement / p.33 financial review; FY2023 comp AR24 corresponding disclosures', 'Used for fuel cost intensity and commercial sensitivity analysis.'),
    ('Operating profit', 696, 589, 453, 'AR25 p.144 income statement / p.174 Note 23; FY2023 comp AR24 p.187 Note 24', 'Statutory operating profit.'),
    ('Trade receivables (gross)', 136, 142, 115, 'AR25 p.167 Note 13; FY2023 comp AR24 corresponding note', 'Before loss allowance.'),
    ('Loss allowance', 6, 7, 5, 'AR25 p.167 Note 13; FY2023 comp AR24 corresponding note', 'Used to arrive at net trade receivables.'),
    ('Trade receivables (net)', 130, 135, 110, 'AR25 p.167 Note 13; FY2023 comp AR24 corresponding note', 'Used for receivables days.'),
    ('Trade and other receivables', 530, 483, 343, 'AR25 p.167 Note 13 / p.146 statement of financial position; FY2023 comp AR24 corresponding note', 'Current receivables line on the face of the balance sheet.'),
    ('Trade payables', 366, 357, 402, 'AR25 p.167 Note 15; FY2023 comp AR24 corresponding note', 'Used for payables days.'),
    ('Trade and other payables', 1654, 1656, 1764, 'AR25 p.167 Note 15 / p.146 statement of financial position; FY2023 comp AR24 corresponding note', 'Total current payables; trade payables are used for days calculations.'),
    ('Current assets', 4625, 4545, 4130, 'AR25 p.146 statement of financial position; FY2023 comp AR24 p.156', 'Includes current intangible assets and short-dated investments.'),
    ('Current liabilities', 4152, 4471, 4144, 'AR25 p.146 statement of financial position; FY2023 comp AR24 p.156', 'Includes customer balances, borrowings, lease liabilities and provisions.'),
    ('Current intangible assets', 518, 572, 676, 'AR25 p.146 statement of financial position / Note 10 p.165; FY2023 comp AR24 corresponding note', 'Mainly ETS and carbon-related assets; excluded from adjusted liquidity metrics.'),
    ('Cash generated from operations', 1875, 1483, 1509, 'AR25 p.174 Note 23; FY2023 comp AR24 p.187 Note 24', 'Used for operating cash conversion.'),
    ('Net cash generated from operating activities', 1625, 1465, 1551, 'AR25 p.148 cash flow statement; FY2023 comp AR24 p.158', 'Used for operating cash flow to revenue.'),
    ('Cash and other investments', 3528, 3461, 2925, 'AR25 p.167 Note 14; FY2023 comp AR24 corresponding note', 'Cash, cash equivalents and other investments.'),
    ('Borrowings', 1881, 2106, 1895, 'AR25 p.168 Note 17; FY2023 comp AR24 corresponding note', 'Current plus non-current borrowings.'),
    ('Lease liabilities', 1045, 1174, 989, 'AR25 p.169 Note 18; FY2023 comp AR24 corresponding note', 'Current plus non-current lease liabilities.'),
    ('Funding obligations (borrowings + leases)', 2926, 3280, 2884, 'Derived from borrowings + lease liabilities', 'Useful capital structure subtotal in an airline context.'),
    ('Adjusted net cash/(debt) incl. leases', 602, 181, 41, 'Derived from cash and other investments less borrowings and lease liabilities; strategic review cross-check AR25 p.5', 'Positive values indicate a net cash position after including lease liabilities.'),
    ('Unearned revenue', 1950, 1741, 1501, 'AR25 pp.167-168 Note 16 / p.146 statement of financial position; FY2023 comp AR24 corresponding note', 'Current and non-current unearned revenue combined.'),
    ('Other customer contract liabilities', 19, 35, 79, 'AR25 pp.167-168 Note 16; FY2023 comp AR24 corresponding note', 'Customer vouchers and unresolved refund / transfer options after cancellations.'),
    ('Total customer contract liabilities', 1969, 1776, 1580, 'Derived from unearned revenue + other customer contract liabilities (AR25 Note 16)', 'Full customer-balance exposure beneath reported revenue.'),
    ('Right-of-use assets', 1015, 1190, 928, 'AR25 pp.168-169 Note 18; FY2023 comp AR24 corresponding note', 'Material lease-related asset base.'),
    ('Property, plant and equipment (excluding ROU assets)', 4791, 4285, 3936, 'AR25 p.146 statement of financial position / Note 11 pp.166-167; FY2023 comp AR24 corresponding note', 'Owned aircraft, spares and other owned assets.'),
    ('Goodwill', 387, 387, 365, 'AR25 p.165 Note 10; FY2023 comp AR24 corresponding note', 'Relevant to impairment planning.'),
    ('Other intangible assets', 384, 406, 276, 'AR25 p.165 Note 10; FY2023 comp AR24 corresponding note', 'Excludes goodwill; includes landing rights and software.'),
    ('Maintenance provision', 939, 894, 753, 'AR25 pp.169-170 Note 19; FY2023 comp AR24 p.181', 'Largest estimation area in provisions and a key audit matter.'),
    ('Total equity', 3498, 2973, 2787, 'AR25 p.149 statement of changes in equity / p.146 net assets; FY2023 comp AR24 corresponding statement', 'Used for leverage and capital structure ratios.'),
    ('Revenue recognised from opening unearned revenue', 1678, 1399, 1006, 'AR25 p.168 Note 16; FY2023 comp AR24 corresponding note', 'Useful bridge from opening balance sheet to current-year revenue.'),
    ('Revenue recognised from opening other customer liabilities', 25, 47, 47, 'AR25 p.168 Note 16; FY2023 comp AR24 corresponding note', 'Supports understanding of voucher / refund liability unwind.'),
    ('Revenue deferred during the year', 11019, 10170, 9233, 'AR25 p.167 Note 16; FY2023 comp AR24 corresponding note', 'Inclusive of APD and other charges.'),
    ('APD on revenue recognised', 765, 711, 'n/a', 'AR25 p.167 Note 16; FY2024 comp AR24 corresponding note', 'Required bridge between contract-liability movement and passenger revenue.'),
]

raw_index = {name: i + 3 for i, (name, *_rest) in enumerate(raw_rows)}

def cell_ref(sheet, col, row):
    return f"{sheet}!{col}{row}"

ratio_rows = [
    {
        'metric': 'Revenue growth',
        'formula_text': '(Current-year revenue / Prior-year revenue) - 1',
        'line_items': 'Revenue (group)',
        'why': 'Standard top-line growth measure using statutory revenue.',
        'fy24_formula': f"=({cell_ref('Raw_Financials','C',raw_index['Revenue'])}/{cell_ref('Raw_Financials','D',raw_index['Revenue'])})-1",
        'fy25_formula': f"=({cell_ref('Raw_Financials','B',raw_index['Revenue'])}/{cell_ref('Raw_Financials','C',raw_index['Revenue'])})-1",
        'fy24_value': 9309/8171-1,
        'fy25_value': 10106/9309-1,
        'style': 'pct',
        'audit': 'Growth remains positive but slowed; the more important audit question is whether growth translates cleanly through cut-off and customer liability movements.'
    },
    {
        'metric': 'Operating margin',
        'formula_text': 'Operating profit / Revenue',
        'line_items': 'Operating profit; Revenue',
        'why': 'Most relevant statutory profitability measure for audit planning.',
        'fy24_formula': f"={cell_ref('Raw_Financials','C',raw_index['Operating profit'])}/{cell_ref('Raw_Financials','C',raw_index['Revenue'])}",
        'fy25_formula': f"={cell_ref('Raw_Financials','B',raw_index['Operating profit'])}/{cell_ref('Raw_Financials','B',raw_index['Revenue'])}",
        'fy24_value': 589/9309,
        'fy25_value': 696/10106,
        'style': 'pct',
        'audit': 'Margin improved, but this should be reconciled to the movement in cash generation, D&A and customer liability releases rather than accepted at face value.'
    },
    {
        'metric': 'EBITDA margin',
        'formula_text': 'EBITDA / Revenue',
        'line_items': 'Statutory EBITDA (derived); Revenue',
        'why': 'Useful operating cash proxy, but clearly identified as a derived subtotal rather than a primary IFRS line item.',
        'fy24_formula': f"={cell_ref('Raw_Financials','C',raw_index['Statutory EBITDA (derived)'])}/{cell_ref('Raw_Financials','C',raw_index['Revenue'])}",
        'fy25_formula': f"={cell_ref('Raw_Financials','B',raw_index['Statutory EBITDA (derived)'])}/{cell_ref('Raw_Financials','B',raw_index['Revenue'])}",
        'fy24_value': 1359/9309,
        'fy25_value': 1439/10106,
        'style': 'pct',
        'audit': 'EBITDA margin softened slightly despite higher EBIT margin, suggesting part of the statutory profit improvement sits in depreciation / amortisation and asset mix rather than underlying cash margin.'
    },
    {
        'metric': 'Depreciation and amortisation as % of revenue',
        'formula_text': 'Depreciation and amortisation / Revenue',
        'line_items': 'Depreciation and amortisation; Revenue',
        'why': 'Useful bridge between EBITDA and EBIT in an asset-intensive airline with significant fleet and lease accounting effects.',
        'fy24_formula': f"={cell_ref('Raw_Financials','C',raw_index['Depreciation and amortisation'])}/{cell_ref('Raw_Financials','C',raw_index['Revenue'])}",
        'fy25_formula': f"={cell_ref('Raw_Financials','B',raw_index['Depreciation and amortisation'])}/{cell_ref('Raw_Financials','B',raw_index['Revenue'])}",
        'fy24_value': 770/9309,
        'fy25_value': 743/10106,
        'style': 'pct',
        'audit': 'Depreciation and amortisation fell as a share of revenue, which explains part of the EBIT versus EBITDA divergence and supports challenge over useful lives, fleet mix and capitalisation judgements.'
    },
    {
        'metric': 'Current ratio',
        'formula_text': 'Current assets / Current liabilities',
        'line_items': 'Current assets; Current liabilities',
        'why': 'Standard liquidity measure using statement of financial position totals.',
        'fy24_formula': f"={cell_ref('Raw_Financials','C',raw_index['Current assets'])}/{cell_ref('Raw_Financials','C',raw_index['Current liabilities'])}",
        'fy25_formula': f"={cell_ref('Raw_Financials','B',raw_index['Current assets'])}/{cell_ref('Raw_Financials','B',raw_index['Current liabilities'])}",
        'fy24_value': 4545/4471,
        'fy25_value': 4625/4152,
        'style': 'num',
        'audit': 'Year-end liquidity improved, but this ratio should not be read in isolation because airline current liabilities include sizeable customer prepayments that are not cash obligations in the same way as trade payables or debt.'
    },
    {
        'metric': 'Quick ratio (adjusted)',
        'formula_text': '(Current assets - Current intangible assets) / Current liabilities',
        'line_items': 'Current assets; Current intangible assets; Current liabilities',
        'why': 'easyJet has no separate inventory line; excluding current intangible assets gives a better near-term liquidity view than using the statutory current ratio alone.',
        'fy24_formula': f"=({cell_ref('Raw_Financials','C',raw_index['Current assets'])}-{cell_ref('Raw_Financials','C',raw_index['Current intangible assets'])})/{cell_ref('Raw_Financials','C',raw_index['Current liabilities'])}",
        'fy25_formula': f"=({cell_ref('Raw_Financials','B',raw_index['Current assets'])}-{cell_ref('Raw_Financials','B',raw_index['Current intangible assets'])})/{cell_ref('Raw_Financials','B',raw_index['Current liabilities'])}",
        'fy24_value': (4545-572)/4471,
        'fy25_value': (4625-518)/4152,
        'style': 'num',
        'audit': 'Adjusted quick liquidity is still only around parity with current liabilities, but the denominator still includes customer prepayments; that is why a customer-balance-adjusted cover metric is also useful.'
    },
    {
        'metric': 'Liquidity cover excluding customer balances',
        'formula_text': '(Current assets - Current intangible assets) / (Current liabilities - Total customer contract liabilities)',
        'line_items': 'Current assets; Current intangible assets; Current liabilities; Total customer contract liabilities',
        'why': 'This is more decision-useful for an airline because it strips out customer prepayments, which inflate statutory current liabilities but are not equivalent to supplier, tax or financing cash outflows.',
        'fy24_formula': f"=({cell_ref('Raw_Financials','C',raw_index['Current assets'])}-{cell_ref('Raw_Financials','C',raw_index['Current intangible assets'])})/({cell_ref('Raw_Financials','C',raw_index['Current liabilities'])}-{cell_ref('Raw_Financials','C',raw_index['Total customer contract liabilities'])})",
        'fy25_formula': f"=({cell_ref('Raw_Financials','B',raw_index['Current assets'])}-{cell_ref('Raw_Financials','B',raw_index['Current intangible assets'])})/({cell_ref('Raw_Financials','B',raw_index['Current liabilities'])}-{cell_ref('Raw_Financials','B',raw_index['Total customer contract liabilities'])})",
        'fy24_value': (4545-572)/(4471-1776),
        'fy25_value': (4625-518)/(4152-1969),
        'style': 'num',
        'audit': 'This adjusted view improved materially and gives a more realistic picture of short-term liquidity than the statutory current ratio alone.'
    },
    {
        'metric': 'Operating cash conversion',
        'formula_text': 'Cash generated from operations / Operating profit',
        'line_items': 'Cash generated from operations; Operating profit',
        'why': 'Useful audit planning bridge from earnings to operating cash before financing and tax flows.',
        'fy24_formula': f"={cell_ref('Raw_Financials','C',raw_index['Cash generated from operations'])}/{cell_ref('Raw_Financials','C',raw_index['Operating profit'])}",
        'fy25_formula': f"={cell_ref('Raw_Financials','B',raw_index['Cash generated from operations'])}/{cell_ref('Raw_Financials','B',raw_index['Operating profit'])}",
        'fy24_value': 1483/589,
        'fy25_value': 1875/696,
        'style': 'num',
        'audit': 'Very strong conversion is consistent with ticket sales in advance of travel, but it increases the importance of auditing customer liabilities and period-end release of revenue.'
    },
    {
        'metric': 'Gross funding obligations to equity (incl. leases)',
        'formula_text': '(Borrowings + Lease liabilities) / Total equity',
        'line_items': 'Borrowings; Lease liabilities; Total equity',
        'why': 'For an airline, lease liabilities are economically material funding obligations and should be considered alongside borrowings.',
        'fy24_formula': f"=({cell_ref('Raw_Financials','C',raw_index['Borrowings'])}+{cell_ref('Raw_Financials','C',raw_index['Lease liabilities'])})/{cell_ref('Raw_Financials','C',raw_index['Total equity'])}",
        'fy25_formula': f"=({cell_ref('Raw_Financials','B',raw_index['Borrowings'])}+{cell_ref('Raw_Financials','B',raw_index['Lease liabilities'])})/{cell_ref('Raw_Financials','B',raw_index['Total equity'])}",
        'fy24_value': (2106+1174)/2973,
        'fy25_value': (1881+1045)/3498,
        'style': 'num',
        'audit': 'Leverage has improved, but the balance sheet remains operationally leveraged and forecast-sensitive because leases remain embedded in the model.'
    },
    {
        'metric': 'Adjusted net debt to equity (incl. leases)',
        'formula_text': '(Borrowings + Lease liabilities - Cash and other investments) / Total equity',
        'line_items': 'Borrowings; Lease liabilities; Cash and other investments; Total equity',
        'why': 'Reflects easyJet’s disclosed net cash / debt concept while making the lease treatment explicit.',
        'fy24_formula': f"=({cell_ref('Raw_Financials','C',raw_index['Borrowings'])}+{cell_ref('Raw_Financials','C',raw_index['Lease liabilities'])}-{cell_ref('Raw_Financials','C',raw_index['Cash and other investments'])})/{cell_ref('Raw_Financials','C',raw_index['Total equity'])}",
        'fy25_formula': f"=({cell_ref('Raw_Financials','B',raw_index['Borrowings'])}+{cell_ref('Raw_Financials','B',raw_index['Lease liabilities'])}-{cell_ref('Raw_Financials','B',raw_index['Cash and other investments'])})/{cell_ref('Raw_Financials','B',raw_index['Total equity'])}",
        'fy24_value': (2106+1174-3461)/2973,
        'fy25_value': (1881+1045-3528)/3498,
        'style': 'num',
        'audit': 'The net cash position strengthened materially. That improves resilience, but does not remove the need to challenge downside scenarios in an inherently volatile sector.'
    },
    {
        'metric': 'Receivables days',
        'formula_text': 'Average net trade receivables / Revenue x 365',
        'line_items': 'Average net trade receivables; Revenue',
        'why': 'Average balance is more appropriate than closing balance because easyJet’s year-end position can be distorted by seasonality and booking timing.',
        'fy24_formula': f"=((({cell_ref('Raw_Financials','D',raw_index['Trade receivables (net)'])}+{cell_ref('Raw_Financials','C',raw_index['Trade receivables (net)'])})/2)/{cell_ref('Raw_Financials','C',raw_index['Revenue'])})*365",
        'fy25_formula': f"=((({cell_ref('Raw_Financials','C',raw_index['Trade receivables (net)'])}+{cell_ref('Raw_Financials','B',raw_index['Trade receivables (net)'])})/2)/{cell_ref('Raw_Financials','B',raw_index['Revenue'])})*365",
        'fy24_value': ((110+135)/2)/9309*365,
        'fy25_value': ((135+130)/2)/10106*365,
        'style': 'num',
        'audit': 'Receivables days are low and stable, which is expected in an advance-booking airline model. That supports the view that working capital risk sits more in customer liabilities than in collection risk.'
    },
    {
        'metric': 'Payables days',
        'formula_text': 'Average trade payables / Total operating costs before D&A x 365',
        'line_items': 'Average trade payables; total operating costs before depreciation and amortisation',
        'why': 'Trade payables are better matched to external operating cost spend than to revenue in an airline model.',
        'fy24_formula': f"=((({cell_ref('Raw_Financials','D',raw_index['Trade payables'])}+{cell_ref('Raw_Financials','C',raw_index['Trade payables'])})/2)/{cell_ref('Raw_Financials','C',raw_index['Total operating costs before depreciation and amortisation'])})*365",
        'fy25_formula': f"=((({cell_ref('Raw_Financials','C',raw_index['Trade payables'])}+{cell_ref('Raw_Financials','B',raw_index['Trade payables'])})/2)/{cell_ref('Raw_Financials','B',raw_index['Total operating costs before depreciation and amortisation'])})*365",
        'fy24_value': ((402+357)/2)/7950*365,
        'fy25_value': ((357+366)/2)/8667*365,
        'style': 'num',
        'audit': 'Payables days shortened, which may indicate less supplier-financing support or different year-end settlement timing; either way, customer prepayments remain the more important working capital driver.'
    },
    {
        'metric': 'Unearned revenue as % of revenue',
        'formula_text': 'Closing unearned revenue / Revenue',
        'line_items': 'Unearned revenue; Revenue',
        'why': 'Direct measure of the size of cash received in advance of travel that remains on the balance sheet at year end.',
        'fy24_formula': f"={cell_ref('Raw_Financials','C',raw_index['Unearned revenue'])}/{cell_ref('Raw_Financials','C',raw_index['Revenue'])}",
        'fy25_formula': f"={cell_ref('Raw_Financials','B',raw_index['Unearned revenue'])}/{cell_ref('Raw_Financials','B',raw_index['Revenue'])}",
        'fy24_value': 1741/9309,
        'fy25_value': 1950/10106,
        'style': 'pct',
        'audit': 'The increase reinforces that revenue cut-off and completeness of customer liabilities are core audit risks for easyJet.'
    },
    {
        'metric': 'Customer contract liabilities as % of revenue',
        'formula_text': '(Unearned revenue + Other customer contract liabilities) / Revenue',
        'line_items': 'Unearned revenue; Other customer contract liabilities; Revenue',
        'why': 'Captures the full customer liability position rather than unearned revenue alone.',
        'fy24_formula': f"=({cell_ref('Raw_Financials','C',raw_index['Unearned revenue'])}+{cell_ref('Raw_Financials','C',raw_index['Other customer contract liabilities'])})/{cell_ref('Raw_Financials','C',raw_index['Revenue'])}",
        'fy25_formula': f"=({cell_ref('Raw_Financials','B',raw_index['Unearned revenue'])}+{cell_ref('Raw_Financials','B',raw_index['Other customer contract liabilities'])})/{cell_ref('Raw_Financials','B',raw_index['Revenue'])}",
        'fy24_value': (1741+35)/9309,
        'fy25_value': (1950+19)/10106,
        'style': 'pct',
        'audit': 'The balance remains close to one-fifth of annual revenue. That is the strongest balance-sheet signal supporting the revenue recognition and cut-off risk assessment.'
    },
    {
        'metric': 'Lease liabilities to total capital',
        'formula_text': 'Lease liabilities / (Lease liabilities + Borrowings + Equity)',
        'line_items': 'Lease liabilities; Borrowings; Total equity',
        'why': 'Useful capital structure metric in an airline with significant lease funding.',
        'fy24_formula': f"={cell_ref('Raw_Financials','C',raw_index['Lease liabilities'])}/({cell_ref('Raw_Financials','C',raw_index['Lease liabilities'])}+{cell_ref('Raw_Financials','C',raw_index['Borrowings'])}+{cell_ref('Raw_Financials','C',raw_index['Total equity'])})",
        'fy25_formula': f"={cell_ref('Raw_Financials','B',raw_index['Lease liabilities'])}/({cell_ref('Raw_Financials','B',raw_index['Lease liabilities'])}+{cell_ref('Raw_Financials','B',raw_index['Borrowings'])}+{cell_ref('Raw_Financials','B',raw_index['Total equity'])})",
        'fy24_value': 1174/(1174+2106+2973),
        'fy25_value': 1045/(1045+1881+3498),
        'style': 'pct',
        'audit': 'Lease dependence eased, but remains material enough to influence going concern, impairment and maintenance provisioning workstreams.'
    },
    {
        'metric': 'Operating cash flow to revenue',
        'formula_text': 'Net cash generated from operating activities / Revenue',
        'line_items': 'Net cash generated from operating activities; Revenue',
        'why': 'Direct statutory cash flow measure using the cash flow statement.',
        'fy24_formula': f"={cell_ref('Raw_Financials','C',raw_index['Net cash generated from operating activities'])}/{cell_ref('Raw_Financials','C',raw_index['Revenue'])}",
        'fy25_formula': f"={cell_ref('Raw_Financials','B',raw_index['Net cash generated from operating activities'])}/{cell_ref('Raw_Financials','B',raw_index['Revenue'])}",
        'fy24_value': 1465/9309,
        'fy25_value': 1625/10106,
        'style': 'pct',
        'audit': 'Cash generation against revenue improved modestly, consistent with the resilience of the prepayment model and tighter liquidity management.'
    },
    {
        'metric': 'Fuel cost intensity',
        'formula_text': 'Fuel expense / Revenue',
        'line_items': 'Fuel expense; Revenue',
        'why': 'Fuel is a major commercial sensitivity and an important bridge between revenue growth and margin.',
        'fy24_formula': f"={cell_ref('Raw_Financials','C',raw_index['Fuel expense'])}/{cell_ref('Raw_Financials','C',raw_index['Revenue'])}",
        'fy25_formula': f"={cell_ref('Raw_Financials','B',raw_index['Fuel expense'])}/{cell_ref('Raw_Financials','B',raw_index['Revenue'])}",
        'fy24_value': 2223/9309,
        'fy25_value': 2253/10106,
        'style': 'pct',
        'audit': 'Fuel remained a large cost base but eased as a share of revenue, helping the trading outcome while leaving forecast assumptions sensitive to fuel price and hedge movements.'
    },
    {
        'metric': 'Maintenance provision as % of revenue',
        'formula_text': 'Maintenance provision / Revenue',
        'line_items': 'Maintenance provision; Revenue',
        'why': 'Useful scale check for the main judgemental provision on the file.',
        'fy24_formula': f"={cell_ref('Raw_Financials','C',raw_index['Maintenance provision'])}/{cell_ref('Raw_Financials','C',raw_index['Revenue'])}",
        'fy25_formula': f"={cell_ref('Raw_Financials','B',raw_index['Maintenance provision'])}/{cell_ref('Raw_Financials','B',raw_index['Revenue'])}",
        'fy24_value': 894/9309,
        'fy25_value': 939/10106,
        'style': 'pct',
        'audit': 'The provision remains very large relative to the trading base even though it fell slightly as a percentage of revenue, which supports continued specialist focus.'
    },
]

# -----------------------------
# Formatting helpers
# -----------------------------
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ET.register_namespace('w', W_NS)


def fmt_int(v):
    if isinstance(v, str):
        return v
    return f"{v:,.0f}"


def fmt_pct(v):
    return f"{v*100:.1f}%"


def fmt_ppt(v):
    sign = '+' if v >= 0 else ''
    return f"{sign}{v*100:.1f} ppts"


def fmt_ratio(v):
    sign = '-' if v < 0 else ''
    return f"{v:.2f}x"


def fmt_days(v):
    sign = '+' if v >= 0 else ''
    return f"{v:.2f} days" if sign == '' else f"{v:.2f} days"


def format_ratio_value(style, v):
    if style == 'pct':
        return fmt_pct(v)
    return fmt_ratio(v) if abs(v) < 10 else f"{v:.2f}"


def format_movement(style, fy24, fy25):
    diff = fy25 - fy24
    if style == 'pct':
        return fmt_ppt(diff)
    if style == 'num':
        if abs(fy24) > 1.5 or abs(fy25) > 1.5:
            return f"{diff:+.2f}"
        return f"{diff:+.2f}x"
    return f"{diff:+.2f}"


def w_run(text: str, *, bold=False, italic=False, size=None):
    props = []
    if bold:
        props.append('<w:b/>')
    if italic:
        props.append('<w:i/>')
    if size is not None:
        props.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    rpr = f'<w:rPr>{"".join(props)}</w:rPr>' if props else ''
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def w_para(text: str, *, bold=False, italic=False, size=None, before=60, after=60):
    return f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}"/></w:pPr>{w_run(text, bold=bold, italic=italic, size=size)}</w:p>'


def w_table(headers, rows, widths, *, font_size=16):
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
            f'<w:p><w:pPr><w:spacing w:before="20" w:after="20"/></w:pPr>'
            f'{w_run(str(text), bold=header, size=18 if header else font_size)}</w:p></w:tc>'
        )
    xml = ['<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' + borders + '</w:tblPr>']
    xml.append('<w:tr>' + ''.join(cell(h, widths[i], True) for i, h in enumerate(headers)) + '</w:tr>')
    for row in rows:
        xml.append('<w:tr>' + ''.join(cell(row[i], widths[i], False) for i in range(len(headers))) + '</w:tr>')
    xml.append('</w:tbl>')
    return ''.join(xml)


# -----------------------------
# Build document.xml
# -----------------------------
with ZipFile(DOCX, 'r') as z:
    doc_items = {name: z.read(name) for name in z.namelist() if name != 'word/document.xml'}
    root = ET.fromstring(z.read('word/document.xml'))
    ns = {'w': W_NS}
    sectpr = root.find('.//w:sectPr', ns)
    sectpr_xml = ET.tostring(sectpr, encoding='unicode') if sectpr is not None else ''

analytical_rows_doc = [
    ['Revenue growth', '13.9%', '8.6%', '-5.4 ppts', 'Growth remained positive, but slowed versus FY2024. The more relevant audit issue is that year-end customer liabilities still represented almost one-fifth of revenue.'],
    ['Operating margin', '6.3%', '6.9%', '+0.6 ppts', 'Statutory EBIT margin improved. The team should still bridge this to cash generation, D&A and estimate movements rather than relying on the headline improvement.'],
    ['EBITDA margin', '14.6%', '14.2%', '-0.4 ppts', 'EBITDA margin eased despite higher EBIT margin, which implies some of the profit improvement is explained by lower D&A or asset mix rather than pure trading leverage.'],
    ['Depreciation and amortisation as % of revenue', '8.3%', '7.4%', '-0.9 ppts', 'This explains part of the EBIT versus EBITDA divergence and supports a more focused challenge over asset lives, lease mix and capitalisation judgements.'],
    ['Adjusted quick ratio', '0.89x', '0.99x', '+0.10x', 'Liquidity improved, but this ratio still includes customer prepayments within current liabilities and should not be read in isolation.'],
    ['Liquidity cover excluding customer balances', '1.47x', '1.88x', '+0.41x', 'Stripping out customer contract liabilities gives a more decision-useful short-term liquidity view for an airline and shows a stronger year-end position than the statutory current ratio alone.'],
    ['Operating cash conversion', '2.52x', '2.69x', '+0.18x', 'Cash conversion remained very strong. That is consistent with cash being received in advance of travel and reinforces the need to audit revenue and contract liabilities together.'],
    ['Customer contract liabilities as % of revenue', '19.1%', '19.5%', '+0.4 ppts', 'The balance-sheet build-up increased. That is the clearest data point supporting the revenue recognition and cut-off risk assessment.'],
    ['Receivables days', '4.80', '4.79', '-0.01', 'Receivables days are low and stable. Collection risk is not the key working capital issue; customer prepayments are.'],
    ['Payables days', '17.42', '15.22', '-2.20', 'Payables days shortened, which may indicate less supplier financing support or different settlement timing; either way, customer cash remains the more important working capital driver.'],
    ['Fuel cost intensity', '23.9%', '22.3%', '-1.6 ppts', 'Fuel remained a major cost driver but eased as a share of revenue, supporting margins while leaving forecasts sensitive to fuel and hedge assumptions.'],
]

contract_bridge_rows_doc = [
    ['Opening total customer contract liabilities', '1,580', '1,776', 'Opening deferred revenue base to be released as travel occurs or options are exercised.'],
    ['Revenue deferred during the year', '10,170', '11,019', 'Cash received in advance of travel and other performance obligations.'],
    ['Revenue recognised during the year', '(9,266)', '(10,064)', 'Includes unearned revenue and other customer liabilities released to revenue.'],
    ['APD on revenue recognised', '(711)', '(765)', 'Technical bridge required because contract-liability movements are shown inclusive of APD.'],
    ['Net transfer / FX movement in other customer liabilities', '3', '3', 'Reflects additional versus reduced voucher/refund liabilities plus FX.'],
    ['Closing total customer contract liabilities', '1,776', '1,969', 'Reconciles to the year-end balance sheet exposure beneath revenue.'],
]

raw_rows_doc = []
for name, fy25, fy24, fy23, source, comment in raw_rows:
    raw_rows_doc.append([name, fmt_int(fy25), fmt_int(fy24), fmt_int(fy23), source, comment])

ratio_rows_doc = []
for row in ratio_rows:
    fy24 = row['fy24_value']
    fy25 = row['fy25_value']
    if row['style'] == 'pct':
        fy24_s = fmt_pct(fy24)
        fy25_s = fmt_pct(fy25)
        move_s = fmt_ppt(fy25 - fy24)
    else:
        fy24_s = f"{fy24:.2f}x" if 'days' not in row['metric'].lower() else f"{fy24:.2f} days"
        fy25_s = f"{fy25:.2f}x" if 'days' not in row['metric'].lower() else f"{fy25:.2f} days"
        if 'days' in row['metric'].lower():
            move_s = f"{fy25 - fy24:+.2f} days"
        else:
            move_s = f"{fy25 - fy24:+.2f}x"
    ratio_rows_doc.append([
        row['metric'],
        row['formula_text'],
        row['line_items'],
        row['why'],
        fy24_s,
        fy25_s,
        move_s,
        row['audit'],
    ])

body_parts = [
    w_para('Audit Risk Review - easyJet plc FY2025', bold=True, size=28, before=80, after=120),
    w_para('Audit planning memorandum for graduate audit, accounting advisory and transaction services applications.', italic=True, size=18, before=0, after=30),
    w_para('Upgraded with statutory raw data, ratio workings and data-linked commentary from easyJet’s FY2025 and FY2024 annual reports.', italic=True, size=18, before=0, after=110),

    w_para('1. Executive Summary', bold=True, size=24, before=140, after=60),
    w_para('- easyJet operates a cash-before-service model. At 30 September 2025, customer contract liabilities were GBP1.969bn, equal to 19.5% of annual revenue. Revenue recognition therefore remains both an income statement and a balance sheet workstream.', size=19),
    w_para('- Liquidity improved on a year-end basis, with the current ratio increasing from 1.02x to 1.11x and adjusted net cash strengthening from GBP181m to GBP602m. However, statutory current liabilities are inflated by customer prepayments, so airline liquidity is better assessed alongside the 1.88x liquidity cover excluding customer balances.', size=19),
    w_para('- The three planning areas most likely to drive manager and specialist attention are: (1) revenue and customer liabilities; (2) the leased aircraft maintenance provision of GBP939m; and (3) the carrying value and resilience of a capital-intensive asset base that includes GBP4.791bn of owned PPE, GBP1.015bn of right-of-use assets and GBP771m of goodwill and other non-current intangible assets.', size=19),

    w_para('2. Key Audit Risks', bold=True, size=24, before=140, after=60),
    w_table(
        ['Risk area', 'Why it matters', 'Assertion / focus', 'Risk outcome', 'Planning response'],
        [
            ['Revenue and customer liabilities', 'IFRS 15 is applied across multiple triggers and the balance-sheet exposure is large: unearned revenue rose from GBP1.741bn to GBP1.950bn and total customer contract liabilities from GBP1.776bn to GBP1.969bn.', 'Cut-off, completeness, accuracy, valuation; presumed fraud risk on revenue recognition.', 'Revenue could be recognised in the wrong period; cancelled flights may be misclassified between unearned revenue, refunds and vouchers; breakage could be overstated; holidays revenue could be released too early.', 'Front-load walkthroughs over booking, flown, cancellation and voucher flows. Reperform the contract-liability roll-forward. Use booking curves, flown sectors and holidays stay dates around 30 September. Inspect journals and challenge breakage and compensation estimates using post year-end evidence.'],
            ['Leased aircraft maintenance provision', 'This is a large judgemental balance and a disclosed key audit matter. The provision increased from GBP894m to GBP939m and remains 9.3% of revenue.', 'Valuation, completeness, accuracy.', 'The provision may be understated if assumptions on utilisation, restoration scope, escalation or discounting are optimistic.', 'Obtain the model early. Test hours and cycles to engineering data. Compare cost assumptions to recent maintenance events and third-party contracts. Review lease terms and challenge discounting and sensitivity analysis.'],
            ['Claims, cancellations, ETS and related accruals', 'Operational disruption affects customer compensation, refund liabilities, ETS accruals and legal exposures. Current intangible assets of GBP518m also reflect carbon allowances that settle against accruals over the annual cycle.', 'Completeness, valuation, cut-off, presentation.', 'Accruals may be incomplete, outdated or inconsistently reflected across provisions, payables and disclosures.', 'Use disruption and cancellation data to challenge completeness. Test post year-end settlements. Inspect legal and claims correspondence. Recalculate ETS-related accruals and check consistent capture across all affected balances.'],
            ['Carrying value of Airline assets', 'The airline asset base is material: owned PPE of GBP4.791bn, right-of-use assets of GBP1.015bn, goodwill of GBP387m and other intangible assets of GBP384m.', 'Valuation and disclosure.', 'Headroom may be overstated if revenue, margin or cost assumptions are too optimistic or downside cases are not severe enough.', 'Tie forecasts to Board-approved plans. Compare assumptions to external market data and internal trading. Involve valuation support on WACC. Challenge downside scenarios, not just the base case.'],
            ['Going concern and liquidity', 'Current liquidity improved and net cash strengthened, but the model remains sensitive to downside assumptions and the execution of mitigations such as delaying deliveries or reducing capex.', 'Disclosure, completeness, presentation.', 'Downside cases may be too mild, or planned mitigations may not be executable in the required timeframe.', 'Review the base case and severe downside model early. Test consistency with principal risks, fleet commitments and financing terms. Inspect facilities and challenge the timing and practicality of management actions.'],
        ],
        [1500, 2350, 1500, 2300, 3050],
        font_size=15,
    ),

    w_para('3. Analytical Review', bold=True, size=24, before=140, after=60),
    w_para('- The review below preserves the existing planning lens but now ties each point back to calculated movements from statutory line items. For easyJet, customer liabilities, liquidity and cost intensity are more informative than retail-style gross margin or inventory metrics.', size=19),
    w_table(['Metric', 'FY2024', 'FY2025', 'Movement', 'Planning view'], analytical_rows_doc, [2200, 900, 900, 950, 4400], font_size=15),

    w_para('4. Revenue Recognition Under IFRS 15', bold=True, size=24, before=140, after=60),
    w_para('- Revenue recognition remains product-specific. Passenger seats and most ancillaries are recognised when the flight takes place. Cancellation fees are recognised when processed. easyJet Plus is recognised over the membership term. Partner revenue is recognised net where easyJet acts as agent. easyJet holidays revenue is recognised over the holiday period for non-flight elements.', size=19),
    w_para('- The contract-liability bridge should be built explicitly. For FY2025, opening customer contract liabilities of GBP1.776bn plus GBP11.019bn deferred during the year, less GBP10.064bn recognised as revenue and GBP765m of APD, plus GBP3m of net transfers and FX, reconciled to closing customer contract liabilities of GBP1.969bn.', size=19),
    w_table(['Bridge item', 'FY2024', 'FY2025', 'Audit read-across'], contract_bridge_rows_doc, [2300, 850, 850, 3600], font_size=15),
    w_para('- Revenue recognised that was included in opening customer contract liabilities was GBP1.703bn in FY2025 (GBP1.678bn from unearned revenue and GBP25m from other customer liabilities). That makes the opening-liability unwind a real audit dataset, not just a disclosure point.', size=19),
    w_para('- Timing remains the primary accounting issue. The correct trigger is usually not booking date or cash receipt. For the airline business it is the flown date; for holidays it is the passage of the stay. The growth in easyJet holidays revenue from GBP1.137bn to GBP1.440bn makes this over-time recognition issue more commercially material than in prior years.', size=19),
    w_para('- Variable consideration is a live issue rather than a textbook one. Delay and cancellation compensation can reduce revenue up to the value of the fare, with any excess recognised in other costs. Vouchers, refunds and breakage judgments determine whether balances stay in liabilities or are released to revenue. The rise in customer liabilities from 19.1% to 19.5% of revenue reinforces that this judgement remains material.', size=19),
    w_para('- Cut-off risk is elevated by the size of the year-end liability bridge. A mis-dated flight status, failed interface or incorrect transfer between unearned revenue and other customer liabilities could create a material period-end misstatement. The APD bridge in the contract-liability note remains an important technical adjustment and should be built explicitly in the audit file.', size=19),

    w_para('5. Audit Approach Summary', bold=True, size=24, before=140, after=60),
    w_para('- Front-load systems and data-flow understanding. Planning quality depends on how booking, departure-control, cancellation, voucher and finance systems interact.', size=19),
    w_para('- Treat revenue as an IT-enabled workstream. The team should test the automated rules that release unearned revenue, create refund or voucher liabilities and recognise holidays revenue over time.', size=19),
    w_para('- Use commercial evidence in the challenge process. The margin and liquidity story should be checked against booking curves, flown sectors around year end, holidays stay dates, fuel cost intensity, disruption experience and year-end liability movements.', size=19),
    w_para('- Keep scepticism focused on judgemental balances and classification decisions rather than cash movements alone, particularly maintenance provisioning, customer liabilities, impairment models and downside planning assumptions.', size=19),
    w_para('- From an accounting advisory and transaction services perspective, the same workstream matters because customer liabilities, vouchers/refunds and maintenance provisions can affect revenue quality, net working capital and debt-like item analysis in diligence.', size=19),

    w_para('6. Conclusion', bold=True, size=24, before=140, after=60),
    w_para('- Revenue and customer liabilities remain the highest-risk area. The combination of multiple IFRS 15 triggers, a customer liability balance equal to 19.5% of revenue, and dependence on complex systems makes this the most pervasive planning risk on the file.', size=19),
    w_para('- The maintenance provision remains the most judgemental single estimate. However, the data-backed analytical review indicates that the wider audit story is still driven by the prepayment revenue model, cash conversion profile and the balance-sheet consequences of recognising revenue at the correct point in time.', size=19),

    w_para('7. Raw Financial Data', bold=True, size=24, before=140, after=60),
    w_table(['Line item', 'FY2025', 'FY2024', 'FY2023', 'Source', 'Comment'], raw_rows_doc, [1550, 700, 700, 700, 2000, 1850], font_size=15),

    w_para('8. Ratio Workings', bold=True, size=24, before=140, after=60),
    w_table(['Metric', 'Formula', 'Line items used', 'Why correct', 'FY2024', 'FY2025', 'Movement', 'Audit interpretation'], ratio_rows_doc, [1850, 1500, 1650, 1700, 700, 700, 800, 1600], font_size=14),

    w_para('9. Revised Analytical Review - Data-Integrated Commentary', bold=True, size=24, before=140, after=60),
    w_para('- Growth and mix: group revenue grew 8.6% in FY2025 after 13.9% in FY2024, while easyJet holidays revenue still grew 26.7% from GBP1.137bn to GBP1.440bn. The mix shift matters because the holidays revenue pattern is different from the airline flown-date trigger and should therefore receive disproportionate planning attention.', size=19),
    w_para('- Margin profile: operating margin improved from 6.3% to 6.9%, but EBITDA margin eased from 14.6% to 14.2% and depreciation and amortisation fell from 8.3% to 7.4% of revenue. That is a stronger audit signal than the headline margin alone: part of the statutory EBIT improvement sits in below-EBITDA movements and should be bridged to depreciation, amortisation, fleet mix and estimate movements.', size=19),
    w_para('- Working capital and cash: receivables days were effectively flat at 4.8 days, while payables days shortened from 17.4 to 15.2 days. This may indicate less supplier financing support or different settlement timing, but either way easyJet’s cash profile is still driven more by customer prepayments, which is consistent with operating cash conversion improving from 2.52x to 2.69x.', size=19),
    w_para('- Customer liabilities and cut-off: unearned revenue increased from 18.7% to 19.3% of revenue and total customer contract liabilities increased from 19.1% to 19.5%. That is the most persuasive data point underpinning the existing revenue recognition commentary: a large proportion of the revenue cycle still sits on the balance sheet at year end.', size=19),
    w_para('- Liquidity and capital structure: the current ratio improved from 1.02x to 1.11x and the adjusted quick ratio from 0.89x to 0.99x, while liquidity cover excluding customer balances improved from 1.47x to 1.88x and adjusted net cash improved from GBP181m to GBP602m. This is positive, but funding obligations including leases still totalled GBP2.926bn and lease liabilities still represented 16.3% of total capital, so liquidity planning and covenant-style downside thinking remain relevant.', size=19),
    w_para('- Asset intensity and estimates: owned PPE rose to GBP4.791bn, right-of-use assets remained over GBP1.0bn, and the maintenance provision increased to GBP939m. Even though maintenance provision as a percentage of revenue eased slightly from 9.6% to 9.3%, the estimate remains large enough to justify continued specialist challenge and early manager focus.', size=19),

    w_para('10. Audit Flags / Key Anomalies', bold=True, size=24, before=140, after=60),
    w_para('- Customer contract liabilities increased by GBP193m year on year and remained close to one-fifth of annual revenue. That is the clearest balance-sheet flag supporting revenue cut-off and completeness work.', size=19),
    w_para('- easyJet holidays revenue grew materially faster than the wider group. A newer and faster-growing revenue stream with over-time recognition should receive more attention than its absolute size alone might suggest.', size=19),
    w_para('- Operating margin improved while EBITDA margin softened and depreciation and amortisation fell as a share of revenue. That does not indicate an error, but it does mean the margin story should be reconciled carefully rather than accepted as a simple trading uplift.', size=19),
    w_para('- Receivables days were stable, but payables days shortened. The working capital model is therefore still more dependent on customer prepayments than on trade creditor support, with some scope for timing effects in creditor settlement.', size=19),
    w_para('- Lease liabilities reduced, but the balance remains economically material. Going concern and impairment work should continue to treat leases as part of the funding structure rather than as a technical afterthought.', size=19),
    w_para('- The same data points would matter in advisory and diligence work: customer liabilities drive revenue-quality and net working capital analysis, while maintenance provisions and leases can become debt-like or normalisation discussion points.', size=19),

    w_para('11. Assumptions and Data Limitations', bold=True, size=24, before=140, after=60),
    w_para('- Gross profit is not presented and is not especially meaningful for easyJet’s nature-based airline cost base. The analysis therefore uses EBITDA, operating profit and total operating costs before depreciation and amortisation.', size=19),
    w_para('- Inventories are not separately disclosed on the face of the balance sheet. Aircraft spares are held within property, plant and equipment, so inventory days have not been forced into the analysis.', size=19),
    w_para('- FY2024 receivables and payables days use FY2023 closing balances from the FY2024 annual report to support average-balance workings. That is more robust than using closing balances alone.', size=19),
    w_para('- Current intangible assets are excluded from the adjusted quick ratio because they primarily comprise carbon allowances and related items that do not provide the same short-term liquidity profile as cash, investments or receivables.', size=19),
    w_para('- easyJet is seasonal. All year-end ratios are therefore planning indicators rather than complete substitutes for monthly trading, cash and booking data.', size=19),

    w_para('12. Suggested Excel Workbook Tab Structure', bold=True, size=24, before=140, after=60),
    w_para('- Raw_Financials: statutory line items, note references and any derived subtotals used in the review.', size=19),
    w_para('- Ratio_Calculations: formula-driven workings linked back to the raw financials sheet.', size=19),
    w_para('- Analytical_Review: short data-led commentary by theme.', size=19),
    w_para('- Audit_Flags: concise summary of the movements most likely to change planning response.', size=19),
    w_para('- Revenue_Workstream: IFRS 15 triggers, contract-liability roll-forward and cut-off considerations.', size=19),
]

doc_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + f'<w:document xmlns:w="{W_NS}"><w:body>{"".join(body_parts)}{sectpr_xml}</w:body></w:document>'

with ZipFile(DOCX, 'w', compression=ZIP_DEFLATED) as z:
    for name, data in doc_items.items():
        z.writestr(name, data)
    z.writestr('word/document.xml', doc_xml)

# -----------------------------
# Build workbook from scratch
# -----------------------------
MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

sheet_names = ['Overview', 'Raw_Financials', 'Ratio_Calculations', 'Analytical_Review', 'Revenue_Workstream', 'Risk_Register', 'Audit_Flags', 'Source_Log']

styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><name val="Aptos"/><family val="2"/></font>
    <font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Aptos"/><family val="2"/></font>
    <font><b/><sz val="11"/><name val="Aptos"/><family val="2"/></font>
    <font><i/><sz val="11"/><name val="Aptos"/><family val="2"/></font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0B3D5C"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDDE8EF"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF4F7F9"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="2" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def col_letter(idx: int) -> str:
    out = ''
    while idx:
        idx, rem = divmod(idx - 1, 26)
        out = chr(65 + rem) + out
    return out


def make_cell(ref, value='', style=0, ctype='str', formula=None):
    if ctype == 'str':
        return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'
    if ctype == 'num':
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    if ctype == 'formula':
        return f'<c r="{ref}" s="{style}"><f>{escape(formula)}</f><v>{value}</v></c>'
    raise ValueError(ctype)


def make_sheet(rows, widths):
    max_col = max(len(r) for r in rows)
    xml_rows = []
    for r_idx, row in enumerate(rows, start=1):
        cells = []
        for c_idx, item in enumerate(row, start=1):
            if item is None:
                continue
            ref = f'{col_letter(c_idx)}{r_idx}'
            cells.append(make_cell(ref, item.get('value', ''), item.get('style', 0), item.get('type', 'str'), item.get('formula')))
        xml_rows.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
    cols = '<cols>' + ''.join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths, start=1)) + '</cols>'
    dim = f'A1:{col_letter(max_col)}{len(rows)}'
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + f'<worksheet xmlns="{MAIN_NS}"><dimension ref="{dim}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18"/>{cols}<sheetData>{"".join(xml_rows)}</sheetData></worksheet>'

# Overview
rows_overview = [
    [ {'value': 'Audit Risk Review - easyJet plc FY2025', 'style': 1}, None, None ],
    [ {'value': 'Workbook purpose', 'style': 2}, {'value': 'Add a defendable statutory data layer and formula-driven calculations underneath the existing audit planning memo.', 'style': 0} ],
    [ {'value': 'What changed', 'style': 2}, {'value': 'Raw financial line items, average-balance workings, airline-adapted ratios, data-linked commentary and a separate audit flags view.', 'style': 0} ],
    [ {'value': 'How to use', 'style': 2}, {'value': 'Start with Raw_Financials, review Ratio_Calculations, then use Analytical_Review and Audit_Flags to tighten planning commentary and interview talking points.', 'style': 0} ],
    [ {'value': 'Important note', 'style': 4}, {'value': 'Gross profit and inventory days are not forced into the model because they are not especially meaningful for easyJet’s statutory presentation. The workbook uses operating-cost and customer-liability metrics instead.', 'style': 4} ],
]

# Raw_Financials
rows_raw = [
    [ {'value': 'Raw Financials', 'style': 1}, None, None, None, None, None ],
    [ {'value': 'Line item', 'style': 2}, {'value': 'FY2025', 'style': 2}, {'value': 'FY2024', 'style': 2}, {'value': 'FY2023', 'style': 2}, {'value': 'Source', 'style': 2}, {'value': 'Comment', 'style': 2} ],
]
for name, fy25, fy24, fy23, source, comment in raw_rows:
    row = [ {'value': name, 'style': 0} ]
    for v in [fy25, fy24, fy23]:
        if isinstance(v, str):
            row.append({'value': v, 'style': 0})
        else:
            row.append({'value': v, 'type': 'num', 'style': 0})
    row.extend([{'value': source, 'style': 0}, {'value': comment, 'style': 0}])
    rows_raw.append(row)

# Convert selected derived rows to formulas in raw sheet
# row numbers on sheet:
raw_sheet_row = {name: i+3 for i,(name,*_) in enumerate(raw_rows)}
for r in rows_raw:
    pass
# overwrite formula-based cells manually by row index in rows_raw list (offset +2 header rows)
def set_formula_row(name, col_letter_target, formula, value):
    row_num = raw_sheet_row[name]
    row = rows_raw[row_num-1]  # because rows_raw starts at row1
    col_idx = {'B':1,'C':2,'D':3}[col_letter_target]
    row[col_idx] = {'value': value, 'type': 'formula', 'formula': formula, 'style': 0}

for col, vals in [('B', ('FY2025','Raw_Financials!B')), ('C', ('FY2024','Raw_Financials!C')), ('D', ('FY2023','Raw_Financials!D'))]:
    pass
set_formula_row('Airline revenue (passenger + ancillary)', 'B', f"={cell_ref('Raw_Financials','B',raw_sheet_row['Passenger revenue'])}+{cell_ref('Raw_Financials','B',raw_sheet_row['Ancillary revenue'])}", 8666)
set_formula_row('Airline revenue (passenger + ancillary)', 'C', f"={cell_ref('Raw_Financials','C',raw_sheet_row['Passenger revenue'])}+{cell_ref('Raw_Financials','C',raw_sheet_row['Ancillary revenue'])}", 8172)
set_formula_row('Airline revenue (passenger + ancillary)', 'D', f"={cell_ref('Raw_Financials','D',raw_sheet_row['Passenger revenue'])}+{cell_ref('Raw_Financials','D',raw_sheet_row['Ancillary revenue'])}", 7395)
set_formula_row('Statutory EBITDA (derived)', 'B', f"={cell_ref('Raw_Financials','B',raw_sheet_row['Operating profit'])}+{cell_ref('Raw_Financials','B',raw_sheet_row['Depreciation and amortisation'])}", 1439)
set_formula_row('Statutory EBITDA (derived)', 'C', f"={cell_ref('Raw_Financials','C',raw_sheet_row['Operating profit'])}+{cell_ref('Raw_Financials','C',raw_sheet_row['Depreciation and amortisation'])}", 1359)
set_formula_row('Statutory EBITDA (derived)', 'D', f"={cell_ref('Raw_Financials','D',raw_sheet_row['Operating profit'])}+{cell_ref('Raw_Financials','D',raw_sheet_row['Depreciation and amortisation'])}", 1126)
set_formula_row('Total operating costs before depreciation and amortisation', 'B', f"={cell_ref('Raw_Financials','B',raw_sheet_row['Revenue'])}-{cell_ref('Raw_Financials','B',raw_sheet_row['Statutory EBITDA (derived)'])}", 8667)
set_formula_row('Total operating costs before depreciation and amortisation', 'C', f"={cell_ref('Raw_Financials','C',raw_sheet_row['Revenue'])}-{cell_ref('Raw_Financials','C',raw_sheet_row['Statutory EBITDA (derived)'])}", 7950)
set_formula_row('Total operating costs before depreciation and amortisation', 'D', f"={cell_ref('Raw_Financials','D',raw_sheet_row['Revenue'])}-{cell_ref('Raw_Financials','D',raw_sheet_row['Statutory EBITDA (derived)'])}", 7045)
set_formula_row('Funding obligations (borrowings + leases)', 'B', f"={cell_ref('Raw_Financials','B',raw_sheet_row['Borrowings'])}+{cell_ref('Raw_Financials','B',raw_sheet_row['Lease liabilities'])}", 2926)
set_formula_row('Funding obligations (borrowings + leases)', 'C', f"={cell_ref('Raw_Financials','C',raw_sheet_row['Borrowings'])}+{cell_ref('Raw_Financials','C',raw_sheet_row['Lease liabilities'])}", 3280)
set_formula_row('Funding obligations (borrowings + leases)', 'D', f"={cell_ref('Raw_Financials','D',raw_sheet_row['Borrowings'])}+{cell_ref('Raw_Financials','D',raw_sheet_row['Lease liabilities'])}", 2884)
set_formula_row('Adjusted net cash/(debt) incl. leases', 'B', f"={cell_ref('Raw_Financials','B',raw_sheet_row['Cash and other investments'])}-{cell_ref('Raw_Financials','B',raw_sheet_row['Borrowings'])}-{cell_ref('Raw_Financials','B',raw_sheet_row['Lease liabilities'])}", 602)
set_formula_row('Adjusted net cash/(debt) incl. leases', 'C', f"={cell_ref('Raw_Financials','C',raw_sheet_row['Cash and other investments'])}-{cell_ref('Raw_Financials','C',raw_sheet_row['Borrowings'])}-{cell_ref('Raw_Financials','C',raw_sheet_row['Lease liabilities'])}", 181)
set_formula_row('Adjusted net cash/(debt) incl. leases', 'D', f"={cell_ref('Raw_Financials','D',raw_sheet_row['Cash and other investments'])}-{cell_ref('Raw_Financials','D',raw_sheet_row['Borrowings'])}-{cell_ref('Raw_Financials','D',raw_sheet_row['Lease liabilities'])}", 41)
set_formula_row('Total customer contract liabilities', 'B', f"={cell_ref('Raw_Financials','B',raw_sheet_row['Unearned revenue'])}+{cell_ref('Raw_Financials','B',raw_sheet_row['Other customer contract liabilities'])}", 1969)
set_formula_row('Total customer contract liabilities', 'C', f"={cell_ref('Raw_Financials','C',raw_sheet_row['Unearned revenue'])}+{cell_ref('Raw_Financials','C',raw_sheet_row['Other customer contract liabilities'])}", 1776)
set_formula_row('Total customer contract liabilities', 'D', f"={cell_ref('Raw_Financials','D',raw_sheet_row['Unearned revenue'])}+{cell_ref('Raw_Financials','D',raw_sheet_row['Other customer contract liabilities'])}", 1580)

# Ratio calculations
rows_ratio = [
    [ {'value': 'Ratio Calculations', 'style': 1}, None, None, None, None, None, None, None ],
    [ {'value': 'Metric', 'style': 2}, {'value': 'Formula', 'style': 2}, {'value': 'Line items used', 'style': 2}, {'value': 'Why this formula is correct', 'style': 2}, {'value': 'FY2024', 'style': 2}, {'value': 'FY2025', 'style': 2}, {'value': 'Movement', 'style': 2}, {'value': 'Audit interpretation', 'style': 2} ],
]
for idx, row in enumerate(ratio_rows, start=3):
    style_map = 3 if row['style'] == 'pct' else 5
    rows_ratio.append([
        {'value': row['metric'], 'style': 0},
        {'value': row['formula_text'], 'style': 0},
        {'value': row['line_items'], 'style': 0},
        {'value': row['why'], 'style': 0},
        {'value': row['fy24_value'], 'type': 'formula', 'formula': row['fy24_formula'][1:] if row['fy24_formula'].startswith('=') else row['fy24_formula'], 'style': style_map},
        {'value': row['fy25_value'], 'type': 'formula', 'formula': row['fy25_formula'][1:] if row['fy25_formula'].startswith('=') else row['fy25_formula'], 'style': style_map},
        {'value': row['fy25_value'] - row['fy24_value'], 'type': 'formula', 'formula': f'F{idx}-E{idx}', 'style': style_map if row['style'] == 'pct' else 5},
        {'value': row['audit'], 'style': 0},
    ])

# Analytical review
rows_analysis = [
    [ {'value': 'Analytical Review', 'style': 1}, None, None, None ],
    [ {'value': 'Theme', 'style': 2}, {'value': 'Data signal', 'style': 2}, {'value': 'Why it matters', 'style': 2}, {'value': 'Audit implication', 'style': 2} ],
    [ {'value': 'Growth and mix', 'style': 0}, {'value': 'Revenue growth slowed from 13.9% to 8.6%, while easyJet holidays revenue still grew 26.7%.', 'style': 0}, {'value': 'The mix shift makes the over-time recognition pattern in Holidays more important than in prior years.', 'style': 0}, {'value': 'Retain revenue as a significant planning risk and give Holidays specific cut-off testing rather than relying solely on airline procedures.', 'style': 0} ],
    [ {'value': 'Margin quality', 'style': 0}, {'value': 'Operating margin improved from 6.3% to 6.9%, EBITDA margin eased from 14.6% to 14.2%, and D&A fell from 8.3% to 7.4% of revenue.', 'style': 0}, {'value': 'The statutory profit improvement does not look like a clean trading story on every measure.', 'style': 0}, {'value': 'Bridge profit movement to D&A, asset mix, useful lives and capitalisation rather than accepting the EBIT uplift at face value.', 'style': 0} ],
    [ {'value': 'Working capital', 'style': 0}, {'value': 'Receivables days were flat at c.4.8 days; payables days shortened from 17.4 to 15.2 days.', 'style': 0}, {'value': 'The working capital model is still driven more by customer prepayments than by collections or supplier stretch, with some scope for settlement timing effects.', 'style': 0}, {'value': 'Audit focus should stay on customer liabilities and period-end revenue release rather than on debtor recoverability.', 'style': 0} ],
    [ {'value': 'Customer liabilities', 'style': 0}, {'value': 'Customer contract liabilities increased from 19.1% to 19.5% of revenue.', 'style': 0}, {'value': 'A large proportion of the revenue cycle still sits on the balance sheet at year end.', 'style': 0}, {'value': 'This is the strongest data point supporting cut-off, completeness and breakage challenge in the IFRS 15 workstream.', 'style': 0} ],
    [ {'value': 'Liquidity and leverage', 'style': 0}, {'value': 'Current ratio improved from 1.02x to 1.11x, liquidity cover excluding customer balances from 1.47x to 1.88x, and adjusted net cash from GBP181m to GBP602m.', 'style': 0}, {'value': 'Year-end liquidity improved materially; customer prepayments inflate statutory current liabilities, so the adjusted cover metric is more decision-useful for an airline.', 'style': 0}, {'value': 'Going concern work should remain robust, with clear challenge over downside assumptions and management actions.', 'style': 0} ],
    [ {'value': 'Capital intensity', 'style': 0}, {'value': 'Owned PPE rose to GBP4.791bn, ROU assets remained at GBP1.015bn and the maintenance provision increased to GBP939m.', 'style': 0}, {'value': 'The asset base and maintenance obligations remain large relative to the trading base.', 'style': 0}, {'value': 'Impairment and maintenance provision work should remain early-planned and specialist-supported.', 'style': 0} ],
]

# Revenue workstream
rows_revenue = [
    [ {'value': 'Revenue Workstream - IFRS 15', 'style': 1}, None, None, None, None, None ],
    [ {'value': 'Technical bridge', 'style': 4}, {'value': 'The contract-liability roll-forward is presented inclusive of APD and other charges, while passenger revenue is recognised excluding APD. The audit file should bridge this explicitly.', 'style': 4}, None, None, None, None ],
    [ {'value': 'Revenue stream', 'style': 2}, {'value': 'FY2025 amount (GBPm)', 'style': 2}, {'value': 'Recognition trigger', 'style': 2}, {'value': 'Judgement / variable consideration', 'style': 2}, {'value': 'Cut-off exposure', 'style': 2}, {'value': 'Audit response', 'style': 2} ],
    [ {'value': 'Passenger revenue', 'style': 0}, {'value': 6072, 'type': 'num', 'style': 0}, {'value': 'Flight takes place', 'style': 0}, {'value': 'Compensation offsets fare revenue; no-show logic and claims estimates matter.', 'style': 0}, {'value': 'Flights either side of 30 September; cancelled flights moved to refund or voucher liabilities.', 'style': 0}, {'value': 'Test flown-date logic, reperform liability unwind, inspect journals and challenge compensation assumptions with post year-end evidence.', 'style': 0} ],
    [ {'value': 'Airline ancillary revenue', 'style': 0}, {'value': 2594, 'type': 'num', 'style': 0}, {'value': 'Usually when the flight takes place; cancellation fees when processed', 'style': 0}, {'value': 'Product-specific triggers may differ across ancillary types.', 'style': 0}, {'value': 'Incorrect event date or product mapping could shift revenue across periods.', 'style': 0}, {'value': 'Walk through ancillary mapping and test event-date logic by product.', 'style': 0} ],
    [ {'value': 'Partner revenue and insurance', 'style': 0}, {'value': 'Included within airline ancillary revenue', 'style': 0}, {'value': 'Net commission when the underlying service occurs; insurance commission at booking', 'style': 0}, {'value': 'Principal-versus-agent judgement and correct trigger for insurance commission.', 'style': 0}, {'value': 'Gross presentation or wrong trigger could overstate revenue.', 'style': 0}, {'value': 'Inspect contracts, confirm control and inventory risk, and compare gross cash with net recognition.', 'style': 0} ],
    [ {'value': 'easyJet Plus', 'style': 0}, {'value': 'Included within airline ancillary revenue', 'style': 0}, {'value': 'Over membership term', 'style': 0}, {'value': 'Deferred income period and renewal logic.', 'style': 0}, {'value': 'Release too early around year end.', 'style': 0}, {'value': 'Test membership dates and reperform release profile.', 'style': 0} ],
    [ {'value': 'easyJet holidays revenue', 'style': 0}, {'value': 1440, 'type': 'num', 'style': 0}, {'value': 'Over holiday period for non-flight elements', 'style': 0}, {'value': 'Refunds, vouchers and holidays spanning year end require earned-to-date judgement.', 'style': 0}, {'value': 'Recognition at departure rather than over the stay.', 'style': 0}, {'value': 'Select bookings spanning 30 September and reperform recognition over the stay.', 'style': 0} ],
    [ {'value': 'Contract liability roll-forward', 'style': 2}, None, None, None, None, None ],
    [ {'value': 'Item', 'style': 2}, {'value': 'FY2024', 'style': 2}, {'value': 'FY2025', 'style': 2}, {'value': 'Audit point', 'style': 2}, None, None ],
    [ {'value': 'Opening unearned revenue', 'style': 0}, {'value': 1501, 'type': 'num', 'style': 0}, {'value': 1741, 'type': 'num', 'style': 0}, {'value': 'Opening balance released as travel occurs.', 'style': 0} ],
    [ {'value': 'Opening other customer liabilities', 'style': 0}, {'value': 79, 'type': 'num', 'style': 0}, {'value': 35, 'type': 'num', 'style': 0}, {'value': 'Primarily vouchers and unresolved refund options.', 'style': 0} ],
    [ {'value': 'Revenue deferred during the year', 'style': 0}, {'value': 10170, 'type': 'num', 'style': 0}, {'value': 11019, 'type': 'num', 'style': 0}, {'value': 'Inclusive of APD and other charges.', 'style': 0} ],
    [ {'value': 'Revenue recognised from opening unearned revenue', 'style': 0}, {'value': 1399, 'type': 'num', 'style': 0}, {'value': 1678, 'type': 'num', 'style': 0}, {'value': 'Useful bridge from opening liability to current-year revenue.', 'style': 0} ],
    [ {'value': 'Revenue recognised from opening other customer liabilities', 'style': 0}, {'value': 47, 'type': 'num', 'style': 0}, {'value': 25, 'type': 'num', 'style': 0}, {'value': 'Shows unwind of voucher / refund liabilities into revenue.', 'style': 0} ],
    [ {'value': 'Closing unearned revenue', 'style': 0}, {'value': 1741, 'type': 'num', 'style': 0}, {'value': 1950, 'type': 'num', 'style': 0}, {'value': 'Key year-end cut-off balance.', 'style': 0} ],
    [ {'value': 'Closing other customer liabilities', 'style': 0}, {'value': 35, 'type': 'num', 'style': 0}, {'value': 19, 'type': 'num', 'style': 0}, {'value': 'Check classification between unearned revenue and other payables after cancellations.', 'style': 0} ],
    [ {'value': 'Explicit total customer liability bridge', 'style': 2}, None, None, None, None, None ],
    [ {'value': 'Bridge item', 'style': 2}, {'value': 'FY2024', 'style': 2}, {'value': 'FY2025', 'style': 2}, {'value': 'Audit point', 'style': 2}, None, None ],
    [ {'value': 'Opening total customer contract liabilities', 'style': 0}, {'value': 1580, 'type': 'formula', 'formula': f"{cell_ref('Raw_Financials','D',raw_sheet_row['Total customer contract liabilities'])}", 'style': 0}, {'value': 1776, 'type': 'formula', 'formula': f"{cell_ref('Raw_Financials','C',raw_sheet_row['Total customer contract liabilities'])}", 'style': 0}, {'value': 'Opening deferred revenue base beneath current-year revenue.', 'style': 0} ],
    [ {'value': 'Revenue deferred during the year', 'style': 0}, {'value': 10170, 'type': 'formula', 'formula': f"{cell_ref('Raw_Financials','C',raw_sheet_row['Revenue deferred during the year'])}", 'style': 0}, {'value': 11019, 'type': 'formula', 'formula': f"{cell_ref('Raw_Financials','B',raw_sheet_row['Revenue deferred during the year'])}", 'style': 0}, {'value': 'Advance cash receipts before performance obligations are satisfied.', 'style': 0} ],
    [ {'value': 'Revenue recognised during the year', 'style': 0}, {'value': 9266, 'type': 'formula', 'formula': f"{cell_ref('Raw_Financials','D',raw_sheet_row['Total customer contract liabilities'])}+{cell_ref('Raw_Financials','C',raw_sheet_row['Revenue deferred during the year'])}-{cell_ref('Raw_Financials','C',raw_sheet_row['APD on revenue recognised'])}+3-{cell_ref('Raw_Financials','C',raw_sheet_row['Total customer contract liabilities'])}", 'style': 0}, {'value': 10064, 'type': 'formula', 'formula': f"{cell_ref('Raw_Financials','C',raw_sheet_row['Total customer contract liabilities'])}+{cell_ref('Raw_Financials','B',raw_sheet_row['Revenue deferred during the year'])}-{cell_ref('Raw_Financials','B',raw_sheet_row['APD on revenue recognised'])}+3-{cell_ref('Raw_Financials','B',raw_sheet_row['Total customer contract liabilities'])}", 'style': 0}, {'value': 'Includes both unearned revenue and other customer liabilities released to revenue.', 'style': 0} ],
    [ {'value': 'APD on revenue recognised', 'style': 0}, {'value': 711, 'type': 'num', 'style': 0}, {'value': 765, 'type': 'num', 'style': 0}, {'value': 'Technical bridge item because the note is presented inclusive of APD.', 'style': 0} ],
    [ {'value': 'Net transfer / FX movement in other customer liabilities', 'style': 0}, {'value': 3, 'type': 'num', 'style': 0}, {'value': 3, 'type': 'num', 'style': 0}, {'value': 'Additional less reduced contract liabilities plus FX in the other-customer-liability bucket.', 'style': 0} ],
    [ {'value': 'Closing total customer contract liabilities', 'style': 0}, {'value': 1776, 'type': 'formula', 'formula': f"{cell_ref('Raw_Financials','C',raw_sheet_row['Total customer contract liabilities'])}", 'style': 0}, {'value': 1969, 'type': 'formula', 'formula': f"{cell_ref('Raw_Financials','B',raw_sheet_row['Total customer contract liabilities'])}", 'style': 0}, {'value': 'Reconciles to the year-end balance sheet exposure.', 'style': 0} ],
]

# Risk register
rows_risk = [
    [ {'value': 'Risk Register', 'style': 1}, None, None, None ],
    [ {'value': 'Risk area', 'style': 2}, {'value': 'Key data signal', 'style': 2}, {'value': 'Why it matters', 'style': 2}, {'value': 'Planning response', 'style': 2} ],
    [ {'value': 'Revenue and customer liabilities', 'style': 0}, {'value': 'Customer contract liabilities increased from GBP1.776bn to GBP1.969bn; 19.5% of revenue at year end.', 'style': 0}, {'value': 'Large balance-sheet exposure beneath reported revenue makes cut-off and completeness central.', 'style': 0}, {'value': 'System walkthroughs, roll-forward reperformance, population analytics and post year-end testing.', 'style': 0} ],
    [ {'value': 'Maintenance provision', 'style': 0}, {'value': 'Provision increased from GBP894m to GBP939m; still 9.3% of revenue.', 'style': 0}, {'value': 'Largest judgemental provision and a disclosed key audit matter.', 'style': 0}, {'value': 'Early model challenge, engineering data testing and specialist review.', 'style': 0} ],
    [ {'value': 'Liquidity and going concern', 'style': 0}, {'value': 'Adjusted net cash improved from GBP181m to GBP602m and liquidity cover excluding customer balances improved from 1.47x to 1.88x.', 'style': 0}, {'value': 'Year-end liquidity improved, but reliance on customer cash timing and operational resilience remains high and the sector remains volatile.', 'style': 0}, {'value': 'Challenge downside cases, mitigations, financing terms and delivery flexibility.', 'style': 0} ],
    [ {'value': 'Impairment and asset carrying value', 'style': 0}, {'value': 'Owned PPE GBP4.791bn, ROU assets GBP1.015bn, goodwill GBP387m, other intangibles GBP384m.', 'style': 0}, {'value': 'Large forecast-dependent asset base in a commercially sensitive sector.', 'style': 0}, {'value': 'Tie VIU assumptions to Board plans and challenge downside headroom.', 'style': 0} ],
    [ {'value': 'Advisory / TS crossover', 'style': 0}, {'value': 'Customer liabilities GBP1.969bn; maintenance provision GBP939m; leases GBP1.045bn.', 'style': 0}, {'value': 'These balances would also matter in diligence through revenue quality, net working capital and debt-like item discussions.', 'style': 0}, {'value': 'Keep the bridge and classification analysis clear enough to support audit, advisory and deal-oriented discussion.', 'style': 0} ],
]

# Audit flags
rows_flags = [
    [ {'value': 'Audit Flags', 'style': 1}, None, None ],
    [ {'value': 'Flag', 'style': 2}, {'value': 'Evidence', 'style': 2}, {'value': 'Why the team should care', 'style': 2} ],
    [ {'value': 'Customer liability build-up', 'style': 0}, {'value': 'Total customer contract liabilities rose by GBP193m and ended at 19.5% of revenue.', 'style': 0}, {'value': 'This is the strongest evidence-led support for the revenue cut-off and completeness risk.', 'style': 0} ],
    [ {'value': 'Mix shift into Holidays', 'style': 0}, {'value': 'easyJet holidays revenue rose 26.7%, well ahead of group revenue growth.', 'style': 0}, {'value': 'Different IFRS 15 timing and a newer process set mean the workstream should not be audited only by reference to the airline ledger.', 'style': 0} ],
    [ {'value': 'Margin quality needs bridging', 'style': 0}, {'value': 'Operating margin improved, but EBITDA margin declined slightly and D&A fell from 8.3% to 7.4% of revenue.', 'style': 0}, {'value': 'Part of the statutory margin improvement may sit in D&A or asset mix rather than trading alone.', 'style': 0} ],
    [ {'value': 'Working capital supported by customers, not suppliers', 'style': 0}, {'value': 'Receivables days were flat; payables days shortened; operating cash conversion improved.', 'style': 0}, {'value': 'The prepayment model remains the core cash driver, with some possibility that creditor timing also affected the year-end position.', 'style': 0} ],
    [ {'value': 'Lease funding still material', 'style': 0}, {'value': 'Lease liabilities remained GBP1.045bn and 16.3% of total capital.', 'style': 0}, {'value': 'Leases remain material to going concern, impairment and maintenance provisioning despite the improvement in leverage.', 'style': 0} ],
    [ {'value': 'Advisory / TS relevance', 'style': 0}, {'value': 'Customer liabilities, vouchers/refunds and maintenance provisions remain material within the audit story.', 'style': 0}, {'value': 'The same balances would affect revenue-quality, debt-like item and net working capital analysis in diligence.', 'style': 0} ],
]

# Source log
rows_source = [
    [ {'value': 'Source Log', 'style': 1}, None, None ],
    [ {'value': 'Source', 'style': 2}, {'value': 'Use', 'style': 2}, {'value': 'Reference', 'style': 2} ],
    [ {'value': 'easyJet Annual Report and Accounts 2025', 'style': 0}, {'value': 'Primary source for FY2025 statements, contract liabilities, receivables, payables, leases, provisions and impairment commentary.', 'style': 0}, {'value': 'p.144 income statement; p.146 statement of financial position; p.148 cash flow statement; Note 8 pp.163-164; Note 10 p.165; Note 11 pp.166-167; Note 13 p.167; Note 14 p.167; Note 15 p.167; Note 16 pp.167-168; Note 17 p.168; Note 18 pp.168-169; Note 19 pp.169-170; Note 23 p.174.', 'style': 0} ],
    [ {'value': 'easyJet Annual Report and Accounts 2024', 'style': 0}, {'value': 'Source of FY2023 opening balances needed for average-balance working capital ratios and comparative raw data.', 'style': 0}, {'value': 'p.154 income statement; p.156 statement of financial position; p.158 cash flow statement; corresponding notes for FY2023 comparatives; Note 24 p.187 for operating cash reconciliation.', 'style': 0} ],
    [ {'value': 'easyJet FY2025 investor annual report page', 'style': 0}, {'value': 'Official filing page confirming the FY2025 report.', 'style': 0}, {'value': 'Investor reports and presentations page for the 2025 annual report.', 'style': 0} ],
]

sheet_xml_map = {
    'xl/worksheets/sheet1.xml': make_sheet(rows_overview, [32, 108, 8]),
    'xl/worksheets/sheet2.xml': make_sheet(rows_raw, [36, 14, 14, 14, 42, 46]),
    'xl/worksheets/sheet3.xml': make_sheet(rows_ratio, [26, 30, 30, 34, 12, 12, 12, 54]),
    'xl/worksheets/sheet4.xml': make_sheet(rows_analysis, [22, 34, 34, 42]),
    'xl/worksheets/sheet5.xml': make_sheet(rows_revenue, [28, 18, 18, 34, 26, 42]),
    'xl/worksheets/sheet6.xml': make_sheet(rows_risk, [24, 34, 34, 44]),
    'xl/worksheets/sheet7.xml': make_sheet(rows_flags, [24, 34, 42]),
    'xl/worksheets/sheet8.xml': make_sheet(rows_source, [32, 36, 72]),
}

workbook_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + \
    '<sheets>' + ''.join(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>' for i, name in enumerate(sheet_names, start=1)) + '</sheets>' + \
    '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'

workbook_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + \
    ''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheet_names)+1)) + \
    f'<Relationship Id="rId{len(sheet_names)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' + \
    '</Relationships>'

content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + \
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + \
    '<Default Extension="xml" ContentType="application/xml"/>' + \
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' + \
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' + \
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + \
    ''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, len(sheet_names)+1)) + \
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + \
    '</Types>'

root_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + \
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' + \
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' + \
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' + \
    '</Relationships>'

now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
core_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Audit Risk Review - easyJet plc FY2025</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''

app_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' + \
    '<Application>Microsoft Excel</Application><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>8</vt:i4></vt:variant></vt:vector></HeadingPairs>' + \
    '<TitlesOfParts><vt:vector size="8" baseType="lpstr">' + ''.join(f'<vt:lpstr>{escape(n)}</vt:lpstr>' for n in sheet_names) + '</vt:vector></TitlesOfParts></Properties>'

with ZipFile(XLSX, 'w', compression=ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', root_rels)
    z.writestr('docProps/core.xml', core_xml)
    z.writestr('docProps/app.xml', app_xml)
    z.writestr('xl/workbook.xml', workbook_xml)
    z.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
    z.writestr('xl/styles.xml', styles_xml)
    for name, xml in sheet_xml_map.items():
        z.writestr(name, xml)

print('Upgraded Word report and Excel workbook with raw-data and ratio layers.')
