"""Offline, explainable B2B deal desk analysis. All amounts are annual, USD."""
from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

MONEY = Decimal('0.01')
COLUMNS = ('deal_id', 'account', 'list_acv', 'discount_pct', 'delivery_cost',
           'implementation_cost', 'payment_days', 'probability_pct')


def _number(value: str, label: str) -> Decimal:
    try:
        x = Decimal(str(value))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f'{label}: invalid decimal') from exc
    if not x.is_finite():
        raise ValueError(f'{label}: must be finite')
    return x


def _usd(value: Decimal) -> Decimal:
    return value.quantize(MONEY, rounding=ROUND_HALF_UP)


@dataclass(frozen=True)
class Deal:
    deal_id: str
    account: str
    list_acv: Decimal
    discount_pct: Decimal
    delivery_cost: Decimal
    implementation_cost: Decimal
    payment_days: int
    probability_pct: Decimal

    @property
    def booked_acv(self) -> Decimal:
        return _usd(self.list_acv * (1 - self.discount_pct / 100))

    @property
    def discount_amount(self) -> Decimal:
        return _usd(self.list_acv - self.booked_acv)

    @property
    def year_one_contribution(self) -> Decimal:
        return _usd(self.booked_acv - self.delivery_cost - self.implementation_cost)

    @property
    def contribution_margin_pct(self) -> Decimal:
        return _usd(100 * self.year_one_contribution / self.booked_acv) if self.booked_acv else Decimal('0')

    @property
    def weighted_contribution(self) -> Decimal:
        return _usd(self.year_one_contribution * self.probability_pct / 100)


def load_deals(path: Path) -> list[Deal]:
    with path.open(newline='', encoding='utf-8-sig') as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames or set(COLUMNS) - set(reader.fieldnames):
            raise ValueError(f'CSV needs columns: {", ".join(COLUMNS)}')
        deals, seen = [], set()
        for rownum, row in enumerate(reader, 2):
            if None in row or any(row.get(k) is None for k in COLUMNS):
                raise ValueError(f'Row {rownum}: malformed CSV')
            deal_id, account = (row[k].strip() for k in ('deal_id', 'account'))
            if not deal_id or not account or deal_id in seen:
                raise ValueError(f'Row {rownum}: missing account/id or duplicate id')
            seen.add(deal_id)
            nums = {k: _number(row[k], f'row {rownum} {k}') for k in
                    ('list_acv', 'discount_pct', 'delivery_cost', 'implementation_cost', 'probability_pct')}
            if any(nums[k] < 0 for k in ('delivery_cost', 'implementation_cost')) or nums['list_acv'] <= 0:
                raise ValueError(f'Row {rownum}: ACV must be positive and costs nonnegative')
            if any(not 0 <= nums[k] <= 100 for k in ('discount_pct', 'probability_pct')):
                raise ValueError(f'Row {rownum}: percentages must be 0..100')
            try:
                days = int(row['payment_days'])
                if str(days) != row['payment_days'].strip() or days < 0 or days > 365:
                    raise ValueError()
            except ValueError as exc:
                raise ValueError(f'Row {rownum}: payment_days must be an integer 0..365') from exc
            deals.append(Deal(deal_id, account, payment_days=days, **nums))
    if not deals:
        raise ValueError('CSV contains no deals')
    return deals


def review(deal: Deal, min_margin: Decimal = Decimal('25'), max_discount: Decimal = Decimal('20'),
           max_payment_days: int = 60) -> dict:
    """Policy flags are independent; score is a routing priority, not default probability."""
    flags = []
    if deal.contribution_margin_pct < min_margin:
        flags.append('MARGIN')
    if deal.discount_pct > max_discount:
        flags.append('DISCOUNT')
    if deal.payment_days > max_payment_days:
        flags.append('PAYMENT_TERMS')
    if deal.year_one_contribution < 0:
        flags.append('LOSS_MAKING')
    score = min(100, 35 * ('LOSS_MAKING' in flags) + 30 * ('MARGIN' in flags)
                + 20 * ('DISCOUNT' in flags) + 15 * ('PAYMENT_TERMS' in flags))
    return {'flags': flags, 'risk_score': score, 'decision': 'REVIEW' if flags else 'PASS'}


def portfolio(deals: list[Deal]) -> dict:
    total = lambda fn: _usd(sum((fn(d) for d in deals), Decimal('0')))
    acv = total(lambda d: d.booked_acv)
    contribution = total(lambda d: d.year_one_contribution)
    return {'count': len(deals), 'list_acv': total(lambda d: d.list_acv),
            'booked_acv': acv, 'discount_amount': total(lambda d: d.discount_amount),
            'delivery_cost': total(lambda d: d.delivery_cost),
            'implementation_cost': total(lambda d: d.implementation_cost),
            'year_one_contribution': contribution,
            'margin_pct': _usd(100 * contribution / acv) if acv else Decimal('0'),
            'weighted_contribution': total(lambda d: d.weighted_contribution),
            'review_count': sum(review(d)['decision'] == 'REVIEW' for d in deals)}


def generate_outputs(deals: list[Deal], outdir: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Font, PatternFill, Alignment
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    outdir.mkdir(parents=True, exist_ok=True)
    p = portfolio(deals)
    wb = Workbook()
    dash = wb.active
    dash.title = 'Dashboard'
    ledger = wb.create_sheet('Deal Ledger')
    policy = wb.create_sheet('Policy & Definitions')
    moneyfmt = '"$"#,##0.00;[Red]("$"#,##0.00)'
    header = ['Deal ID', 'Account', 'List ACV', 'Discount %', 'Booked ACV', 'Discount $',
              'Delivery Cost', 'Implementation Cost', 'Year 1 Contribution', 'Margin %',
              'Payment Days', 'Win Probability %', 'Weighted Contribution', 'Risk Score', 'Decision', 'Flags']
    ledger.append(header)
    for i, d in enumerate(deals, 2):
        verdict = review(d)
        ledger.append([d.deal_id, d.account, float(d.list_acv), float(d.discount_pct)/100,
                       f'=ROUND(C{i}*(1-D{i}),2)', f'=ROUND(C{i}-E{i},2)',
                       float(d.delivery_cost), float(d.implementation_cost),
                       f'=ROUND(E{i}-G{i}-H{i},2)', f'=IF(E{i}=0,0,I{i}/E{i})',
                       d.payment_days, float(d.probability_pct)/100, f'=ROUND(I{i}*L{i},2)',
                       verdict['risk_score'], verdict['decision'], ', '.join(verdict['flags'])])
        for col in ('C', 'E', 'F', 'G', 'H', 'I', 'M'):
            ledger[f'{col}{i}'].number_format = moneyfmt
        for col in ('D', 'J', 'L'):
            ledger[f'{col}{i}'].number_format = '0.0%'
    n = len(deals) + 1
    ledger.freeze_panes = 'C2'
    ledger.auto_filter.ref = f'A1:P{n}'
    ledger.conditional_formatting.add(f'J2:J{n}', CellIsRule(operator='lessThan', formula=['0.25'], fill=PatternFill('solid', fgColor='FCE8E6')))
    dash.append(['DEAL DESK | YEAR-ONE VIEW', 'Value'])
    metrics = [('Deals', f'=COUNTA(\'Deal Ledger\'!A2:A{n})'),
               ('List ACV', f'=SUM(\'Deal Ledger\'!C2:C{n})'),
               ('Booked ACV', f'=SUM(\'Deal Ledger\'!E2:E{n})'),
               ('Discount $', f'=SUM(\'Deal Ledger\'!F2:F{n})'),
               ('Delivery Cost', f'=SUM(\'Deal Ledger\'!G2:G{n})'),
               ('Implementation Cost', f'=SUM(\'Deal Ledger\'!H2:H{n})'),
               ('Year 1 Contribution', f'=SUM(\'Deal Ledger\'!I2:I{n})'),
               ('Contribution Margin %', '=IF(B4=0,0,B8/B4)'),
               ('Probability-Weighted Contribution', f'=SUM(\'Deal Ledger\'!M2:M{n})'),
               ('Deals Requiring Review', f'=COUNTIF(\'Deal Ledger\'!O2:O{n},"REVIEW")')]
    for label, formula in metrics:
        dash.append([label, formula])
    for i in range(3, 13):
        dash[f'B{i}'].number_format = '0.0%' if i == 9 else ("0" if i in (3, 12) else moneyfmt)
    chart = BarChart()
    chart.title = 'ACV and first-year contribution'
    chart.add_data(Reference(dash, min_col=2, min_row=4, max_row=4), from_rows=True)
    chart.add_data(Reference(dash, min_col=2, min_row=8, max_row=8), from_rows=True)
    chart.series[0].title = 'Booked ACV'
    chart.series[1].title = 'Year 1 Contribution'
    dash.add_chart(chart, 'D3')
    policy.append(['Policy / limitation', 'Explanation'])
    for row in [('Contribution', 'Booked annual contract value - annual delivery cost - one-time implementation cost.'),
                ('Margin', 'Year-one contribution / booked annual contract value; 0 if booked ACV is zero.'),
                ('Discount', 'List ACV - booked ACV. NOT equal to recoverable revenue or realized leakage.'),
                ('Weighted contribution', 'Year-one contribution * supplied win probability; scenario estimate only.'),
                ('Approval', 'Review if margin <25%, discount >20%, terms >60d, or negative contribution.'),
                ('Risk score', 'Routing heuristic, NOT statistically calibrated credit or churn risk.'),
                ('Cash timing', 'Payment days trigger a flag; financing costs and cash-flow effects NOT modeled.'),
                ('Excel formulas', 'Open in Excel/LibreOffice to recalculate formula cells and dashboard.')]:
        policy.append(row)
    for ws in wb:
        ws.freeze_panes = ws.freeze_panes or 'A2'
        ws.row_dimensions[1].height = 27
        for cell in ws[1]:
            cell.fill = PatternFill('solid', fgColor='16324F')
            cell.font = Font(color='FFFFFF', bold=True)
            cell.alignment = Alignment(wrap_text=True)
    for col, width in {'A':36,'B':24}.items():
        dash.column_dimensions[col].width = width
        policy.column_dimensions[col].width = 65 if col == 'B' else 27
    for col in 'ABCDEFGHIJKLMNOP':
        ledger.column_dimensions[col].width = 26 if col in ('B', 'P') else 20
    wb.save(outdir / 'deal_desk.xlsx')

    # Each bar reconciles to the contribution, with no probabilistic or implied recovery claim.
    fig, ax = plt.subplots(figsize=(9, 4.6))
    labels = ['List ACV', 'Discount', 'Booked ACV', 'Delivery', 'Implementation', 'Contribution']
    vals = [float(p[k]) for k in ('list_acv','discount_amount','booked_acv','delivery_cost','implementation_cost','year_one_contribution')]
    ax.bar(labels, vals, color=['#16324F','#DC7F53','#307E86','#DFAF63','#A3727A','#439768'])
    ax.set_title('Year-one commercial bridge (USD; bars are amounts, not additive)')
    ax.tick_params(axis='x', rotation=18)
    ax.grid(axis='y', alpha=.2)
    fig.tight_layout()
    fig.savefig(outdir / 'commercial_bridge.png', dpi=160)
    plt.close(fig)
    fig, ax = plt.subplots(figsize=(9, 4.6))
    ax.scatter([float(d.booked_acv) for d in deals], [review(d)['risk_score'] for d in deals],
               c=['#C45151' if review(d)['decision']=='REVIEW' else '#439768' for d in deals], s=75)
    ax.set(xlabel='Booked annual contract value (USD)', ylabel='Policy routing score (0–100)',
           title='Deal size vs policy routing score')
    ax.set_ylim(-5, 105)
    ax.grid(alpha=.2)
    fig.tight_layout()
    fig.savefig(outdir / 'risk_map.png', dpi=160)
    plt.close(fig)

    pdf = canvas.Canvas(str(outdir / 'executive_brief.pdf'), pagesize=(612, 792))
    pdf.setTitle('Deal Desk | Fictional Executive Brief')
    def line(text: str, y: int, size: int = 11):
        pdf.setFillColor(colors.HexColor('#16324F'))
        pdf.setFont('Helvetica', size)
        pdf.drawString(42, y, text)
    line('DEAL DESK | EXECUTIVE BRIEF', 745, 18)
    line('Fictional sample data; scenario analysis, not company performance.', 718, 10)
    details = [('Pipeline deals', str(p['count'])), ('List ACV', f"${p['list_acv']:,.2f}"),
               ('Booked ACV', f"${p['booked_acv']:,.2f}"), ('Discount dollars (not recoverable)', f"${p['discount_amount']:,.2f}"),
               ('Year-one contribution', f"${p['year_one_contribution']:,.2f}"),
               ('Portfolio contribution margin', f"{p['margin_pct']}%"),
               ('Probability-weighted contribution', f"${p['weighted_contribution']:,.2f}"),
               ('Deals requiring human review', str(p['review_count']))]
    for i, (label, val) in enumerate(details):
        line(f'{label}: {val}', 683 - 24*i)
    line('Review priorities (illustrative)', 455, 13)
    for i, d in enumerate(sorted(deals, key=lambda x: review(x)['risk_score'], reverse=True)[:5]):
        line(f"{d.deal_id}: score {review(d)['risk_score']} | {', '.join(review(d)['flags']) or 'PASS'}", 432-i*22, 10)
    line('Actions: inspect low margins; approve discounts and payment terms case by case.', 290, 10)
    line('Discounts are observed vs list price; they are NOT proved revenue leakage.', 272, 10)
    line('No CAC, LTV, churn, payback or financing assumptions in this model.', 254, 10)
    pdf.drawImage(ImageReader(str(outdir / 'commercial_bridge.png')), 38, 32, width=536, height=196)
    pdf.showPage()
    line('DEAL REVIEW MAP', 745, 18)
    line('Score is a deterministic policy priority; it does not predict deal failure.', 717, 10)
    pdf.drawImage(ImageReader(str(outdir / 'risk_map.png')), 34, 390, width=544, height=285)
    line('Policy thresholds: margin <25%; discount >20%; payment terms >60d.', 348, 10)
    line('Loss-making deals are separately flagged. Human approval required.', 328, 10)
    line('Source: input CSV. All calculations are year-one USD, before taxes.', 308, 10)
    pdf.save()


def main() -> None:
    parser = argparse.ArgumentParser(description='Auditable offline deal desk analysis')
    parser.add_argument('--input', type=Path, default=Path(__file__).parent/'examples'/'fictional_deals.csv')
    parser.add_argument('--out', type=Path, default=Path('output'))
    args = parser.parse_args()
    try:
        deals = load_deals(args.input)
        generate_outputs(deals, args.out)
        print(f'Analyzed {len(deals)} deals. Wrote Excel, PDF and two PNG charts to {args.out.resolve()}')
    except (ValueError, OSError) as exc:
        parser.exit(2, f'Error: {exc}\n')


if __name__ == '__main__':
    main()
