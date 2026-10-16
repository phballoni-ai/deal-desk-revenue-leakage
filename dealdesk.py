"""Offline B2B deal desk analysis; USD, annual ACV, one-time setup cost."""
from __future__ import annotations
import argparse
import csv
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

FIELDS = ('deal_id','account','list_acv','discount_pct','delivery_cost','implementation_cost','payment_days','probability_pct')
CENT = Decimal('.01')

def dec(value, name):
    try:
        x = Decimal(str(value))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f'{name}: invalid number') from exc
    if not x.is_finite():
        raise ValueError(f'{name}: finite number required')
    return x

def usd(x):
    return x.quantize(CENT, rounding=ROUND_HALF_UP)

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
    def booked_acv(self):
        return usd(self.list_acv * (1 - self.discount_pct / 100))
    @property
    def discount_amount(self):
        return usd(self.list_acv - self.booked_acv)
    @property
    def year_one_contribution(self):
        return usd(self.booked_acv - self.delivery_cost - self.implementation_cost)
    @property
    def contribution_margin_pct(self):
        return usd(100 * self.year_one_contribution / self.booked_acv) if self.booked_acv else Decimal('0')
    @property
    def weighted_contribution(self):
        return usd(self.year_one_contribution * self.probability_pct / 100)

def load_deals(path):
    with Path(path).open(newline='', encoding='utf-8-sig') as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames or set(FIELDS) - set(reader.fieldnames):
            raise ValueError('Missing CSV columns: ' + ', '.join(sorted(set(FIELDS)-set(reader.fieldnames or []))))
        deals, seen = [], set()
        for rownum, row in enumerate(reader, 2):
            if None in row or any(row.get(k) is None for k in FIELDS):
                raise ValueError(f'Row {rownum}: malformed CSV')
            did, account = row['deal_id'].strip(), row['account'].strip()
            if not did or not account or did in seen:
                raise ValueError(f'Row {rownum}: blank or duplicate identifier/account')
            seen.add(did)
            vals = {k:dec(row[k], f'row {rownum} {k}') for k in ('list_acv','discount_pct','delivery_cost','implementation_cost','probability_pct')}
            if vals['list_acv'] <= 0 or vals['delivery_cost'] < 0 or vals['implementation_cost'] < 0:
                raise ValueError(f'Row {rownum}: ACV must be positive, costs nonnegative')
            if not 0 <= vals['discount_pct'] <= 100 or not 0 <= vals['probability_pct'] <= 100:
                raise ValueError(f'Row {rownum}: percentage outside 0..100')
            try:
                days = int(row['payment_days'])
                if str(days) != row['payment_days'].strip() or not 0 <= days <= 365:
                    raise ValueError()
            except ValueError as exc:
                raise ValueError(f'Row {rownum}: payment days must be integer 0..365') from exc
            deals.append(Deal(did, account, payment_days=days, **vals))
    if not deals:
        raise ValueError('CSV contains no deals')
    return deals

def review(d):
    flags = []
    if d.contribution_margin_pct < 25: flags.append('MARGIN')
    if d.discount_pct > 20: flags.append('DISCOUNT')
    if d.payment_days > 60: flags.append('PAYMENT_TERMS')
    if d.year_one_contribution < 0: flags.append('LOSS_MAKING')
    score = min(100, 30*('MARGIN' in flags)+20*('DISCOUNT' in flags)+15*('PAYMENT_TERMS' in flags)+35*('LOSS_MAKING' in flags))
    return {'flags':flags, 'risk_score':score, 'decision':'REVIEW' if flags else 'PASS'}

def portfolio(deals):
    total = lambda attr: usd(sum((getattr(d, attr) for d in deals), Decimal(0)))
    booked = total('booked_acv')
    contrib = total('year_one_contribution')
    return {'count':len(deals), 'list_acv':total('list_acv'), 'booked_acv':booked,
            'discount_amount':total('discount_amount'), 'delivery_cost':total('delivery_cost'),
            'implementation_cost':total('implementation_cost'), 'year_one_contribution':contrib,
            'margin_pct':usd(100*contrib/booked) if booked else Decimal(0),
            'weighted_contribution':total('weighted_contribution'),
            'review_count':sum(review(d)['decision']=='REVIEW' for d in deals)}

def generate_outputs(deals, outdir):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Font, PatternFill
    from openpyxl.formatting.rule import CellIsRule
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    p = portfolio(deals)
    wb = Workbook()
    dash = wb.active
    dash.title = 'Dashboard'
    ledger = wb.create_sheet('Deal Ledger')
    policy = wb.create_sheet('Policy & Definitions')
    ledger.append(['Deal ID','Account','List ACV','Discount %','Booked ACV','Discount $',
                   'Delivery Cost','Implementation Cost','Year 1 Contribution','Margin %',
                   'Payment Days','Win Probability %','Weighted Contribution','Risk Score','Decision','Flags'])
    money = '"$"#,##0.00;[Red]("$"#,##0.00)'
    for i,d in enumerate(deals, 2):
        v = review(d)
        ledger.append([d.deal_id,d.account,float(d.list_acv),float(d.discount_pct/100),
                       f'=ROUND(C{i}*(1-D{i}),2)',f'=ROUND(C{i}-E{i},2)',float(d.delivery_cost),
                       float(d.implementation_cost),f'=ROUND(E{i}-G{i}-H{i},2)',f'=IF(E{i}=0,0,I{i}/E{i})',
                       d.payment_days,float(d.probability_pct/100),f'=ROUND(I{i}*L{i},2)',
                       v['risk_score'],v['decision'],', '.join(v['flags'])])
        for col in ('C','E','F','G','H','I','M'): ledger[f'{col}{i}'].number_format = money
        for col in ('D','J','L'): ledger[f'{col}{i}'].number_format = '0.0%'
    n = len(deals)+1
    ledger.freeze_panes = 'C2'
    ledger.auto_filter.ref = f'A1:P{n}'
    ledger.conditional_formatting.add(f'J2:J{n}', CellIsRule(operator='lessThan', formula=['0.25'], fill=PatternFill('solid', fgColor='FCE8E6')))
    dash.append(['DEAL DESK | YEAR ONE', 'Value'])
    for label, formula in [
        ('Deals',f"=COUNTA('Deal Ledger'!A2:A{n})"),
        ('List ACV',f"=SUM('Deal Ledger'!C2:C{n})"),
        ('Booked ACV',f"=SUM('Deal Ledger'!E2:E{n})"),
        ('Discount $',f"=SUM('Deal Ledger'!F2:F{n})"),
        ('Delivery Cost',f"=SUM('Deal Ledger'!G2:G{n})"),
        ('Implementation Cost',f"=SUM('Deal Ledger'!H2:H{n})"),
        ('Year 1 Contribution',f"=SUM('Deal Ledger'!I2:I{n})"),
        ('Contribution Margin %','=IF(B5=0,0,B9/B5)'),
        ('Weighted Contribution',f"=SUM('Deal Ledger'!M2:M{n})"),
        ('Deals Requiring Review',f'=COUNTIF(\'Deal Ledger\'!O2:O{n},"REVIEW")')]:
        dash.append([label,formula])
    for i in range(3,13):
        dash[f'B{i}'].number_format = '0.0%' if i==10 else ('0' if i in (3,12) else money)
    chart = BarChart()
    chart.title = 'Year-one contribution by deal'
    chart.y_axis.title = 'USD'
    chart.add_data(Reference(ledger, min_col=9, min_row=1, max_row=n), titles_from_data=True)
    chart.set_categories(Reference(ledger, min_col=1, min_row=2, max_row=n))
    dash.add_chart(chart, 'D3')
    policy.append(['Policy / limitation','Explanation'])
    for label, explanation in [
        ('Contribution','Booked annual ACV minus annual delivery and one-time implementation costs.'),
        ('Margin','Contribution / booked ACV; zero where booked ACV is zero.'),
        ('Discount','List minus booked ACV. NOT proven recoverable revenue.'),
        ('Weighted contribution','Contribution multiplied by user-supplied probability; scenario only.'),
        ('Review','Margin <25%, discount >20%, payment days >60, or negative contribution.'),
        ('Score','Explainable routing priority, not a calibrated prediction.'),
        ('Cash timing','Terms flag only; no financing costs or cash-flow forecast.'),
        ('Excel','Open in Excel or LibreOffice to calculate formula cells.')]: policy.append([label, explanation])
    for ws in wb:
        ws.freeze_panes = ws.freeze_panes or 'A2'
        for cell in ws[1]:
            cell.fill = PatternFill('solid', fgColor='16324F')
            cell.font = Font(color='FFFFFF', bold=True)
    dash.column_dimensions['A'].width = 35
    dash.column_dimensions['B'].width = 24
    policy.column_dimensions['A'].width = 27
    policy.column_dimensions['B'].width = 90
    for col in 'ABCDEFGHIJKLMNOP': ledger.column_dimensions[col].width = 28 if col in ('B','P') else 21
    wb.save(outdir/'deal_desk.xlsx')

    fig, ax = plt.subplots(figsize=(9,4.6))
    labels = ['List ACV','Discount','Booked ACV','Delivery','Implementation','Contribution']
    keys = ('list_acv','discount_amount','booked_acv','delivery_cost','implementation_cost','year_one_contribution')
    ax.bar(labels,[float(p[k]) for k in keys],color=['#16324F','#DC7F53','#307E86','#DFAF63','#A3727A','#439768'])
    ax.set_title('Year-one commercial amounts (USD; bars are not additive)')
    ax.tick_params(axis='x',rotation=18)
    ax.grid(axis='y',alpha=.2)
    fig.tight_layout()
    fig.savefig(outdir/'commercial_bridge.png',dpi=160)
    plt.close(fig)
    fig, ax = plt.subplots(figsize=(9,4.6))
    ax.scatter([float(d.booked_acv) for d in deals],[review(d)['risk_score'] for d in deals],
               c=['#C45151' if review(d)['flags'] else '#439768' for d in deals],s=75)
    ax.set(xlabel='Booked ACV (USD)',ylabel='Policy routing score (0–100)',title='Deal size vs routing priority')
    ax.set_ylim(-5,105)
    ax.grid(alpha=.2)
    fig.tight_layout()
    fig.savefig(outdir/'risk_map.png',dpi=160)
    plt.close(fig)

    pdf = canvas.Canvas(str(outdir/'executive_brief.pdf'), pagesize=(612,792))
    pdf.setTitle('Deal Desk | Fictional Executive Brief')
    def write(text,y,size=10):
        pdf.setFont('Helvetica',size)
        pdf.drawString(42,y,text)
    write('DEAL DESK | EXECUTIVE BRIEF',745,18)
    write('Fictional sample data. Scenario analysis, not company performance.',718)
    entries = [('Deals',p['count']),('List ACV',f"${p['list_acv']:,.2f}"),
               ('Booked ACV',f"${p['booked_acv']:,.2f}"),('Discount (not recoverable)',f"${p['discount_amount']:,.2f}"),
               ('Year-one contribution',f"${p['year_one_contribution']:,.2f}"),
               ('Contribution margin',f"{p['margin_pct']}%"),
               ('Probability-weighted contribution',f"${p['weighted_contribution']:,.2f}"),
               ('Deals requiring review',p['review_count'])]
    for i,(label,value) in enumerate(entries): write(f'{label}: {value}',680-i*24)
    write('HIGHEST REVIEW PRIORITIES',450,13)
    for i,d in enumerate(sorted(deals,key=lambda x:review(x)['risk_score'],reverse=True)[:5]):
        write(f"{d.deal_id}: score {review(d)['risk_score']} | {', '.join(review(d)['flags']) or 'PASS'}",427-i*22)
    write('Action: human review of thin margins, discount approvals and long terms.',292)
    write('Discounts vs list price are NOT proven revenue leakage.',272)
    write('No CAC, LTV, churn, payback or financing costs are modeled.',252)
    pdf.drawImage(ImageReader(str(outdir/'commercial_bridge.png')),38,32,width=536,height=195)
    pdf.showPage()
    write('DEAL REVIEW MAP',745,18)
    write('Routing score is deterministic; it does not predict deal failure.',715)
    pdf.drawImage(ImageReader(str(outdir/'risk_map.png')),34,390,width=544,height=285)
    write('Illustrative policy: margin <25%; discount >20%; payment days >60.',345)
    write('Loss-making contracts are separately flagged. Human decision required.',325)
    write('Source: input CSV. All amounts are first-year USD, before tax.',305)
    pdf.save()

def main():
    parser = argparse.ArgumentParser(description='Auditable offline deal desk')
    parser.add_argument('--input',type=Path,default=Path(__file__).parent/'examples'/'fictional_deals.csv')
    parser.add_argument('--out',type=Path,default=Path('output'))
    args = parser.parse_args()
    try:
        deals = load_deals(args.input)
        generate_outputs(deals,args.out)
        print(f'Analyzed {len(deals)} deals; Excel, PDF and two PNGs saved in {args.out.resolve()}')
    except (ValueError,OSError) as exc:
        parser.exit(2,f'Error: {exc}\n')

if __name__ == '__main__':
    main()
