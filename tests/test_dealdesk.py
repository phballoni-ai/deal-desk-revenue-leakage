import csv
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from dealdesk import Deal, load_deals, portfolio, review


class DealDeskTests(unittest.TestCase):
    def deal(self, **kwargs):
        data = dict(deal_id='D1', account='Demo', list_acv=Decimal('100'),
                    discount_pct=Decimal('20'), delivery_cost=Decimal('40'),
                    implementation_cost=Decimal('10'), payment_days=30,
                    probability_pct=Decimal('50'))
        data.update(kwargs)
        return Deal(**data)

    def test_reconciliation_and_threshold(self):
        d = self.deal()
        self.assertEqual(d.booked_acv, Decimal('80.00'))
        self.assertEqual(d.discount_amount, Decimal('20.00'))
        self.assertEqual(d.year_one_contribution, Decimal('30.00'))
        self.assertEqual(d.contribution_margin_pct, Decimal('37.50'))
        self.assertEqual(d.weighted_contribution, Decimal('15.00'))
        self.assertEqual(review(d)['decision'], 'PASS')
        self.assertEqual(portfolio([d])['year_one_contribution'], Decimal('30.00'))

    def test_loss_making_route(self):
        d = self.deal(delivery_cost=Decimal('85'), payment_days=90)
        self.assertEqual(review(d)['decision'], 'REVIEW')
        self.assertIn('LOSS_MAKING', review(d)['flags'])
        self.assertIn('PAYMENT_TERMS', review(d)['flags'])

    def test_csv_invalid_percentage_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'deals.csv'
            with path.open('w', newline='', encoding='utf-8') as h:
                w = csv.writer(h)
                w.writerow(['deal_id','account','list_acv','discount_pct','delivery_cost',
                            'implementation_cost','payment_days','probability_pct'])
                w.writerow(['D1','Demo','100','101','40','10','30','50'])
            with self.assertRaises(ValueError):
                load_deals(path)

    def test_demo_portfolio_reconciles(self):
        deals = load_deals(Path(__file__).resolve().parents[1] / 'examples' / 'fictional_deals.csv')
        p = portfolio(deals)
        self.assertEqual(p['list_acv'] - p['discount_amount'], p['booked_acv'])
        self.assertEqual(p['booked_acv']-p['delivery_cost']-p['implementation_cost'], p['year_one_contribution'])
        self.assertEqual(p['count'], 12)


if __name__ == '__main__':
    unittest.main()
