import csv
import sys
from collections import Counter
from decimal import Decimal
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from engine import load_fx, scan_workbooks, scenario


def test_fixture_quality():
    fx=load_fx(ROOT/'sample_files/fx_rates.csv')
    result=scan_workbooks([(p.name,p.read_bytes()) for p in (ROOT/'sample_files').glob('*.xlsx')],fx)
    assert len(result.lines)==216
    assert len(result.sources)==6
    assert len(result.issues)==5
    assert set(i['code'] for i in result.issues)=={'DUPLICATE_KEY','MISSING_CUSTOMER','NEGATIVE_UNITS','REBATE_OUT_OF_RANGE','FORMULA_DRIFT'}
    assert all(s['mapping_coverage']==100 for s in result.sources)
    assert sum(r['status']=='READY' for r in result.lines)==211


def test_financial_ground_truth():
    fx=load_fx(ROOT/'sample_files/fx_rates.csv')
    result=scan_workbooks([(p.name,p.read_bytes()) for p in (ROOT/'sample_files').glob('*.xlsx')],fx)
    truth_file=ROOT/'tests/ground_truth.csv'
    with truth_file.open(newline='') as f: oracle={(r['source_file'],int(r['excel_row'])):r for r in csv.DictReader(f)}
    assert len(oracle)==216
    for r in result.lines:
        truth=oracle[(r['file'],r['row'])]
        assert abs(Decimal(str(r['net_usd']))-Decimal(truth['expected_net_usd']))<=Decimal('.02')
        assert abs(Decimal(str(r['gross_usd']))-Decimal(truth['expected_gross_usd']))<=Decimal('.02')


def test_scenario_neutral():
    fx=load_fx(ROOT/'sample_files/fx_rates.csv')
    result=scan_workbooks([(p.name,p.read_bytes()) for p in (ROOT/'sample_files').glob('*.xlsx')],fx)
    baseline,changed=scenario(result.lines,0,0,fx)
    assert baseline==changed
    baseline,changed=scenario(result.lines,10,0,fx)
    assert changed>baseline
