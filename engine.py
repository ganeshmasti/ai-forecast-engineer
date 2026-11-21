"""A small, auditable forecast consolidation engine. All financial values are synthetic.

This is an educational prototype: it does NOT execute Excel formulas, macros or
external links and it never writes to SAP. Money is computed with Decimal.
"""
from __future__ import annotations

import csv
import io
import re
from collections import Counter
from dataclasses import dataclass, asdict
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from difflib import get_close_matches
from pathlib import Path

import openpyxl

Q = Decimal('0.01')
FIELDS = ('period','region','subregion','bu','customer_id','sku','units','price_local','frontend_rate','contra_rate','volume_rate','ccy','gross_formula','net_formula')
ALIASES = {
    'period': ('month','forecast month','period'),
    'region': ('region',),
    'subregion': ('subregion','sub region'),
    'bu': ('bu','business unit'),
    'customer_id': ('customer id','account code','customer','customer code'),
    'sku': ('sku','product code','product id'),
    'units': ('units','forecast qty','forecast quantity','quantity'),
    'price_local': ('unit price','asp local','unit price local','average selling price'),
    'frontend_rate': ('front-end %','fep discount','frontend %','front-end discount'),
    'contra_rate': ('contra %','contra rebate','contra rate'),
    'volume_rate': ('volume %','tier rebate','volume rebate'),
    'ccy': ('ccy','currency'),
    'gross_formula': ('gross revenue','gross sales'),
    'net_formula': ('net revenue','net sales'),
}
REQUIRED = {'period','bu','customer_id','sku','units','price_local','frontend_rate','contra_rate','volume_rate','ccy'}
SEVERITY = {'MISSING_CUSTOMER':'High','NEGATIVE_UNITS':'High','REBATE_OUT_OF_RANGE':'High','DUPLICATE_KEY':'High','FORMULA_DRIFT':'Medium','UNKNOWN_CURRENCY':'High','UNKNOWN_REGION':'Medium','INVALID_VALUE':'High','UNMAPPED_FIELDS':'High'}


def decimal(v):
    if isinstance(v, Decimal): return v
    if v is None or isinstance(v,str) and not v.strip(): return Decimal(0)
    return Decimal(str(v))


def money(v): return decimal(v).quantize(Q, rounding=ROUND_HALF_UP)


def clean(v): return re.sub(r'[^a-z0-9]+', ' ',str(v or '').lower()).strip()


def guess_mapping(headers):
    """Deterministic recognition; LLM suggestions can be layered on later."""
    mapped={}; detail=[]; used=set()
    for j,h in enumerate(headers):
        if h is None: continue
        norm=clean(h)
        for field, names in ALIASES.items():
            if field in used:continue
            if norm in {clean(a) for a in names}:
                mapped[field]=j; used.add(field)
                detail.append({'source':str(h),'target':field,'method':'approved alias','confidence':1.0})
                break
        else:
            match=get_close_matches(norm, [clean(a) for v in ALIASES.values() for a in v], n=1, cutoff=.78)
            if match:
                detail.append({'source':str(h),'target':f'SUGGESTED: {match[0]}','method':'review required','confidence':.7})
            else:
                detail.append({'source':str(h),'target':'Unmapped','method':'manual review','confidence':0.0})
    return mapped,detail


def find_header(sheet):
    best=(-1,None,None)
    for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row or 12, 12), values_only=True):
        mapping,_=guess_mapping(row)
        n=len(REQUIRED & set(mapping))
        if n>best[0]:best=(n, row, mapping)
    if best[0]<6: raise ValueError('No compatible forecast table detected (need at least 6 recognized required headers).')
    for i,row in enumerate(sheet.iter_rows(min_row=1,max_row=min(sheet.max_row or 12,12),values_only=True),1):
        if tuple(row)==tuple(best[1]): return i,list(best[1]),best[2]
    raise ValueError('Header not found')


def norm_formula(v):
    if not isinstance(v,str) or not v.startswith('='): return ''
    return re.sub(r'\$','',re.sub(r'\s+','',v.upper()))


def formula_expected(mapping,r,measure):
    def cell(field):
        return f"{openpyxl.utils.get_column_letter(mapping[field]+1)}{r}"
    try:
        if measure=='gross_formula':return '=' + cell('units')+'*'+cell('price_local')
        return '='+cell('gross_formula')+'*(1-'+cell('frontend_rate')+'-'+cell('contra_rate')+'-'+cell('volume_rate')+')'
    except KeyError: return ''


def load_fx(path):
    with open(path, newline='', encoding='utf-8') as f:
        return {(r['period'],r['currency']):decimal(r['usd_per_local_unit']) for r in csv.DictReader(f)}


@dataclass
class ScanResult:
    lines:list[dict]
    issues:list[dict]
    mappings:list[dict]
    sources:list[dict]


def scan_workbooks(items:list[tuple[str,bytes]], fx:dict):
    lines=[]; issues=[]; mappings=[]; sources=[]
    for filename, content in items:
        safe_name=Path(filename).name
        if len(content)>8*1024*1024:
            raise ValueError(f'{safe_name}: file larger than 8 MB demo limit')
        # openpyxl does not execute macros / formula calculations. Keep links disabled.
        try:
            wb=openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=False, keep_links=False)
        except Exception as e:
            raise ValueError(f'{safe_name}: unable to parse XLSX: {e}') from e
        if 'Submission' not in wb.sheetnames:
            raise ValueError(f'{safe_name}: Submission worksheet not found')
        sheet=wb['Submission']
        head_idx,headers,mapping=find_header(sheet)
        suggestions=guess_mapping(headers)[1]
        mappings += [dict(file=safe_name,**m) for m in suggestions]
        missing=REQUIRED-set(mapping)
        if missing:issues.append(issue(safe_name,head_idx,'UNMAPPED_FIELDS', 'Missing required columns: '+', '.join(sorted(missing))))
        count=0
        seen={}
        region_from_filename=safe_name.split('_')[0].upper() if safe_name.split('_')[0].upper() in ('AMER','EMEA','APAC') else None
        subregion_from_filename=safe_name.split('_')[1] if '_' in safe_name else ''
        for ridx,row in enumerate(sheet.iter_rows(min_row=head_idx+1,values_only=True),head_idx+1):
            if not any(v is not None and v!='' for v in row):continue
            def cell(field):
                j=mapping.get(field)
                return row[j] if j is not None and j<len(row) else None
            region=str(cell('region') or region_from_filename or '').strip().upper()
            subregion=str(cell('subregion') or subregion_from_filename).strip()
            customer=str(cell('customer_id') or '').strip()
            period=str(cell('period') or '').strip()
            bu=str(cell('bu') or '').strip()
            sku=str(cell('sku') or '').strip()
            ccy=str(cell('ccy') or '').strip().upper()
            line_id=f'{safe_name}|{ridx}'
            local_issues=[]
            for field in ['units','price_local','frontend_rate','contra_rate','volume_rate']:
                try: decimal(cell(field))
                except (InvalidOperation,ValueError):local_issues.append(('INVALID_VALUE',f'{field}: {cell(field)!r}'))
            if local_issues: 
                for code,msg in local_issues:issues.append(issue(safe_name,ridx,code,msg))
                continue
            units=decimal(cell('units')); price=decimal(cell('price_local'))
            fep=decimal(cell('frontend_rate'));contra=decimal(cell('contra_rate'));volume=decimal(cell('volume_rate'))
            gross=money(units*price)
            net=money(gross*(Decimal(1)-fep-contra-volume))
            rate=fx.get((period,ccy))
            usd=money(net*rate) if rate is not None else None
            gross_usd=money(gross*rate) if rate is not None else None
            lineno=dict(line_id=line_id,file=safe_name,row=ridx,period=period,region=region,subregion=subregion,
                        bu=bu,customer_id=customer,sku=sku,ccy=ccy,units=float(units),price_local=float(price),
                        frontend_rate=float(fep),contra_rate=float(contra),volume_rate=float(volume),
                        gross_local=float(gross),net_local=float(net),gross_usd=float(gross_usd) if gross_usd is not None else None,
                        net_usd=float(usd) if usd is not None else None,
                        gross_formula=str(cell('gross_formula') or ''),net_formula=str(cell('net_formula') or ''))
            lines.append(lineno);count+=1
            if not customer:issues.append(issue(safe_name,ridx,'MISSING_CUSTOMER','Customer identifier is blank',line_id))
            if not region:issues.append(issue(safe_name,ridx,'UNKNOWN_REGION','Region could not be determined',line_id))
            if units<0:issues.append(issue(safe_name,ridx,'NEGATIVE_UNITS',f'Forecast units {units} below zero',line_id))
            for field,v in [('front-end',fep),('contra',contra),('volume',volume)]:
                if v<0 or v>Decimal('.50'):
                    issues.append(issue(safe_name,ridx,'REBATE_OUT_OF_RANGE',f'{field} rate {v:.2%} outside 0%-50% demo policy',line_id))
            if rate is None:issues.append(issue(safe_name,ridx,'UNKNOWN_CURRENCY',f'Missing synthetic budget FX rate {ccy}/{period}',line_id))
            # Within-submission duplicate detection only. Cross-regional overlap needs ownership mapping.
            key=(period,bu,customer,sku)
            if all(key) and key in seen:issues.append(issue(safe_name,ridx,'DUPLICATE_KEY',f'Same key as Excel row {seen[key]}',line_id))
            seen[key]=ridx
            for field in ('gross_formula','net_formula'):
                val=cell(field)
                if field in mapping and isinstance(val,str) and val.startswith('='):
                    expected=formula_expected(mapping,ridx,field)
                    if expected and norm_formula(val)!=norm_formula(expected):
                        issues.append(issue(safe_name,ridx,'FORMULA_DRIFT',f'{field}: {val} differs from expected {expected}',line_id))
                elif field in mapping:
                    issues.append(issue(safe_name,ridx,'FORMULA_DRIFT',f'{field}: missing/unrecognized formula',line_id))
        sources.append({'file':safe_name,'rows':count,'mapping_coverage':round(len(REQUIRED & set(mapping))/len(REQUIRED)*100,1),'region':region_from_filename or 'UNKNOWN'})
        wb.close()
    flagged={i['line_id'] for i in issues if i.get('line_id')}
    for line in lines:line['status']='REVIEW' if line['line_id'] in flagged else 'READY'
    return ScanResult(lines,issues,mappings,sources)


def issue(file,row,code,detail,line_id=None):
    return {'file':file,'row':row,'code':code,'severity':SEVERITY.get(code,'Medium'),'detail':detail,'line_id':line_id or f'{file}|{row}'}


def scenario(lines, demand_pct:float, rebate_bps:float,fx):
    # Delta is applied only to ready records; UI labels this explicitly.
    qfactor=Decimal(1)+decimal(demand_pct)/Decimal(100)
    rebate_delta=decimal(rebate_bps)/Decimal(10000)
    baseline=scenario_value=Decimal(0)
    for r in lines:
        if r['status']!='READY':continue
        rate=fx.get((r['period'],r['ccy']))
        if rate is None:continue
        units=decimal(r['units']);price=decimal(r['price_local'])
        gross=money(units*price)
        base_net=money(gross*(Decimal(1)-decimal(r['frontend_rate'])-decimal(r['contra_rate'])-decimal(r['volume_rate'])))
        alt_gross=money(gross*qfactor)
        alt_net=money(alt_gross*(Decimal(1)-decimal(r['frontend_rate'])-decimal(r['contra_rate'])-decimal(r['volume_rate'])-rebate_delta))
        baseline += money(base_net*rate)
        scenario_value += money(alt_net*rate)
    return float(baseline),float(scenario_value)
