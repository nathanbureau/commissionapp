# ================================================================
# bureau booths commission calculator  |  app.py
# ================================================================

import re, io, zipfile, random, time
from datetime import date, timedelta

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from sqlalchemy import create_engine, text

st.set_page_config(
    page_title='Bureau Commission',
    page_icon='💷',
    layout='wide',
    initial_sidebar_state='expanded',
)


# ================================================================
# section 1: configuration
# ================================================================

REPS = {
    'Joe Marshall':            {'level': 'AE1',     'region': 'AUS',   'currency': 'AUD'},
    'Jack Eagle':              {'level': 'AE1',     'region': 'UK',    'currency': 'GBP'},
    'Jay Hudon':               {'level': 'AE1',     'region': 'CAN',   'currency': 'CAD'},
    'Kevin Craig':             {'level': 'AE1',     'region': 'CAN',   'currency': 'CAD'},
    'John Yienger':            {'level': 'AE1',     'region': 'US TX', 'currency': 'USD'},
    'Tyler Givens':            {'level': 'AE1',     'region': 'US NY', 'currency': 'USD'},
    'Brandon Brown':           {'level': 'AE1',     'region': 'CAN',   'currency': 'CAD'},
    'Brett Robinson':          {'level': 'AE1',     'region': 'US TX', 'currency': 'USD'},
    'Connor Harper':           {'level': 'AE1',     'region': 'US NY', 'currency': 'USD'},
    'Jett Laws':               {'level': 'AE1',     'region': 'US TX', 'currency': 'USD'},
    'Branson Wilson':          {'level': 'AE1',     'region': 'US TX', 'currency': 'USD'},
    'Peter Holt':              {'level': 'AE1',     'region': 'CAN',   'currency': 'CAD'},
    'Reuben Zuidhof':          {'level': 'AE1',     'region': 'CAN',   'currency': 'CAD'},
    'Sam Quennell':            {'level': 'AE1',     'region': 'AUS',   'currency': 'AUD'},
    'Adam Morgan':             {'level': 'AE1',     'region': 'US NY', 'currency': 'USD'},
    'Grace Aicardi':           {'level': 'AE1',     'region': 'AUS',   'currency': 'AUD'},
    'Anton Weininger':         {'level': 'AE1',     'region': 'UK',    'currency': 'GBP'},
    'Nik Balashov':            {'level': 'AE1',     'region': 'UK',    'currency': 'GBP'},
    'Harm Magis':              {'level': 'AE1',     'region': 'UK',    'currency': 'GBP'},
    'Megan Scholz':            {'level': 'AE1',     'region': 'AUS',   'currency': 'AUD'},
    'Alex Kretowicz':          {'level': 'AE1',     'region': 'CAN',   'currency': 'CAD'},
    'Harry Steele':            {'level': 'AE2',     'region': 'AUS',   'currency': 'AUD'},
    'Joshua Cherry':           {'level': 'AE2',     'region': 'CAN',   'currency': 'CAD'},
    'Lachie Topp':             {'level': 'AE2',     'region': 'UK',    'currency': 'GBP'},
    'Ryan Lenz':               {'level': 'AE2',     'region': 'US NY', 'currency': 'USD'},
    'Kyle Harms':              {'level': 'AE2',     'region': 'CAN',   'currency': 'CAD'},
    'Alex DeRenzis':           {'level': 'CAM',     'region': 'CAN',   'currency': 'CAD'},
    'Graeme Hodson-Walker':    {'level': 'CAM',     'region': 'CAN',   'currency': 'CAD'},
    'Natasha Lewis':           {'level': 'CAM',     'region': 'US NY', 'currency': 'USD'},
    'Ella Horner':             {'level': 'CAM',     'region': 'AUS',   'currency': 'AUD'},
    'Grace Randell':           {'level': 'CAM',     'region': 'AUS',   'currency': 'AUD'},
    'Halle Smith':             {'level': 'CAM',     'region': 'US TX', 'currency': 'USD'},
    'Ragan Sims':              {'level': 'CAM',     'region': 'US TX', 'currency': 'USD'},
    'Kyle McCulloch':          {'level': 'CAM',     'region': 'US NY', 'currency': 'USD'},
    'Danielle Celentano':      {'level': 'CAM',     'region': 'CAN',   'currency': 'CAD'},
    'Rachelle Sampson':        {'level': 'CAM',     'region': 'CAN',   'currency': 'CAD'},
    'Kathryn Nicholson-Brown': {'level': 'CAM',     'region': 'UK',    'currency': 'GBP'},
    'Nicole Murphy':           {'level': 'CAM',     'region': 'CAN',   'currency': 'CAD'},
    'Marta Menendez':          {'level': 'CAM',     'region': 'UK',    'currency': 'GBP'},
    'Marcus De Verteuil':      {'level': 'SPECIAL', 'region': 'CAN',   'currency': 'CAD'},
    'Geddes Carrington':       {'level': 'MANAGER', 'region': 'US',    'currency': 'USD'},
}

OWNER_MAP = {
    'josh cherry':         'Joshua Cherry',
    'marcus de verteuil':  'Marcus De Verteuil',
    'jett (phillip) laws': 'Jett Laws',
}

_CHANNEL_KEYS = [
    ('dealer',   ['dealer', 'architect', 'rfp', 'tender']),
    ('outbound', ['outbound']),
    ('return',   ['return', 'expansion', 'customer']),
]

def detect_channel(raw):
    if not raw or str(raw).strip().lower() in ('', 'none', 'nan', '(no value)'):
        return 'inbound'
    s = str(raw).lower()
    for ch, kws in _CHANNEL_KEYS:
        if any(k in s for k in kws):
            return ch
    return 'inbound'

def clean_owner(raw):
    if not raw:
        return None
    s = str(raw).replace('(Deactivated User)', '').strip()
    return OWNER_MAP.get(s.lower(), s)

# commission rates: {region: {channel: rate}}
_AE1 = {
    'CAN':   {'inbound': .020, 'return': .020, 'outbound': .050, 'dealer': .010},
    'UK':    {'inbound': .020, 'return': .020, 'outbound': .050, 'dealer': .010},
    'US TX': {'inbound': .020, 'return': .020, 'outbound': .050, 'dealer': .010},
    'AUS':   {'inbound': .015, 'return': .015, 'outbound': .050, 'dealer': .010},
    'US NY': {'inbound': .015, 'return': .015, 'outbound': .050, 'dealer': .010},
}
_AE2 = {
    'AUS':   {'inbound': .015, 'return': .030, 'outbound': .050, 'dealer': .010},
    'CAN':   {'inbound': .020, 'return': .030, 'outbound': .050, 'dealer': .010},
    'US TX': {'inbound': .020, 'return': .030, 'outbound': .050, 'dealer': .010},
    'US NY': {'inbound': .020, 'return': .025, 'outbound': .060, 'dealer': .010},
    'UK':    {'inbound': .015, 'return': .020, 'outbound': .050, 'dealer': .010},
}
_CAM     = {'inbound': .000, 'return': .000, 'outbound': .010, 'dealer': .000}
_SPECIAL = {'Marcus De Verteuil': {'inbound': .020, 'return': .020, 'outbound': .050, 'dealer': .000}}

def get_comm_rate(name, level, region, channel):
    if level == 'MANAGER':
        return 0.0
    if name in _SPECIAL:
        return _SPECIAL[name].get(channel, 0.0)
    if level == 'CAM':
        return _CAM.get(channel, 0.0)
    if level == 'AE2':
        return _AE2.get(region, {}).get(channel, 0.0)
    if level == 'AE1':
        return _AE1.get(region, {}).get(channel, 0.0)
    return 0.0

# accelerator tiers: list of (threshold, rate) ascending
_AE1_T = {
    'CAN':   [(273600,.0020),(307800,.0050),(342000,.0140),(376200,.0160),(410400,.0180),(444600,.0200)],
    'AUS':   [(259200,.0020),(291600,.0050),(324000,.0120),(356400,.0140),(388800,.0180),(421200,.0200)],
    'UK':    [(180000,.0020),(202500,.0050),(225000,.0140),(247500,.0160),(270000,.0180),(292500,.0200)],
    'US TX': [(302400,.0020),(340200,.0050),(378000,.0140),(415800,.0160),(453600,.0180),(491400,.0200)],
    'US NY': [(302400,.0020),(340200,.0050),(378000,.0140),(415800,.0160),(453600,.0180),(491400,.0200)],
}
_AE2_T = {
    'AUS':   [(288000,.0020),(324000,.0050),(360000,.0120),(396000,.0140),(432000,.0180),(468000,.0200)],
    'CAN':   [(364800,.0020),(410400,.0050),(456000,.0140),(501600,.0180),(547200,.0200),(592800,.0220)],
    'UK':    [(300000,.0020),(337500,.0050),(375000,.0140),(412500,.0180),(450000,.0200),(487500,.0220)],
    'US TX': [(391200,.0020),(440100,.0050),(489000,.0140),(537900,.0180),(586800,.0200),(635700,.0220)],
    'US NY': [(391200,.0020),(440100,.0050),(489000,.0140),(537900,.0180),(586800,.0200),(635700,.0220)],
}
_HARRY_T = [(288000,.0020),(324000,.0050),(360000,.0120),(396000,.0140),(432000,.0180),(468000,.0200)]
_CAM_T = {
    'UK':    [(360000,.0020),(405000,.0040),(450000,.0170),(495000,.0190),(540000,.0210),(585000,.0220)],
    'CAN':   [(500000,.0100),(562500,.0125),(625000,.0150),(687500,.0175),(750000,.0200),(812500,.0225)],
    'US NY': [(500000,.0100),(562500,.0125),(625000,.0150),(687500,.0175),(750000,.0200),(812500,.0225)],
    'US TX': [(500000,.0150),(562500,.0175),(625000,.0200),(687500,.0225),(750000,.0250),(812500,.0275)],
    'AUS':   [(500000,.0075),(562500,.0100),(625000,.0125),(687500,.0150),(750000,.0175),(812500,.0200)],
}

def get_tier_rate(name, level, region, attainment):
    if level in ('MANAGER', 'SPECIAL'):
        return 0.0, None
    if level == 'AE1':
        tiers = _AE1_T.get(region, [])
    elif level == 'AE2':
        tiers = _HARRY_T if name == 'Harry Steele' else _AE2_T.get(region, [])
    elif level == 'CAM':
        tiers = _CAM_T.get(region, [])
    else:
        return 0.0, None
    rate, label = 0.0, None
    for threshold, r in tiers:
        if attainment >= threshold:
            rate, label = r, str(threshold)
        else:
            break
    return rate, label

def _get_tiers_for(name, level, region):
    if level == 'AE1':   return _AE1_T.get(region, [])
    if level == 'AE2':   return _HARRY_T if name == 'Harry Steele' else _AE2_T.get(region, [])
    if level == 'CAM':   return _CAM_T.get(region, [])
    return []

DEFAULT_FX  = {'USD': 1.0, 'CAD': 0.74, 'AUD': 0.65, 'GBP': 1.26, 'EUR': 1.08}
CCY_SYM     = {'USD': '$', 'CAD': 'C$', 'AUD': 'A$', 'GBP': '£', 'EUR': '€'}
WELCOMES    = ['Good to see you, Teya.', 'Welcome back, Teya.',
               "Hey Teya, let's get into it.", 'Ready when you are, Teya.']


# ================================================================
# section 2: database
# ================================================================

@st.cache_resource
def _engine():
    url = st.secrets['database']['url']
    return create_engine(url, pool_pre_ping=True, pool_size=5, max_overflow=10, pool_recycle=1800)

def db_read(sql, params=None):
    with _engine().connect() as conn:
        result = conn.execute(text(sql), params or {})
        return pd.DataFrame(result.fetchall(), columns=list(result.keys()))

def db_exec(sql, params=None):
    with _engine().connect() as conn:
        conn.execute(text(sql), params or {})
        conn.commit()

_SCHEMA = [
    '''create table if not exists reps (
        id serial primary key, name text unique not null,
        level text not null, region text not null, currency text not null, active boolean default true
    )''',
    '''create table if not exists deals (
        id serial primary key, hubspot_id text unique not null,
        deal_name text, owner text, deal_currency text, close_date date, close_quarter text,
        sales_channel text, booth_items_revenue numeric, invoice_total numeric,
        paid_total numeric, overdue_total numeric, hubspot_invoice_status text,
        hubspot_paid_date date, invoice_numbers text,
        booth_missing boolean default false, is_refund boolean default false,
        last_imported timestamp default now()
    )''',
    '''create table if not exists invoices (
        id serial primary key, invoice_number text unique not null, deal_id int references deals(id),
        source text, customer_name text, invoice_date date, gross_amount numeric, balance numeric,
        status text, is_credit_note boolean default false, last_imported timestamp default now()
    )''',
    '''create table if not exists payments (
        id serial primary key, invoice_id int references invoices(id),
        payment_date date, amount numeric, source text, match_method text, match_confidence text,
        last_imported timestamp default now(),
        unique(invoice_id, payment_date, amount, source)
    )''',
    '''create table if not exists commission_lines (
        id serial primary key, deal_id int references deals(id), payment_id int,
        period text, payment_date date, cash_landed numeric, booth_payable numeric,
        comm_rate numeric, commission numeric, accel_rate numeric, accelerator numeric,
        payout numeric, close_quarter text
    )''',
    '''create table if not exists attainment (
        id serial primary key, rep_name text, close_quarter text,
        total_booth_items numeric, tier_name text, tier_rate numeric, deal_count int,
        unique(rep_name, close_quarter)
    )''',
    '''create table if not exists import_log (
        id serial primary key, ts timestamp default now(), source text, file_name text,
        records_imported int, records_matched int, records_flagged int, notes text
    )''',
]

@st.cache_resource
def init_schema():
    engine = _engine()
    with engine.connect() as conn:
        for stmt in _SCHEMA:
            conn.execute(text(stmt))
        for name, cfg in REPS.items():
            conn.execute(text('''
                insert into reps (name, level, region, currency)
                values (:n, :l, :r, :c)
                on conflict (name) do update set level=excluded.level, region=excluded.region, currency=excluded.currency
            '''), {'n': name, 'l': cfg['level'], 'r': cfg['region'], 'c': cfg['currency']})
        conn.commit()
    return True

def match_invoices_to_deals():
    engine = _engine()
    matched = 0
    with engine.connect() as conn:
        deals = conn.execute(text(
            "select id, invoice_numbers from deals where invoice_numbers is not null and invoice_numbers != ''"
        )).fetchall()
        for deal in deals:
            for raw in str(deal.invoice_numbers).split(','):
                clean = re.sub(r'\D', '', raw.strip())
                if not clean:
                    continue
                r = conn.execute(text(
                    "update invoices set deal_id=:did "
                    "where regexp_replace(invoice_number,'[^0-9]','','g')=:c "
                    "and (deal_id is null or deal_id=:did)"
                ), {'did': deal.id, 'c': clean})
                matched += r.rowcount
        conn.commit()
    return matched

def log_import(source, fname, imported, matched, flagged, notes=''):
    with _engine().connect() as conn:
        conn.execute(text(
            'insert into import_log (source,file_name,records_imported,records_matched,records_flagged,notes) '
            'values (:s,:f,:i,:m,:fl,:n)'
        ), {'s': source, 'f': fname, 'i': imported, 'm': matched, 'fl': flagged, 'n': notes})
        conn.commit()

def last_import_info(source):
    df = db_read(
        'select ts, records_imported, records_flagged from import_log where source=:s order by ts desc limit 1',
        {'s': source}
    )
    return df.iloc[0] if len(df) else None


# ================================================================
# section 3: importers
# ================================================================

def _find_col(df, kws, excl=None):
    for col in df.columns:
        cl = col.lower()
        if excl and any(e in cl for e in excl):
            continue
        if any(k in cl for k in kws):
            return col
    return None

def _cfloat(val):
    if val is None or str(val).strip().lower() in ('', 'nan', 'none', '(no value)'):
        return None
    try:
        return float(str(val).replace(',', '').replace('$', '').strip())
    except Exception:
        return None

def _quarter(d):
    try:
        dt = pd.to_datetime(d)
        return f'{dt.year} Q{(dt.month-1)//3+1}'
    except Exception:
        return None

def import_hubspot(file):
    df = pd.read_csv(file, dtype=str, encoding='utf-8-sig')
    df = df.replace({'(No value)': None, '': None})
    df.columns = df.columns.str.strip()

    id_col      = _find_col(df, ['record id', 'deal id'])
    owner_col   = _find_col(df, ['deal owner', 'owner'])
    ccy_col     = _find_col(df, ['currency'])
    close_col   = _find_col(df, ['close date'])
    ch_col      = _find_col(df, ['channel', 'sales channel', 'source'], excl=['utm'])
    booth_col   = _find_col(df, ['booth items', 'booth item', 'booths items'])
    itotal_col  = _find_col(df, ['invoice total', 'is_invoice_total'])
    paid_col    = _find_col(df, ['paid total', 'is_paidtotal', 'is_paid_total'])
    over_col    = _find_col(df, ['overdue', 'is_overduetotal'])
    status_col  = _find_col(df, ['invoice status', 'is_invoice_status'])
    pdate_col   = _find_col(df, ['latest invoice paid', 'is_latest_paid_date', 'paid date'])
    inum_col    = _find_col(df, ['invoice number', 'is_invoicenumbers'])
    name_col    = _find_col(df, ['deal name'])

    if not id_col:
        raise ValueError('Cannot find Record ID column in HubSpot file')

    df = df.drop_duplicates(subset=[id_col])
    engine = _engine()
    upserted = 0

    with engine.connect() as conn:
        for _, row in df.iterrows():
            hid = str(row[id_col]).strip() if row[id_col] else None
            if not hid:
                continue

            owner   = clean_owner(row[owner_col]) if owner_col else None
            ccy     = str(row[ccy_col]).strip() if ccy_col and row[ccy_col] else 'USD'
            close_r = row[close_col] if close_col else None
            cdate   = None
            cq      = None
            try:
                cdate = pd.to_datetime(close_r).date() if close_r else None
                cq    = _quarter(close_r) if close_r else None
            except Exception:
                pass

            channel = detect_channel(row[ch_col]) if ch_col else 'inbound'
            booth   = _cfloat(row[booth_col]) if booth_col else None
            itotal  = _cfloat(row[itotal_col]) if itotal_col else None
            paid    = _cfloat(row[paid_col]) if paid_col else None
            over    = _cfloat(row[over_col]) if over_col else None
            if paid is None and itotal is not None and over is not None:
                paid = itotal - over

            status  = str(row[status_col]).strip() if status_col and row[status_col] else None
            dname   = str(row[name_col]).strip() if name_col and row[name_col] else None
            inums   = str(row[inum_col]).strip() if inum_col and row[inum_col] else None

            pdate = None
            if pdate_col and row[pdate_col]:
                try:
                    pdate = pd.to_datetime(row[pdate_col]).date()
                except Exception:
                    pass

            bv = booth or 0.0
            iv = itotal or 0.0
            pv = paid or 0.0
            booth_missing = (bv == 0.0) and (pv > 0 or iv > 0)
            is_refund     = any((v or 0) < 0 for v in [paid, itotal, booth])

            conn.execute(text('''
                insert into deals (
                    hubspot_id, deal_name, owner, deal_currency, close_date, close_quarter,
                    sales_channel, booth_items_revenue, invoice_total, paid_total, overdue_total,
                    hubspot_invoice_status, hubspot_paid_date, invoice_numbers,
                    booth_missing, is_refund, last_imported
                ) values (
                    :hid,:dn,:ow,:cy,:cd,:cq,:ch,:bo,:it,:pa,:ov,:st,:pd,:in,:bm,:ir,now()
                )
                on conflict (hubspot_id) do update set
                    deal_name=excluded.deal_name, owner=excluded.owner,
                    deal_currency=excluded.deal_currency, close_date=excluded.close_date,
                    close_quarter=excluded.close_quarter, sales_channel=excluded.sales_channel,
                    booth_items_revenue=excluded.booth_items_revenue, invoice_total=excluded.invoice_total,
                    paid_total=excluded.paid_total, overdue_total=excluded.overdue_total,
                    hubspot_invoice_status=excluded.hubspot_invoice_status,
                    hubspot_paid_date=excluded.hubspot_paid_date, invoice_numbers=excluded.invoice_numbers,
                    booth_missing=excluded.booth_missing, is_refund=excluded.is_refund, last_imported=now()
            '''), {
                'hid': hid, 'dn': dname, 'ow': owner, 'cy': ccy, 'cd': cdate, 'cq': cq,
                'ch': channel, 'bo': booth, 'it': itotal, 'pa': paid, 'ov': over,
                'st': status, 'pd': pdate, 'in': inums, 'bm': booth_missing, 'ir': is_refund,
            })
            upserted += 1

        conn.commit()

    log_import('hubspot', getattr(file, 'name', 'upload'), upserted, 0, 0)
    return upserted, 0

# ---- quickbooks parser ----

def _qb_float(val):
    if not val or str(val).strip().lower() in ('', '-', 'none', 'nan'):
        return 0.0
    try:
        return float(re.sub(r'[,$\s]', '', str(val)))
    except Exception:
        return 0.0

def _is_date(val):
    try:
        pd.to_datetime(str(val), dayfirst=False)
        return bool(str(val).strip())
    except Exception:
        return False

def _parse_qb_customers(file):
    raw = pd.read_csv(file, header=None, dtype=str, encoding='utf-8-sig').fillna('')
    header_idx = None
    for i, row in raw.iterrows():
        if 'date' in [str(v).strip().lower() for v in row.values]:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError('Cannot find header row in QuickBooks file')

    headers = [str(v).strip() for v in raw.iloc[header_idx].values]
    data    = raw.iloc[header_idx + 1:].reset_index(drop=True)
    col     = {h.lower(): i for i, h in enumerate(headers) if h}

    date_c = col.get('date', 0)
    type_c = next((col[k] for k in ['type', 'transaction type'] if k in col), 1)
    num_c  = next((col[k] for k in ['no.', 'num', 'ref no.', 'number', 'ref'] if k in col), 3)

    num_cols = [i for i, h in enumerate(headers) if any(k in h.lower() for k in ['debit','credit','amount'])]
    if len(num_cols) >= 2:
        credit_c = num_cols[-1]
        debit_c  = num_cols[-2]
    else:
        credit_c = len(headers) - 1
        debit_c  = len(headers) - 2

    customers = {}
    current   = None

    for _, row in data.iterrows():
        vals = [str(row.iloc[i]).strip() if i < len(row) else '' for i in range(len(headers))]
        if not any(vals):
            continue
        date_v = vals[date_c] if date_c < len(vals) else ''
        type_v = vals[type_c] if type_c < len(vals) else ''
        col0   = vals[0]

        if not _is_date(date_v) and not type_v and col0 and not col0.lower().startswith('total'):
            current = col0
            customers.setdefault(current, {'invoices': [], 'payments': []})
            continue

        if current is None or col0.lower().startswith('total'):
            continue

        if type_v.lower() == 'invoice':
            inv_num = vals[num_c] if num_c < len(vals) else ''
            if not inv_num:
                inv_num = next((v for v in vals if re.match(r'^\d{3,}$', v)), '')
            amount = _qb_float(vals[credit_c]) or _qb_float(vals[debit_c])
            if inv_num:
                customers[current]['invoices'].append({'date': date_v, 'inv': inv_num, 'amount': amount})

        elif type_v.lower() == 'payment':
            amount = _qb_float(vals[credit_c])
            if amount == 0:
                amount = abs(_qb_float(vals[debit_c]))
            customers[current]['payments'].append({'date': date_v, 'amount': amount})

    return customers

def import_qb(file, source):
    customers = _parse_qb_customers(file)
    invoices_out, payments_out, flagged = [], [], []

    for customer, data in customers.items():
        invs = [dict(i, remaining=i['amount']) for i in sorted(data['invoices'], key=lambda x: x['date'])]

        for inv in invs:
            if inv['inv']:
                invoices_out.append({
                    'inv_num': inv['inv'], 'customer': customer, 'date': inv['date'],
                    'gross': inv['amount'], 'is_cn': inv['amount'] < 0,
                })

        for pay in data['payments']:
            p_amt, p_date, rem = pay['amount'], pay['date'], pay['amount']

            # exact match on remaining balance
            for inv in invs:
                if inv['remaining'] > 0 and abs(inv['remaining'] - rem) < 0.02:
                    payments_out.append({'inv': inv['inv'], 'date': p_date, 'amount': rem, 'method': 'exact', 'conf': 'high'})
                    inv['remaining'] -= rem
                    rem = 0
                    break
            if rem == 0:
                continue

            # exact match on invoice total
            for inv in invs:
                if inv['remaining'] > 0 and abs(inv['amount'] - rem) < 0.02:
                    payments_out.append({'inv': inv['inv'], 'date': p_date, 'amount': rem, 'method': 'exact', 'conf': 'high'})
                    inv['remaining'] -= rem
                    rem = 0
                    break
            if rem == 0:
                continue

            # fifo split
            splits = []
            for inv in invs:
                if rem <= 0:
                    break
                if inv['remaining'] <= 0:
                    continue
                apply = min(rem, inv['remaining'])
                splits.append({'inv': inv['inv'], 'date': p_date, 'amount': apply, 'method': 'split', 'conf': 'medium'})
                inv['remaining'] -= apply
                rem -= apply
            if splits:
                payments_out.extend(splits)

            if rem > 0.02:
                flagged.append({'customer': customer, 'date': p_date, 'original': p_amt, 'unmatched': rem})
                payments_out.append({'inv': None, 'date': p_date, 'amount': rem, 'method': 'unmatched', 'conf': 'flagged'})

    engine = _engine()
    upserted = linked = 0

    with engine.connect() as conn:
        conn.execute(text('delete from payments where source=:s'), {'s': source})
        for inv in invoices_out:
            conn.execute(text('''
                insert into invoices (invoice_number,source,customer_name,invoice_date,gross_amount,balance,status,is_credit_note)
                values (:n,:s,:c,:d,:g,:b,'unpaid',:cn)
                on conflict (invoice_number) do update set
                    customer_name=excluded.customer_name, invoice_date=excluded.invoice_date,
                    gross_amount=excluded.gross_amount, balance=excluded.balance, last_imported=now()
            '''), {'n': inv['inv_num'], 's': source, 'c': inv['customer'],
                   'd': inv['date'] or None, 'g': inv['gross'], 'b': inv['gross'], 'cn': inv['is_cn']})
            upserted += 1

        for pay in payments_out:
            if pay['conf'] == 'flagged' or not pay['inv']:
                continue
            inv_row = conn.execute(text('select id from invoices where invoice_number=:n'), {'n': pay['inv']}).fetchone()
            if inv_row:
                conn.execute(text('''
                    insert into payments (invoice_id,payment_date,amount,source,match_method,match_confidence)
                    values (:i,:d,:a,:s,:m,:c)
                    on conflict (invoice_id,payment_date,amount,source) do nothing
                '''), {'i': inv_row.id, 'd': pay['date'] or None, 'a': pay['amount'],
                       's': source, 'm': pay['method'], 'c': pay['conf']})
                linked += 1

        conn.commit()

    log_import(source, getattr(file, 'name', 'upload'), upserted, linked, len(flagged))
    return upserted, linked, flagged

# ---- xero parser ----

def _xfloat(val):
    if not val or str(val).strip().lower() in ('', 'nan', 'none', '-'):
        return 0.0
    try:
        return float(re.sub(r'[,$\s]', '', str(val)))
    except Exception:
        return 0.0

def _xdate(val):
    if not val or str(val).strip().lower() in ('', 'nan', 'none', '-'):
        return None
    try:
        return pd.to_datetime(str(val), dayfirst=True).date()
    except Exception:
        return None

def import_xero(file, source):
    raw = pd.read_csv(file, header=None, dtype=str, encoding='utf-8-sig').fillna('')
    if len(raw) < 6:
        raise ValueError(f'Xero file too short: {source}')

    headers = [str(v).strip() for v in raw.iloc[4].values]
    data    = raw.iloc[5:].reset_index(drop=True)
    col     = {h.lower(): i for i, h in enumerate(headers) if h}

    def gc(keys):
        for k in keys:
            if k in col:
                return col[k]
        return None

    inv_c   = gc(['invoice number', 'invoice no.'])
    con_c   = gc(['contact'])
    idate_c = gc(['invoice date', 'date'])
    pdate_c = gc(['last payment date', 'last paid date'])
    gross_c = gc(['gross (source)', 'gross'])
    bal_c   = gc(['balance (source)', 'balance'])
    stat_c  = gc(['status'])

    engine = _engine()
    upserted = linked = 0

    with engine.connect() as conn:
        conn.execute(text('delete from payments where source=:s'), {'s': source})

        for _, row in data.iterrows():
            vals = [str(row.iloc[i]).strip() if i < len(row) else '' for i in range(len(headers))]
            inv_raw = vals[inv_c] if inv_c is not None else ''
            if not inv_raw or inv_raw.lower() in ('total', '', 'nan'):
                continue

            is_cn   = inv_raw.upper().startswith('CN-')
            status  = vals[stat_c].strip() if stat_c is not None else None
            contact = vals[con_c].strip() if con_c is not None else None
            raw_g   = vals[gross_c] if gross_c is not None else '0'
            gross   = _xfloat(raw_g) * (-1 if is_cn and _xfloat(raw_g) > 0 else 1)
            balance = _xfloat(vals[bal_c]) if bal_c is not None else gross
            idate   = _xdate(vals[idate_c]) if idate_c is not None else None
            pdate   = _xdate(vals[pdate_c]) if pdate_c is not None else None

            conn.execute(text('''
                insert into invoices (invoice_number,source,customer_name,invoice_date,gross_amount,balance,status,is_credit_note)
                values (:n,:s,:c,:d,:g,:b,:st,:cn)
                on conflict (invoice_number) do update set
                    customer_name=excluded.customer_name, invoice_date=excluded.invoice_date,
                    gross_amount=excluded.gross_amount, balance=excluded.balance,
                    status=excluded.status, last_imported=now()
            '''), {'n': inv_raw, 's': source, 'c': contact, 'd': idate,
                   'g': gross, 'b': balance, 'st': status, 'cn': is_cn})
            upserted += 1

            if pdate and gross != 0:
                inv_row = conn.execute(text('select id from invoices where invoice_number=:n'), {'n': inv_raw}).fetchone()
                if inv_row:
                    conn.execute(text('''
                        insert into payments (invoice_id,payment_date,amount,source,match_method,match_confidence)
                        values (:i,:d,:a,:s,'xero_direct','high')
                        on conflict (invoice_id,payment_date,amount,source) do nothing
                    '''), {'i': inv_row.id, 'd': pdate, 'a': gross, 's': source})
                    linked += 1

        conn.commit()

    log_import(source, getattr(file, 'name', 'upload'), upserted, linked, 0)
    return upserted, linked


# ================================================================
# section 4: calculation engine
# ================================================================

def calculate_all():
    engine = _engine()
    with engine.connect() as conn:
        conn.execute(text('delete from commission_lines'))
        conn.execute(text('delete from attainment'))
        conn.commit()

        # attainment: per rep per close quarter
        rows = conn.execute(text('''
            select d.owner, d.close_quarter, d.booth_items_revenue, r.level, r.region
            from deals d join reps r on r.name=d.owner
            where d.booth_items_revenue is not null and d.booth_items_revenue!=0
              and d.close_quarter is not null
        ''')).fetchall()

        groups = {}
        for row in rows:
            key = (row.owner, row.close_quarter, row.level, row.region)
            groups.setdefault(key, []).append(float(row.booth_items_revenue or 0))

        for (owner, quarter, level, region), vals in groups.items():
            total = sum(vals)
            rate, label = get_tier_rate(owner, level, region, total)
            conn.execute(text('''
                insert into attainment (rep_name,close_quarter,total_booth_items,tier_name,tier_rate,deal_count)
                values (:n,:q,:t,:tn,:r,:c)
                on conflict (rep_name,close_quarter) do update set
                    total_booth_items=excluded.total_booth_items, tier_name=excluded.tier_name,
                    tier_rate=excluded.tier_rate, deal_count=excluded.deal_count
            '''), {'n': owner, 'q': quarter, 't': total, 'tn': label, 'r': rate, 'c': len(vals)})

        conn.commit()

        # commission lines from payment records
        pay_rows = conn.execute(text('''
            select p.id as pid, p.payment_date, p.amount, d.id as did,
                   d.owner, d.close_quarter, d.booth_items_revenue, d.invoice_total,
                   d.sales_channel, r.level, r.region
            from payments p join invoices i on i.id=p.invoice_id
            join deals d on d.id=i.deal_id join reps r on r.name=d.owner
            where p.payment_date is not null and d.invoice_total is not null and d.invoice_total!=0
        ''')).fetchall()

        for row in pay_rows:
            _insert_line(conn, row.did, row.owner, row.close_quarter, float(row.booth_items_revenue or 0),
                         float(row.invoice_total), float(row.amount), row.payment_date,
                         row.sales_channel, row.level, row.region, int(row.pid))

        # hubspot fallback: paid deals with no payment records
        hs_rows = conn.execute(text('''
            select d.id as did, d.owner, d.close_quarter, d.booth_items_revenue,
                   d.invoice_total, d.paid_total, d.hubspot_paid_date, d.sales_channel, r.level, r.region
            from deals d join reps r on r.name=d.owner
            where d.hubspot_paid_date is not null
              and d.paid_total is not null and d.paid_total!=0
              and d.invoice_total is not null and d.invoice_total!=0
              and not exists (
                  select 1 from invoices i join payments p on p.invoice_id=i.id where i.deal_id=d.id
              )
        ''')).fetchall()

        for row in hs_rows:
            _insert_line(conn, row.did, row.owner, row.close_quarter, float(row.booth_items_revenue or 0),
                         float(row.invoice_total), float(row.paid_total or 0), row.hubspot_paid_date,
                         row.sales_channel, row.level, row.region, None)

        conn.commit()

    return len(pay_rows), len(hs_rows)

def _insert_line(conn, deal_id, owner, close_quarter, booth, invoice_total, amount,
                 payment_date, sales_channel, level, region, payment_id):
    if invoice_total == 0:
        return
    proportion    = max(-2.0, min(2.0, amount / invoice_total))
    booth_payable = booth * proportion
    channel       = detect_channel(sales_channel)
    comm_rate     = get_comm_rate(owner, level, region, channel)
    att = conn.execute(text(
        'select tier_rate from attainment where rep_name=:n and close_quarter=:q'
    ), {'n': owner, 'q': close_quarter}).fetchone()
    accel_rate  = float(att.tier_rate) if att else 0.0
    commission  = booth_payable * comm_rate
    accelerator = booth_payable * accel_rate
    payout      = commission + accelerator
    period      = pd.to_datetime(payment_date).strftime('%Y-%m')
    conn.execute(text('''
        insert into commission_lines
            (deal_id,payment_id,period,payment_date,cash_landed,booth_payable,
             comm_rate,commission,accel_rate,accelerator,payout,close_quarter)
        values (:did,:pid,:per,:pd,:cl,:bp,:cr,:co,:ar,:ac,:po,:cq)
    '''), {
        'did': deal_id, 'pid': payment_id, 'per': period, 'pd': str(payment_date),
        'cl': amount, 'bp': booth_payable, 'cr': comm_rate, 'co': commission,
        'ar': accel_rate, 'ac': accelerator, 'po': payout, 'cq': close_quarter,
    })


# ================================================================
# section 5: shared helpers
# ================================================================

def _period_clause(period):
    if not period:
        return '', {}
    return 'and cl.payment_date between :p0 and :p1', {'p0': period[0], 'p1': period[1]}

def _fx_deals(fx):
    cases = ' '.join(f"when deal_currency='{k}' then {v}" for k, v in fx.items())
    return f'(case {cases} else 1.0 end)'

def _fx_cl(fx):
    cases = ' '.join(f"when d.deal_currency='{k}' then {v}" for k, v in fx.items())
    return f'(case {cases} else 1.0 end)'

def _scalar(sql, params):
    try:
        return float(db_read(sql, params).iloc[0, 0] or 0)
    except Exception:
        return 0.0

def _fmt(v, sym='$'):
    return f'{sym}{v:,.2f}'


# ================================================================
# section 6: page renderers
# ================================================================

# ---- import ----

def page_import():
    st.title('Import')
    st.caption('Upload source files, then run Calculate. Each re-import is fully idempotent.')

    # hubspot
    st.subheader('HubSpot Deals')
    _import_status('hubspot')
    f = st.file_uploader('HubSpot Deals CSV', type='csv', key='hs')
    if f and st.button('Import HubSpot', key='hs_btn'):
        with st.spinner('Importing...'):
            try:
                n, _ = import_hubspot(f)
                m = match_invoices_to_deals()
                st.success(f'{n} deals upserted -- {m} invoices matched to deals.')
            except Exception as e:
                st.error(str(e))

    st.divider()

    # quickbooks
    st.subheader('QuickBooks')
    col_us, col_ca = st.columns(2)

    with col_us:
        st.markdown('**US -- Inbox Booths LLC**')
        _import_status('QB_US')
        f = st.file_uploader('QB US CSV', type='csv', key='qb_us')
        if f and st.button('Import QB US', key='qb_us_btn'):
            with st.spinner('Importing...'):
                try:
                    n, l, fl = import_qb(f, 'QB_US')
                    match_invoices_to_deals()
                    st.success(f'{n} invoices, {l} payments linked, {len(fl)} flagged.')
                except Exception as e:
                    st.error(str(e))

    with col_ca:
        st.markdown('**Canada -- Inbox Design Inc.**')
        _import_status('QB_CA')
        f = st.file_uploader('QB Canada CSV', type='csv', key='qb_ca')
        if f and st.button('Import QB Canada', key='qb_ca_btn'):
            with st.spinner('Importing...'):
                try:
                    n, l, fl = import_qb(f, 'QB_CA')
                    match_invoices_to_deals()
                    st.success(f'{n} invoices, {l} payments linked, {len(fl)} flagged.')
                except Exception as e:
                    st.error(str(e))

    st.divider()

    # xero
    st.subheader('Xero')
    col_uk, col_au = st.columns(2)

    with col_uk:
        st.markdown('**UK -- Bureau Booths UK Limited**')
        _import_status('XERO_UK')
        f = st.file_uploader('Xero UK CSV', type='csv', key='xero_uk')
        if f and st.button('Import Xero UK', key='xero_uk_btn'):
            with st.spinner('Importing...'):
                try:
                    n, l = import_xero(f, 'XERO_UK')
                    match_invoices_to_deals()
                    st.success(f'{n} invoices, {l} payments linked.')
                except Exception as e:
                    st.error(str(e))

    with col_au:
        st.markdown('**AUS -- Bureau Booths Pty Limited**')
        st.caption('Ensure the export includes Last Payment Date column.')
        _import_status('XERO_AU')
        f = st.file_uploader('Xero AUS CSV', type='csv', key='xero_au')
        if f and st.button('Import Xero AUS', key='xero_au_btn'):
            with st.spinner('Importing...'):
                try:
                    n, l = import_xero(f, 'XERO_AU')
                    match_invoices_to_deals()
                    st.success(f'{n} invoices, {l} payments linked.')
                except Exception as e:
                    st.error(str(e))

    st.divider()

    # calculate
    st.subheader('Calculate Commission')
    st.caption('Recalculates everything from scratch. Run after any import.')
    if st.button('Calculate Commission', type='primary', key='calc'):
        with st.spinner('Calculating...'):
            try:
                np, nh = calculate_all()
                st.success(f'Done -- {np} lines from payment records, {nh} from HubSpot fallback.')
            except Exception as e:
                st.error(str(e))

    st.divider()

    # flagged payments
    st.subheader('Flagged Payments')
    df = db_read('''
        select p.id, p.source, p.payment_date, p.amount, i.customer_name
        from payments p left join invoices i on i.id=p.invoice_id
        where p.match_confidence='flagged' order by p.payment_date desc
    ''')
    if df.empty:
        st.success('No flagged payments.')
    else:
        st.warning(f'{len(df)} payments could not be matched to invoices.')
        st.dataframe(df, hide_index=True, use_container_width=True)

    st.divider()

    # import log
    st.subheader('Import Log')
    log = db_read('select ts, source, file_name, records_imported, records_matched, records_flagged from import_log order by ts desc limit 50')
    if not log.empty:
        st.dataframe(log, hide_index=True, use_container_width=True)

def _import_status(source):
    row = last_import_info(source)
    if row is not None:
        ts = str(row['ts'])[:16]
        st.caption(f'Last import: {ts} UTC -- {row["records_imported"]} records')
    else:
        st.caption('Not yet imported')


# ---- dashboard ----

def page_dashboard():
    st.title('Dashboard')
    fx     = st.session_state.get('fx', DEFAULT_FX)
    period = st.session_state.get('period_filter')
    pc, pp = _period_clause(period)
    fxd    = _fx_deals(fx)
    fxc    = _fx_cl(fx)

    invoiced    = _scalar(f'select coalesce(sum(invoice_total*{fxd}),0) from deals', {})
    outstanding = _scalar(f'select coalesce(sum(overdue_total*{fxd}),0) from deals', {})
    collected   = _scalar(
        f'select coalesce(sum(cl.cash_landed*{fxc}),0) from commission_lines cl join deals d on d.id=cl.deal_id where 1=1 {pc}', pp)
    payout = _scalar(
        f'select coalesce(sum(cl.payout*{fxc}),0) from commission_lines cl join deals d on d.id=cl.deal_id where 1=1 {pc}', pp)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric('Total Invoiced (USD)',  _fmt(invoiced))
    c2.metric('Cash Collected (USD)',  _fmt(collected))
    c3.metric('Outstanding (USD)',     _fmt(outstanding))
    c4.metric('Total Payout (USD)',    _fmt(payout))

    if invoiced > 0:
        st.progress(min(collected / invoiced, 1.0), text=f'Collection Rate: {min(collected/invoiced,1.0):.1%}')

    st.divider()
    cl, cr = st.columns(2)

    with cl:
        st.subheader('Revenue by Region')
        df = db_read(
            f'select r.region, sum(cl.cash_landed*{fxc}) as usd '
            f'from commission_lines cl join deals d on d.id=cl.deal_id join reps r on r.name=d.owner '
            f'where 1=1 {pc} group by r.region order by usd desc', pp)
        if not df.empty:
            fig = px.bar(df, x='region', y='usd', labels={'region':'Region','usd':'USD'},
                         color='region', color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_layout(showlegend=False, margin=dict(t=20,b=20))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info('No data for selected period.')

    with cr:
        st.subheader('Revenue by Channel')
        df = db_read(
            f'select d.sales_channel, sum(cl.cash_landed*{fxc}) as usd '
            f'from commission_lines cl join deals d on d.id=cl.deal_id where 1=1 {pc} '
            f'group by d.sales_channel', pp)
        if not df.empty:
            fig = px.pie(df, names='sales_channel', values='usd',
                         color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_layout(margin=dict(t=20,b=20))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info('No data for selected period.')

    st.divider()
    st.subheader('Top Reps by Payout')
    df = db_read(
        f'select d.owner as rep, r.region, '
        f'sum(cl.commission*{fxc}) as comm_usd, sum(cl.accelerator*{fxc}) as acc_usd, '
        f'sum(cl.payout*{fxc}) as payout_usd '
        f'from commission_lines cl join deals d on d.id=cl.deal_id join reps r on r.name=d.owner '
        f'where 1=1 {pc} group by d.owner,r.region order by payout_usd desc limit 20', pp)

    if not df.empty:
        cc, ct = st.columns([2, 1])
        with cc:
            fig = px.bar(df, x='rep', y='payout_usd', labels={'rep':'','payout_usd':'Payout (USD)'},
                         color='region', color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_layout(margin=dict(t=20,b=60), xaxis_tickangle=-30)
            st.plotly_chart(fig, use_container_width=True)
        with ct:
            disp = df[['rep','region','payout_usd']].copy()
            disp['payout_usd'] = disp['payout_usd'].apply(lambda x: f'${x:,.2f}')
            st.dataframe(disp, hide_index=True, use_container_width=True)
    else:
        st.info('No commission data. Import files and run Calculate.')


# ---- monthly payout ----

def page_monthly_payout():
    st.title('Monthly Payout')
    period = st.session_state.get('period_filter')
    pc, pp = _period_clause(period)

    df = db_read(
        f'select d.owner, r.region, r.currency, r.level, '
        f'sum(cl.commission) as commission, sum(cl.accelerator) as accelerator, sum(cl.payout) as payout '
        f'from commission_lines cl join deals d on d.id=cl.deal_id join reps r on r.name=d.owner '
        f'where 1=1 {pc} group by d.owner,r.region,r.currency,r.level order by r.region,d.owner', pp)

    if df.empty:
        st.info('No commission data for the selected period.')
        return

    missing = _scalar('select count(*) from deals where booth_missing=true', {})
    if missing:
        st.warning(f'{int(missing)} deals are missing Booth Items revenue and are excluded from commission.')

    regions = df['region'].unique().tolist()
    tabs = st.tabs(regions + ['All Reps'])

    for i, region in enumerate(regions):
        with tabs[i]:
            sub = df[df['region'] == region].copy()
            sym = CCY_SYM.get(sub.iloc[0]['currency'], '$')
            disp = sub[['owner', 'level', 'currency', 'commission', 'accelerator', 'payout']].copy()
            disp.columns = ['Rep', 'Level', 'Currency', 'Commission', 'Accelerator', 'Total Payout']
            for c in ['Commission', 'Accelerator', 'Total Payout']:
                disp[c] = disp[c].apply(lambda x: f'{sym}{x:,.2f}')
            total_row = pd.DataFrame([{'Rep': 'TOTAL', 'Level': '', 'Currency': '',
                'Commission': '', 'Accelerator': '', 'Total Payout': f'{sym}{sub["payout"].sum():,.2f}'}])
            st.dataframe(pd.concat([disp, total_row], ignore_index=True), hide_index=True, use_container_width=True)

    with tabs[-1]:
        disp = df[['owner','region','level','currency','commission','accelerator','payout']].copy()
        disp.columns = ['Rep','Region','Level','Currency','Commission','Accelerator','Total Payout']
        st.dataframe(disp, hide_index=True, use_container_width=True)

    st.divider()
    st.subheader('Rep Drill-Down')
    selected = st.selectbox('Select Rep', df['owner'].tolist())

    if selected:
        lines = db_read(
            f'select cl.payment_date, cl.period, cl.close_quarter, d.deal_name, d.sales_channel, '
            f'cl.cash_landed, cl.booth_payable, cl.comm_rate, cl.commission, '
            f'cl.accel_rate, cl.accelerator, cl.payout '
            f'from commission_lines cl join deals d on d.id=cl.deal_id '
            f'where d.owner=:rep {pc} order by cl.payment_date desc',
            {'rep': selected, **pp})

        if lines.empty:
            st.info('No lines for this rep in the selected period.')
        else:
            ccy = REPS.get(selected, {}).get('currency', 'USD')
            sym = CCY_SYM.get(ccy, '$')
            disp = lines.copy()
            for c in ['cash_landed','booth_payable','commission','accelerator','payout']:
                disp[c] = disp[c].apply(lambda x: f'{sym}{x:,.2f}')
            for c in ['comm_rate','accel_rate']:
                disp[c] = disp[c].apply(lambda x: f'{float(x):.2%}')
            st.dataframe(disp, hide_index=True, use_container_width=True)

            xl = _build_excel(selected, lines, ccy)
            st.download_button(
                f'Download {selected} Excel', data=xl,
                file_name=f'{selected.replace(" ","_")}_commission.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    st.divider()
    if st.button('Download All Reps (ZIP)'):
        zb = _build_zip(df, pp, pc)
        st.download_button('Save ZIP', data=zb, file_name='all_reps_commission.zip', mime='application/zip')

def _build_excel(rep, lines, currency):
    sym = CCY_SYM.get(currency, '$')
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        wb   = writer.book
        hdr  = wb.add_format({'bold': True, 'bg_color': '#1a1a2e', 'font_color': 'white', 'border': 1})
        mon  = wb.add_format({'num_format': f'"{sym}"#,##0.00'})
        pct  = wb.add_format({'num_format': '0.00%'})
        bold = wb.add_format({'bold': True})

        summary = lines.groupby('period').agg(
            commission=('commission','sum'), accelerator=('accelerator','sum'), payout=('payout','sum')
        ).reset_index()
        summary.columns = ['Period','Commission','Accelerator','Total Payout']
        summary.to_excel(writer, sheet_name='Summary', index=False, startrow=1)
        ws = writer.sheets['Summary']
        ws.write(0, 0, rep, bold)
        for j, c in enumerate(summary.columns):
            ws.write(1, j, c, hdr)
        ws.set_column('A:A', 10)
        ws.set_column('B:D', 16, mon)

        detail = lines[['payment_date','deal_name','sales_channel','close_quarter',
                         'cash_landed','booth_payable','comm_rate','commission',
                         'accel_rate','accelerator','payout']].copy()
        detail.columns = ['Payment Date','Deal','Channel','Close Quarter',
                          'Cash Landed','Booth Payable','Comm Rate','Commission',
                          'Accel Rate','Accelerator','Total Payout']
        detail.to_excel(writer, sheet_name='Deal Detail', index=False)
        ws2 = writer.sheets['Deal Detail']
        for j, c in enumerate(detail.columns):
            ws2.write(0, j, c, hdr)
            if 'Rate' in c:
                ws2.set_column(j, j, 12, pct)
            elif c in ('Cash Landed','Booth Payable','Commission','Accelerator','Total Payout'):
                ws2.set_column(j, j, 16, mon)
            else:
                ws2.set_column(j, j, 20)
    buf.seek(0)
    return buf.read()

def _build_zip(df, pp, pc):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for _, row in df.iterrows():
            rep = row['owner']
            ccy = row['currency']
            lines = db_read(
                f'select cl.payment_date, cl.period, cl.close_quarter, d.deal_name, '
                f'd.sales_channel, cl.cash_landed, cl.booth_payable, cl.comm_rate, cl.commission, '
                f'cl.accel_rate, cl.accelerator, cl.payout '
                f'from commission_lines cl join deals d on d.id=cl.deal_id '
                f'where d.owner=:rep {pc} order by cl.payment_date desc',
                {'rep': rep, **pp})
            if lines.empty:
                continue
            zf.writestr(f'{rep.replace(" ","_")}_commission.xlsx', _build_excel(rep, lines, ccy))
    buf.seek(0)
    return buf.read()


# ---- payout history ----

def page_payout_history():
    st.title('Payout History')
    st.caption('Full history -- payment period filter not applied.')

    region_filter = st.selectbox('Region', ['All'] + sorted(set(r['region'] for r in REPS.values())))

    df = db_read('''
        select d.owner, r.region, r.currency, cl.period, cl.close_quarter,
               cl.commission, cl.accelerator, cl.payout
        from commission_lines cl join deals d on d.id=cl.deal_id join reps r on r.name=d.owner
        order by cl.period
    ''')

    if df.empty:
        st.info('No commission history yet.')
        return

    if region_filter != 'All':
        df = df[df['region'] == region_filter]

    if df.empty:
        st.info('No data for selected region.')
        return

    st.subheader('Monthly Commission')
    _history_grid(df, 'period', 'commission')

    st.subheader('Monthly Accelerator')
    _history_grid(df, 'period', 'accelerator')

    st.subheader('Monthly Total Payout')
    _history_grid(df, 'period', 'payout')

    st.subheader('Quarterly Total Payout')
    _history_grid(df, 'close_quarter', 'payout')

def _history_grid(df, time_col, val_col):
    pivot = df.pivot_table(index=['owner','region','currency'], columns=time_col,
                           values=val_col, aggfunc='sum', fill_value=0).reset_index()
    time_cols = sorted([c for c in pivot.columns if c not in ('owner','region','currency')])
    pivot.columns.name = None

    def fmt_row(row):
        sym = CCY_SYM.get(row['currency'], '$')
        for c in time_cols:
            row[c] = f'{sym}{row[c]:,.2f}'
        return row

    pivot = pivot.apply(fmt_row, axis=1)
    pivot = pivot.rename(columns={'owner': 'Rep', 'region': 'Region', 'currency': 'Currency'})
    st.dataframe(pivot, hide_index=True, use_container_width=True)


# ---- quarterly review ----

def page_quarterly_review():
    st.title('Quarterly Review')
    st.caption('Based on close quarter. Period filter not applied.')

    quarters = db_read("select distinct close_quarter from attainment where close_quarter is not null order by close_quarter desc")
    if quarters.empty:
        st.info('No attainment data. Import deals and run Calculate.')
        return

    quarter = st.selectbox('Quarter', quarters['close_quarter'].tolist())

    df = db_read('''
        select a.rep_name, r.level, r.region, r.currency,
               a.total_booth_items, a.tier_name, a.tier_rate, a.deal_count,
               coalesce((select sum(cl.payout) from commission_lines cl
                         join deals d on d.id=cl.deal_id
                         where d.owner=a.rep_name and cl.close_quarter=a.close_quarter), 0) as payout
        from attainment a join reps r on r.name=a.rep_name
        where a.close_quarter=:q order by r.region, a.total_booth_items desc
    ''', {'q': quarter})

    if df.empty:
        st.info('No data for this quarter.')
        return

    disp = df.copy()
    disp['total_booth_items'] = disp.apply(lambda r: f'{CCY_SYM.get(r["currency"],"$")}{r["total_booth_items"]:,.2f}', axis=1)
    disp['tier_rate'] = disp['tier_rate'].apply(lambda x: f'{float(x):.2%}' if x else '0.00%')
    disp['payout']    = disp.apply(lambda r: f'{CCY_SYM.get(r["currency"],"$")}{float(r["payout"]):,.2f}', axis=1)
    disp.columns = ['Rep','Level','Region','Currency','Booth Items','Tier Threshold','Tier Rate','Deals','Payout']
    st.dataframe(disp, hide_index=True, use_container_width=True)

    st.divider()
    st.subheader('Attainment Gauges')
    cols = st.columns(3)
    for i, (_, row) in enumerate(df.iterrows()):
        with cols[i % 3]:
            _gauge(row)

def _gauge(row):
    name    = row['rep_name']
    level   = row['level']
    region  = row['region']
    total   = float(row['total_booth_items'] or 0)
    sym     = CCY_SYM.get(row['currency'], '$')
    tiers   = _get_tiers_for(name, level, region)
    if not tiers:
        return
    max_v   = tiers[-1][0] * 1.3
    colours = ['#f0f4f8','#d0e4f7','#a8c8f0','#6fa8dc','#4285c3','#1a5fa0']
    steps, prev = [], 0
    for j, (th, _) in enumerate(tiers):
        steps.append({'range': [prev, th], 'color': colours[min(j, len(colours)-1)]})
        prev = th
    steps.append({'range': [prev, max_v], 'color': '#0d3b75'})
    fig = go.Figure(go.Indicator(
        mode='gauge+number', value=total,
        title={'text': name, 'font': {'size': 13}},
        number={'prefix': sym, 'valueformat': ',.0f'},
        gauge={'axis': {'range': [0, max_v]}, 'bar': {'color': '#e8622a'},
               'steps': steps,
               'threshold': {'line': {'color': 'red', 'width': 2}, 'thickness': .75,
                             'value': tiers[2][0] if len(tiers) > 2 else total}},
    ))
    fig.update_layout(height=220, margin=dict(t=40,b=10,l=20,r=20))
    st.plotly_chart(fig, use_container_width=True)


# ---- data quality ----

def page_data_quality():
    st.title('Data Quality')
    st.caption('Full unfiltered view.')

    st.subheader('Missing Booth Items Revenue')
    df = db_read('''
        select d.owner, r.region, d.deal_currency, d.deal_name, d.close_date,
               d.invoice_total, d.paid_total, d.hubspot_invoice_status
        from deals d join reps r on r.name=d.owner
        where d.booth_missing=true order by r.region, d.owner, d.close_date desc
    ''')
    if df.empty:
        st.success('No missing booth items.')
    else:
        st.warning(f'{len(df)} deals affected.')
        cc, cr = st.columns(2)
        with cc:
            by_ccy = df.groupby('deal_currency').agg(deals=('deal_name','count'), total=('invoice_total','sum')).reset_index()
            by_ccy['total'] = by_ccy.apply(lambda r: f'{CCY_SYM.get(r["deal_currency"],"$")}{r["total"]:,.2f}', axis=1)
            by_ccy.columns = ['Currency','Deals','Invoice Total']
            st.dataframe(by_ccy, hide_index=True, use_container_width=True)
        with cr:
            by_rep = df.groupby(['owner','deal_currency']).size().reset_index(name='deals')
            by_rep.columns = ['Rep','Currency','Deals']
            st.dataframe(by_rep, hide_index=True, use_container_width=True)
        st.dataframe(df[['owner','region','deal_name','deal_currency','close_date','invoice_total','paid_total']],
                     hide_index=True, use_container_width=True)
        st.download_button('Download CSV', df.to_csv(index=False).encode(), 'missing_booth.csv', 'text/csv')

    st.divider()

    st.subheader('Paid Deals with No Payment Date')
    df2 = db_read('''
        select d.owner, r.region, d.deal_name, d.deal_currency, d.close_date, d.paid_total
        from deals d join reps r on r.name=d.owner
        where d.hubspot_paid_date is null and d.paid_total>0
          and not exists (select 1 from invoices i join payments p on p.invoice_id=i.id where i.deal_id=d.id)
        order by r.region, d.owner
    ''')
    if df2.empty:
        st.success('All paid deals have payment dates.')
    else:
        st.warning(f'{len(df2)} deals affected.')
        st.dataframe(df2, hide_index=True, use_container_width=True)

    st.divider()

    st.subheader('Accounts Receivable')
    df3 = db_read('''
        select i.invoice_number, i.customer_name, i.source, i.invoice_date,
               i.gross_amount, i.balance, i.status, d.owner, d.deal_name
        from invoices i left join deals d on d.id=i.deal_id
        where i.balance>0 and i.status not in ('Paid','Voided','Deleted') and not i.is_credit_note
        order by i.invoice_date
    ''')
    if df3.empty:
        st.success('No outstanding invoices.')
    else:
        st.info(f'{len(df3)} outstanding invoices.')
        st.dataframe(df3, hide_index=True, use_container_width=True)

    st.divider()

    st.subheader('QB Unmatched Payments')
    df4 = db_read('''
        select p.id, p.source, p.payment_date, p.amount, i.customer_name
        from payments p left join invoices i on i.id=p.invoice_id
        where p.match_confidence='flagged' order by p.payment_date desc
    ''')
    if df4.empty:
        st.success('No unmatched QB payments.')
    else:
        st.warning(f'{len(df4)} unmatched.')
        st.dataframe(df4, hide_index=True, use_container_width=True)


# ================================================================
# section 7: main app
# ================================================================

init_schema()

if 'welcomed' not in st.session_state:
    ph = st.empty()
    ph.info(random.choice(WELCOMES))
    time.sleep(3)
    ph.empty()
    st.session_state.welcomed = True

with st.sidebar:
    st.markdown('## Bureau Booths')
    st.caption('Commission Calculator')
    st.divider()

    page = st.radio('', ['Dashboard','Monthly Payout','Payout History','Quarterly Review','Data Quality','Import'],
                    label_visibility='collapsed')

    st.divider()
    st.markdown('**Payment Period**')
    mode = st.selectbox('', ['All Dates','Month','Quarter','Date Range'], label_visibility='collapsed')

    period_filter = None
    if mode == 'Month':
        months = pd.date_range('2023-01-01', date.today(), freq='MS').strftime('%Y-%m').tolist()[::-1]
        m = st.selectbox('', months, label_visibility='collapsed', key='m_sel')
        if m:
            period_filter = (f'{m}-01', f'{m}-31')
    elif mode == 'Quarter':
        qs = [f'{y} Q{q}' for y in range(2023, date.today().year + 2) for q in range(1, 5)
              if f'{y} Q{q}' <= f'{date.today().year} Q{(date.today().month-1)//3+1}'][::-1]
        q = st.selectbox('', qs, label_visibility='collapsed', key='q_sel')
        if q:
            yr, qn = q.split(' Q')
            s = {'1':'01-01','2':'04-01','3':'07-01','4':'10-01'}
            e = {'1':'03-31','2':'06-30','3':'09-30','4':'12-31'}
            period_filter = (f'{yr}-{s[qn]}', f'{yr}-{e[qn]}')
    elif mode == 'Date Range':
        d1 = st.date_input('From', date.today() - timedelta(days=30), label_visibility='collapsed')
        d2 = st.date_input('To',   date.today(), label_visibility='collapsed')
        period_filter = (str(d1), str(d2))

    st.session_state.period_filter = period_filter

    st.divider()
    with st.expander('Exchange Rates (USD base)'):
        fx = {'USD': 1.0}
        for ccy, default in DEFAULT_FX.items():
            if ccy == 'USD':
                continue
            fx[ccy] = st.number_input(ccy, value=default, step=0.01, format='%.4f', key=f'fx_{ccy}')
        st.session_state.fx = fx

match page:
    case 'Dashboard':        page_dashboard()
    case 'Monthly Payout':   page_monthly_payout()
    case 'Payout History':   page_payout_history()
    case 'Quarterly Review': page_quarterly_review()
    case 'Data Quality':     page_data_quality()
    case 'Import':           page_import()
