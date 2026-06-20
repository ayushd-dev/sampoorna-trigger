"""
sampoorna_portfolio_csv.py
---------------------------
Runs daily via GitHub Actions cron.
Fetches today's Sampoorna portfolio snapshot, generates Excel,
and uploads it to the 'sampoorna-portfolio-latest' GitHub Release
so it's available as a direct download link.
"""
import io, os, requests, json
from datetime import datetime, timezone, timedelta
import pandas as pd

BASE_URL     = 'https://superset.bkosh.com'
LMS_DB_ID  = 46
IST        = timezone(timedelta(hours=5, minutes=30))
TODAY      = datetime.now(IST).strftime('%Y-%m-%d')

GITHUB_TOKEN = os.environ.get('GITHUB_TOKEN', '')
REPO         = 'ayushd-dev/sampoorna-trigger'   # public repo — no login required to download
RELEASE_TAG  = 'sampoorna-portfolio-latest'
FILENAME     = f'Sampoorna_Portfolio_{TODAY}.xlsx'

# ── Auth Superset ─────────────────────────────────────────────────────────────
s = requests.Session()
s.headers.update({'User-Agent': 'Mozilla/5.0'})
token = s.post(f'{BASE_URL}/api/v1/security/login',
    json={'username': os.environ['SUPERSET_UN'], 'password': os.environ['SUPERSET_PASS'],
          'provider': 'db', 'refresh': True}, timeout=60).json()['access_token']
s.headers.update({'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'})
csrf = s.get(f'{BASE_URL}/api/v1/security/csrf_token/', timeout=30).json()['result']
s.headers.update({'X-CSRFToken': csrf, 'Referer': f'{BASE_URL}/superset/sqllab/'})
print(f'[{datetime.now(IST):%H:%M:%S}] Logged in to Superset.')


def run_sql(db_id, sql, limit=10000):
    r = s.post(f'{BASE_URL}/api/v1/sqllab/execute/',
        json={'database_id': db_id, 'sql': sql, 'schema': 'sampoorna',
              'runAsync': False, 'queryLimit': limit}, timeout=120)
    r.raise_for_status()
    d = r.json()
    if 'error' in d:
        raise RuntimeError(f'SQL error: {d["error"][:400]}')
    return d.get('data', [])


def parse_agent(desc):
    if desc and '_uid_' in desc:
        return desc[:desc.index('_uid_')]
    return desc or ''


# ── Fetch data ────────────────────────────────────────────────────────────────
print(f'[{datetime.now(IST):%H:%M:%S}] Fetching loan portfolio...')
loan_rows = run_sql(LMS_DB_ID, """
    SELECT
        ll.id                                                            AS loan_id,
        ll.share                                                         AS loan_amount,
        rd.borrower_name,
        rd.borrower_mobile,
        rt.name                                                          AS branch_name,
        lg.description                                                   AS agent_raw,
        lg.status                                                        AS loan_status,
        SUM(li.amount)                                                   AS total_demand,
        SUM(CASE WHEN li.status = 'CLOSED' THEN li.amount ELSE 0 END)   AS total_collected,
        SUM(CASE WHEN li.status = 'OPENED' THEN li.amount ELSE 0 END)   AS outstanding,
        MAX(CASE WHEN li.status = 'OPENED' AND li.due_date < CURRENT_DATE
                 THEN (CURRENT_DATE - li.due_date) ELSE 0 END)          AS dpd
    FROM sampoorna.loan_loan ll
    JOIN sampoorna.loan_installment li
        ON li.repayment_schedule_id = ll.repayment_schedule_id
    JOIN sampoorna.loan_group_loangroup_loans lgl
        ON lgl.loan_id = ll.id
    JOIN sampoorna.loan_group_loangroup lg
        ON lg.id = lgl.loangroup_id
    JOIN sampoorna.recovery_demand rd
        ON rd.loan_id = ll.id
    JOIN sampoorna.recovery_team rt
        ON rt.id = rd.recovery_team_id
    WHERE lg.status IN ('DISBURSED', 'COMPLETE')
    GROUP BY ll.id, ll.share, rd.borrower_name, rd.borrower_mobile,
             rt.name, lg.description, lg.status
""", limit=10000)
print(f'  {len(loan_rows):,} loans')

print(f'[{datetime.now(IST):%H:%M:%S}] Fetching last payment dates...')
lpd_rows = run_sql(LMS_DB_ID, """
    SELECT lpr.loan_id,
           MAX(lpr.created_at::date::text) AS last_payment_date,
           COUNT(*)                        AS payment_count
    FROM sampoorna.loan_payments_loanrepayment lpr
    GROUP BY lpr.loan_id
""", limit=10000)
lpd_map = {r['loan_id']: r for r in lpd_rows}

# ── Build DataFrame ───────────────────────────────────────────────────────────
records = []
for r in loan_rows:
    loan_id    = r['loan_id']
    lpd        = lpd_map.get(loan_id, {})
    demand     = int(float(r['total_demand']     or 0))
    collected  = int(float(r['total_collected']  or 0))
    outstanding= int(float(r['outstanding']      or 0))
    dpd        = int(float(r['dpd']              or 0))
    records.append({
        'Snapshot Date':   TODAY,
        'Loan ID':         loan_id,
        'Borrower Name':   r['borrower_name'],
        'Mobile':          r['borrower_mobile'],
        'Branch':          r['branch_name'],
        'Agent':           parse_agent(r['agent_raw']),
        'Loan Status':     r['loan_status'],
        'Loan Amount':     int(float(r['loan_amount'] or 0)),
        'Total Demand':    demand,
        'Total Collected': collected,
        'Outstanding':     outstanding,
        'Collection %':    round(collected / demand * 100, 1) if demand else 0,
        'DPD':             dpd,
        'DPD Bucket':      ('0' if dpd == 0 else '1-30' if dpd <= 30 else
                            '31-60' if dpd <= 60 else '61-90' if dpd <= 90 else '90+'),
        'Payment Count':   int(lpd.get('payment_count') or 0),
        'Last Payment Date': lpd.get('last_payment_date') or '',
    })

df = pd.DataFrame(records).sort_values(['Branch', 'Agent', 'Borrower Name'])

# ── Write Excel to memory ─────────────────────────────────────────────────────
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='All Loans', index=False)
    ws = writer.sheets['All Loans']
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)
buf.seek(0)
excel_bytes = buf.read()
print(f'[{datetime.now(IST):%H:%M:%S}] Excel generated ({len(excel_bytes):,} bytes).')

# ── Upload to GitHub Release ──────────────────────────────────────────────────
if GITHUB_TOKEN:
    gh = requests.Session()
    gh.headers.update({
        'Authorization': f'token {GITHUB_TOKEN}',
        'Accept': 'application/vnd.github+json',
    })

    # Get or create the release
    r = gh.get(f'https://api.github.com/repos/{REPO}/releases/tags/{RELEASE_TAG}', timeout=30)
    if r.status_code == 200:
        release = r.json()
        release_id = release['id']
        # If today's file already exists (re-run), delete it so we can re-upload
        for asset in release.get('assets', []):
            if asset['name'] == FILENAME:
                gh.delete(f'https://api.github.com/repos/{REPO}/releases/assets/{asset["id"]}', timeout=30)
                print(f'  Replaced existing asset: {asset["name"]}')
        print(f'[{datetime.now(IST):%H:%M:%S}] Using existing release ID={release_id}')
    else:
        r = gh.post(f'https://api.github.com/repos/{REPO}/releases', timeout=30, json={
            'tag_name':         RELEASE_TAG,
            'name':             'Sampoorna Portfolio Snapshot (Latest)',
            'body':             'Auto-updated daily. Download the Excel for today\'s portfolio data.',
            'prerelease':       True,
            'target_commitish': 'main',
        })
        r.raise_for_status()
        release_id = r.json()['id']
        print(f'[{datetime.now(IST):%H:%M:%S}] Created release ID={release_id}')

    # Upload the Excel asset
    up = requests.post(
        f'https://uploads.github.com/repos/{REPO}/releases/{release_id}/assets',
        params={'name': FILENAME},
        headers={
            'Authorization':  f'token {GITHUB_TOKEN}',
            'Content-Type':   'application/octet-stream',
            'Accept':         'application/vnd.github+json',
        },
        data=excel_bytes,
        timeout=60,
    )
    up.raise_for_status()
    download_url = up.json().get('browser_download_url', '')
    print(f'[{datetime.now(IST):%H:%M:%S}] Uploaded to GitHub Release.')
    print(f'  Download URL: {download_url}')

    # ── Update Superset dashboard MARKDOWN node with new link list ────────────
    print(f'[{datetime.now(IST):%H:%M:%S}] Updating Superset dashboard tab...')
    rel_fresh = gh.get(f'https://api.github.com/repos/{REPO}/releases/tags/{RELEASE_TAG}', timeout=30).json()
    dated_assets = sorted(
        [a for a in rel_fresh.get('assets', []) if 'Portfolio_20' in a['name']],
        key=lambda a: a['name'], reverse=True)

    link_rows = ''
    for a in dated_assets:
        date_str = a['name'].replace('Sampoorna_Portfolio_', '').replace('.xlsx', '')
        link_rows += (
            f'<a href="{a["browser_download_url"]}" '
            'style="display:block;margin:6px 0;color:#1890ff;font-size:14px;text-decoration:none;">'
            f'&#11123; {date_str}</a>'
        )
    md_content = (
        '<div style="padding:16px 0;">'
        '<p style="font-size:15px;font-weight:600;color:#333;margin-bottom:4px;">Sampoorna Portfolio Snapshot</p>'
        '<p style="font-size:13px;color:#666;margin-bottom:14px;">'
        'One row per loan &mdash; loan amount, demand, collected, outstanding, DPD. '
        'New file added every morning at 8 AM IST.</p>'
        '<div style="background:#f5f5f5;border-radius:6px;padding:12px 16px;">'
        + link_rows + '</div></div>'
    )
    ss = requests.Session()
    ss.headers.update({'User-Agent': 'Mozilla/5.0'})
    ss_token = ss.post(f'{BASE_URL}/api/v1/security/login',
        json={'username': os.environ['SUPERSET_UN'], 'password': os.environ['SUPERSET_PASS'],
              'provider': 'db', 'refresh': True}, timeout=60).json()['access_token']
    ss.headers.update({'Authorization': f'Bearer {ss_token}', 'Content-Type': 'application/json'})
    ss_csrf = ss.get(f'{BASE_URL}/api/v1/security/csrf_token/', timeout=30).json()['result']
    ss.headers.update({'X-CSRFToken': ss_csrf, 'Referer': f'{BASE_URL}/'})

    dash = ss.get(f'{BASE_URL}/api/v1/dashboard/114', timeout=30).json()['result']
    pos  = json.loads(dash.get('position_json') or '{}')
    meta = json.loads(dash.get('json_metadata') or '{}')
    if 'MARKDOWN-portfolio-dl' in pos:
        pos['MARKDOWN-portfolio-dl']['meta']['code'] = md_content
        ss.put(f'{BASE_URL}/api/v1/dashboard/114', json={
            'position_json': json.dumps(pos), 'json_metadata': json.dumps(meta),
        }, timeout=30).raise_for_status()
        print(f'  Dashboard updated with {len(dated_assets)} link(s).')
    else:
        print('  MARKDOWN node not found — skipping dashboard update.')
else:
    print('No GITHUB_TOKEN — skipping release upload.')

print(f'\n[{datetime.now(IST):%H:%M:%S}] Done.')
print(f'  Loans:       {len(df):,}')
print(f'  Loan Amt:    Rs {df["Loan Amount"].sum()/1e7:.2f} Cr')
print(f'  Collected:   Rs {df["Total Collected"].sum()/1e7:.2f} Cr')
print(f'  Outstanding: Rs {df["Outstanding"].sum()/1e7:.2f} Cr')
