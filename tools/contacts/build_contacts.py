"""Rebuild reference/contacts.json from Noah's "GMs, BMs, Controllers, Fixed Ops" workbook.

Usage (from the repo root):
  python3 -I tools/contacts/build_contacts.py "<workbook.xlsx>" reference/stores.json reference/contacts.json [--confirmed YYYY-MM-DD]

Reads tabs by name: 'GM, BM, Controller ...', 'Fixed Ops ...', 'FMGSVCMGR', 'Parts EXT', 'Store  Addresses'
(prefix match, so re-dated tab names still work). Prints coverage gaps; review the git diff before committing.
A new store needs a line in MAP (and ADDR) below plus a row in stores.json.
"""
import datetime, json, re, sys, openpyxl
args = [a for a in sys.argv[1:] if not a.startswith('--')]
src, stores_path, out = args[:3]
confirmed = next((sys.argv[i + 1] for i, a in enumerate(sys.argv) if a == '--confirmed'), datetime.date.today().isoformat())
wb = openpyxl.load_workbook(src, data_only=True)
def tab(prefix):
    return next(ws for ws in wb if ws.title.strip().startswith(prefix))
stores = json.load(open(stores_path))['stores']
by_name = {s['name']: s for s in stores}
MAP = {  # workbook store label -> stores.json name
 'Acura':'Acura','Audi Henderson':'Audi Henderson','Audi Reno':'Audi Reno','Buick GMC Prescott':'Buick GMC Prescott',
 'Cadillac':'Cadillac','Chevy':'Chevrolet Las Vegas','CJDR':'CDJR Post Falls','Findlay Customs':'Findlay Customs',
 'Honda Flagstaff':'Honda Flagstaff','Honda Henderson':'Honda Henderson','Honda North':'Honda North','Honda Spokane':'Honda Spokane',
 'Hyundai Prescott':'Hyundai Prescott','Hyundai St. George':'Hyundai St George','Ineos':'INEOS','JLR LV':'Land Rover Las Vegas',
 'Kia LV':'Kia Las Vegas','Kia St George':'Kia St George','Lexus Spokane':'Lexus Spokane','Lincoln':'Lincoln',
 'LR Henderson':'JLR Henderson','LR Reno':'Land Rover Reno','Mazda':'Mazda Henderson','Motor Co':'Chevy GMC (Motor Co)',
 'Nissan':'Nissan Henderson','Subaru LV':'Subaru Las Vegas','Subaru Prescott':'Subaru Prescott','Subaru St. George':'Subaru St George',
 'Toyota Flagstaff':'Toyota Flagstaff','Toyota Henderson':'Toyota Henderson','Toyota Prescott':'Toyota Prescott',
 'Toyota Spokane':'Toyota Spokane','Volvo':'Volvo Las Vegas','VW Henderson':'VW Henderson','VW St George':'VW St George'}
ADDR = {  # address-sheet label -> workbook store label
 'Audi Henderson':'Audi Henderson','Audi Reno Tahoe':'Audi Reno','Findlay Acura':'Acura','Findlay Buick GMC Prescott':'Buick GMC Prescott',
 'Findlay Cadillac':'Cadillac','Findlay Chevrolet':'Chevy','Findlay Chysler Jeep Dodge Ram':'CJDR','Findlay Customs':'Findlay Customs',
 'Findlay Honda Flagstaff':'Honda Flagstaff','Findlay Honda Henderson':'Honda Henderson','Findlay Honda North':'Honda North',
 'Findlay Honda Spokane':'Honda Spokane','Findlay Hyundai Prescott':'Hyundai Prescott','Findlay Hyundai St. George':'Hyundai St. George',
 'Findlay INEOS Grenadier':'Ineos','Findlay Kia':'Kia LV','Findlay Kia St. George':'Kia St George','Findlay Lexus Spokane':'Lexus Spokane',
 'Findlay Lincoln':'Lincoln','Findlay Mazda':'Mazda','Findlay Motor Co':'Motor Co','Findlay Nissan':'Nissan','Findlay Subaru Prescott':'Subaru Prescott',
 'Findlay Subaru St. George':'Subaru St. George','Findlay Toyota Flagstaff':'Toyota Flagstaff','Findlay Toyota Henderson':'Toyota Henderson',
 'Findlay Toyota Prescott':'Toyota Prescott','Findlay Toyota Spokane':'Toyota Spokane','Findlay Volkswagen Henderson':'VW Henderson',
 'Findlay Volkswagen St. George':'VW St George','Findlay Volvo Cars':'Volvo','Jaguar Land Rover Las Vegas':'JLR LV',
 'Land Rover Henderson':'LR Henderson','Land Rover Reno':'LR Reno','Subaru Las Vegas':'Subaru LV'}
NAME_FIX = {'Jim Willaims': 'Jim Williams'}  # typo on Fixed Ops sheet; FMGSVCMGR sheet spells it Williams

def clean(s): return re.sub(r'\s+', ' ', str(s).replace(' ', ' ')).strip() if s is not None else None
def phone(p):
    if p is None: return None
    d = re.sub(r'\D', '', str(p))
    return f'{d[:3]}-{d[3:6]}-{d[6:]}' if len(d) == 10 else clean(p)
def email(e): return clean(e).lower() if e else None

def tab_date(prefix):
    m = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{2,4})', tab(prefix).title)
    if not m: return None
    mo, d, y = map(int, m.groups()); y += 2000 if y < 100 else 0
    return f'{y:04d}-{mo:02d}-{d:02d}'
GA, FA = tab_date('GM, BM, Controller'), tab_date('Fixed Ops')
recs = {lab: {'adp_code': by_name[n]['adp_code'], 'store': n, 'state': by_name[n]['state'] or None, 'workbook_label': lab,
              'address': None, 'parts_ext': None, 'exec': None, 'contacts': []} for lab, n in MAP.items()}
def add(lab, role, name, em, ph, as_of, **kw):
    name = NAME_FIX.get(clean(name), clean(name))
    recs[lab]['contacts'].append({'role': role, 'name': name, 'email': email(em), 'phone': phone(ph), 'as_of': as_of, **kw})

ws = tab('GM, BM, Controller')
for r in ws.iter_rows(min_row=3, values_only=True):
    _, lab, bm, bme, bmp, gm, gme, gmp, ct, cte, ctp, ex, st = r
    if not lab: continue
    lab = clean(lab)
    if '/' in gm:  # Toyota Henderson lists two GMs in one cell
        for n, e in zip(gm.split('/'), gme.split('/')):
            add(lab, 'General Manager', n, e, gmp, GA)
    else:
        add(lab, 'General Manager', gm, gme, gmp, GA)
    add(lab, 'Business Manager', bm, bme, bmp, GA)
    add(lab, 'Controller', ct, cte, ctp, GA)
    recs[lab]['exec'] = clean(ex)

ws = tab('Fixed Ops')
primary_svc = {}
for r in ws.iter_rows(min_row=3, values_only=True):
    _, lab, sm, sme, smp, pm, pme, pmp, ext = r[:9]
    if not lab: continue
    lab = clean(lab)
    add(lab, 'Service Manager', sm, sme, smp, FA)
    add(lab, 'Parts Manager', pm, pme, pmp, FA)
    recs[lab]['parts_ext'] = ext
    primary_svc[lab] = email(sme)

ws = tab('FMGSVCMGR')
for r in ws.iter_rows(min_row=3, values_only=True):
    lab, n, e = r
    if not lab: continue
    lab = clean(lab)
    if email(e) != primary_svc.get(lab):
        add(lab, 'Service (FMGSVCMGR list)', n, e, None, None, note='On the FMGSVCMGR service-manager email list; not the primary Service Manager on the Fixed Ops sheet')

pe = {clean(r[0]): r[1] for r in tab('Parts EXT').iter_rows(min_row=3, values_only=True) if r[0] and r[1]}
for lab, r in recs.items():
    if lab in pe and pe[lab] != r['parts_ext']:
        r['parts_ext_note'] = f"Fixed Ops tab said {r['parts_ext']}; Parts EXT tab says {pe[lab]} - using the Parts EXT tab"
        r['parts_ext'] = pe[lab]

ws = tab('Store')
for r in ws.iter_rows(min_row=2, values_only=True):
    lab = ADDR.get(clean(r[0]))
    if lab: recs[lab]['address'] = clean(r[1])

FLAGS = {
 ('Cadillac', 'Bryan Austin'): 'Email format differs from the name - confirmed correct by Noah 2026-11-19',
 ('Subaru LV', 'Yasir Abrahams'): 'Email format differs from the name - confirmed correct by Noah 2026-11-19',
}
for (lab, n), msg in FLAGS.items():
    for c in recs[lab]['contacts']:
        if c['name'] == n: c['note'] = msg

order = {s['adp_code']: i for i, s in enumerate(stores)}
data = {'_about': ('Store leadership contacts (GM, Service, Parts, Business Manager, Controller). One row per store, keyed by ADP company code '
                   '(same as reference/stores.json). Use this for names/emails when drafting store emails - never guess an address format. '
                   'Source: "GMs, BMs, Controllers, Fixed Ops" workbook from Noah. as_of = date in the source tab name, '
                   'labels only; confirmed_current = date Noah confirmed the list is up to date. Emails lower-cased; state comes from stores.json. '
                   'Rebuild with tools/contacts/build_contacts.py when Noah sends an updated workbook; do not hand-edit.'),
        'confirmed_current': confirmed,
        'stores': sorted(recs.values(), key=lambda r: order[r['adp_code']])}
json.dump(data, open(out, 'w'), indent=1, ensure_ascii=False)
open(out, 'a').write('\n')
missing = [s['name'] for s in stores if s['adp_code'] not in {r['adp_code'] for r in recs.values()}]
print('stores', len(recs), 'contacts', sum(len(r['contacts']) for r in recs.values()), 'stores.json not covered:', missing)
for r in recs.values():
    roles = {c['role'] for c in r['contacts']}
    need = {'General Manager', 'Service Manager', 'Parts Manager'} - roles
    if need or not r['address']: print('GAP', r['store'], need, r['address'])
