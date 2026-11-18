import json, openpyxl, datetime as dt
U='/mnt/user-data/uploads/'
def rows(f,sheet='Equity'):
    ws=openpyxl.load_workbook(U+f,data_only=True)[sheet]
    return [[c for c in r if c is not None] for r in ws.iter_rows(values_only=True)]
trades=[];hdr=False
for r in rows('tradebook-YN5207-EQ.xlsx'):
    if r and r[0]=='Symbol': hdr=True; continue
    if hdr and len(r)>8:
        trades.append(dict(sym=r[0],date=str(r[2])[:10],type=r[6],qty=float(r[8]),price=float(r[9])))
hold={};hdr=False
for r in rows('holdings-YN5207.xlsx'):
    if r and r[0]=='Symbol': hdr=True; continue
    if hdr and len(r)>10:
        hold[r[0].replace('-E','')]=dict(sector=r[2],ltp=float(r[9]),avg=float(r[8]),qty=float(r[3]))
flows=[];hdr=False
for r in rows('ledger-YN5207.xlsx'):
    if r and r[0]=='Particulars': hdr=True; continue
    if hdr and len(r)>=6 and isinstance(r[1],(str,dt.date,dt.datetime)) and r[3] in('Bank Receipts','Bank Payments'):
        d=str(r[1])[:10]; amt=float(r[4]) if r[3]=='Bank Receipts' else -float(r[4])
        # columns: particulars,date,cc,voucher,debit,credit,bal -> empty cells dropped, so recompute
        flows.append(dict(date=d,amt=None,raw=r))
# robust ledger parse (zero cells are 0.0 not None, so positions hold)
flows=[]
for r in rows('ledger-YN5207.xlsx'):
    if len(r)==7 and r[3] in('Bank Receipts','Bank Payments'):
        flows.append(dict(date=str(r[1])[:10],amt=float(r[5])-float(r[4])))
assert all(t['type']=='buy' for t in trades)
data=dict(asof='2026-11-17',fyStart='2026-04-01',trades=trades,holdings=hold,flows=flows)
json.dump(data,open('data.json','w'),indent=1)
print(len(trades),len(hold),flows)
