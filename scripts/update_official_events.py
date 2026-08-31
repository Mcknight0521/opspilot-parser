#!/usr/bin/env python3
"""Build/update OpsPilot's small official-event store using only stdlib.
Runs in GitHub Actions, never in the user-facing Render request path.
"""
from __future__ import annotations
import argparse, concurrent.futures, datetime as dt, html, json, re, time
from pathlib import Path
from urllib.parse import urljoin
from urllib.request import Request, urlopen

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'data'/'official_events.json'
DGPA='https://www.dgpa.gov.tw'
CWA='https://rdc28.cwa.gov.tw'
COUNTIES=["基隆市","臺北市","新北市","桃園市","新竹市","新竹縣","苗栗縣","臺中市","彰化縣","南投縣","雲林縣","嘉義市","嘉義縣","臺南市","高雄市","屏東縣","宜蘭縣","花蓮縣","臺東縣","澎湖縣","金門縣","連江縣"]
UA='Mozilla/5.0 OpsPilot-EventSync/4.4'

# Terms that can appear near a generic「名稱」label on CWA pages but are not
# typhoon names.  In particular, this prevents wildlife/news labels such as
#「白海豚」from ever being promoted to a typhoon event.
INVALID_TYPHOON_NAMES={
    '白海豚','豪雨','大豪雨','強風','天氣','警報','颱風警報','海上','陸上',
    '發布','解除','臺灣','台灣','中央氣象署','氣象署','名稱','概況'
}

def _decode_quality(text):
    # Higher is better. Replacement characters are a strong sign of a bad decode.
    cjk=sum(1 for ch in text if '\u3400' <= ch <= '\u9fff')
    replacement=text.count('\ufffd')
    controls=sum(1 for ch in text if ord(ch)<32 and ch not in '\n\r\t')
    return cjk*4 - replacement*200 - controls*20

def decode_web_bytes(raw, declared=None):
    """Decode Taiwanese government pages defensively.

    Some legacy pages/attachments advertise an incorrect or incomplete charset.
    Prefer a strict UTF-8 decode when possible, then honor meta/header hints, then
    try Big5/CP950.  Pick the candidate that yields real CJK text without U+FFFD.
    """
    head=raw[:8192].decode('ascii',errors='ignore')
    meta=[]
    for pat in (r'charset\s*=\s*["\']?([A-Za-z0-9._-]+)', r'encoding\s*=\s*["\']([^"\']+)'):
        m=re.search(pat,head,re.I)
        if m: meta.append(m.group(1))
    candidates=[]
    for enc in ['utf-8-sig','utf-8',*(meta or []),declared,'cp950','big5']:
        if not enc or enc.lower() in {x.lower() for x,_ in candidates}:
            continue
        try:
            text=raw.decode(enc,errors='strict')
            candidates.append((enc,text))
        except (UnicodeDecodeError,LookupError):
            pass
    if not candidates:
        return raw.decode('utf-8',errors='replace')
    candidates.sort(key=lambda x:_decode_quality(x[1]),reverse=True)
    return candidates[0][1]

def get(url, timeout=12):
    req=Request(url,headers={'User-Agent':UA,'Accept-Language':'zh-TW,zh;q=0.9','Connection':'close'})
    with urlopen(req,timeout=timeout) as r:
        raw=r.read(2*1024*1024); declared=r.headers.get_content_charset()
    return decode_web_bytes(raw,declared)

def strip(s):
    s=re.sub(r'<script[\s\S]*?</script>|<style[\s\S]*?</style>',' ',s or '',flags=re.I)
    return re.sub(r'\s+',' ',html.unescape(re.sub(r'<[^>]+>',' ',s))).strip()

def iso_roc(y,m,d): return f'{int(y)+1911:04d}-{int(m):02d}-{int(d):02d}'

def dgpa_list(page):
    url=f'{DGPA}/informationlist?page={page}&uid=374'; text=get(url); out=[]; seen=set()
    for m in re.finditer(r'<a[^>]+href=["\']([^"\']*information\?[^"\']*pid=\d+[^"\']*)["\'][^>]*>([\s\S]*?)</a>',text,re.I):
        href,label=m.group(1),strip(m.group(2)); around=strip(text[max(0,m.start()-260):m.end()+260])
        dm=re.search(r'(\d{2,3})年\s*(\d{1,2})月\s*(\d{1,2})日',label) or re.search(r'(\d{2,3})年\s*(\d{1,2})月\s*(\d{1,2})日',around)
        if not dm: continue
        u=urljoin(url,html.unescape(href)); key=(u,dm.group(0))
        if key in seen: continue
        seen.add(key); out.append({'date':iso_roc(*dm.groups()),'url':u,'title':label or around[:160]})
    return out

def nds_links(text,base):
    out=[]
    for m in re.finditer(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>([\s\S]*?)</a>',text or '',re.I):
        href,label=html.unescape(m.group(1)),strip(m.group(2))
        if re.search(r'nds(?:E)?\.html|停止上班|停止辦公|停止上課',href+' '+label,re.I) and not re.search(r'ndsE\.html',href,re.I):
            u=urljoin(base,href)
            if u not in out: out.append(u)
    return out

def parse_counties(text,date,source_url):
    plain=strip(text); events=[]
    for i,county in enumerate(COUNTIES):
        pos=plain.find(county)
        if pos<0: continue
        tail=plain[pos+len(county):]
        stops=[tail.find(c) for c in COUNTIES if tail.find(c)>=0]
        chunk=tail[:min(stops) if stops else 700].strip()
        closed=bool(re.search(r'停止上班|停止上課|停止辦公|停班|停課',chunk)) and not bool(re.search(r'照常上班.{0,50}照常上課|正常上班.{0,50}正常上課',chunk))
        if not closed: continue
        countywide=bool(re.match(r'^[：: ]*(?:今天|明天)?停止上班[、,， ]*(?:今天|明天)?停止上課',chunk))
        events.append({'type':'closure','date':date,'region':county,'name':'全縣停班停課' if countywide else '部分地區停班停課','scope':'countywide' if countywide else 'partial','detail':chunk[:500],'source':'行政院人事行政總處','sourceUrl':source_url})
    return events

def resolve_dgpa(item):
    try:
        detail=get(item['url']); links=nds_links(detail,item['url'])
        for u in links[:3]:
            try:
                ev=parse_counties(get(u),item['date'],u)
                if ev: return ev
            except Exception: pass
    except Exception: pass
    return []

def sync_dgpa(existing, full, years_back):
    today=dt.date.today(); floor=dt.date(today.year-years_back,1,1) if full else today-dt.timedelta(days=45)
    items=[]
    for page in range(1,301):
        try: rows=dgpa_list(page)
        except Exception as e:
            print('DGPA page error',page,e); break
        if not rows: break
        dated=[]
        for x in rows:
            try: d=dt.date.fromisoformat(x['date']); dated.append(d)
            except: continue
            if d>=floor: items.append(x)
        if dated and min(dated)<floor: break
        time.sleep(.08)
    uniq={x['url']:x for x in items}.values(); fresh=[]
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as ex:
        for evs in ex.map(resolve_dgpa,uniq): fresh.extend(evs)
    keep=[e for e in existing if e.get('type')!='closure' or e.get('date','')<floor.isoformat()]
    return keep+fresh, floor.isoformat()

def _clean_typhoon_name(raw):
    name=strip(raw or '')
    name=re.sub(r'^[：:|｜\s]+|[：:|｜\s]+$','',name)
    name=re.sub(r'(?:颱風|台風)$','',name).strip()
    if not name or name in INVALID_TYPHOON_NAMES:
        return None
    if len(name)>12 or re.search(r'\d|20\d{2}|發布|解除|警報|豪雨|強風|海上|陸上|氣象|白海豚',name):
        return None
    if not any('\u3400' <= ch <= '\u9fff' for ch in name):
        return None
    return name

def parse_typhoon(text,tyid):
    p=strip(text)
    if not re.search(r'颱風概況表|發布時間',p): return None

    # Do not use a free-floating「名稱」match: CWA pages contain other named
    # entities, which previously allowed unrelated labels (e.g. 白海豚) through.
    # Prefer labels that explicitly identify the typhoon name and only fall back
    # to a tightly bounded「名稱」occurring next to「颱風」.
    name=None
    name_patterns=(
        r'(?:颱風名稱|中文名稱)\s*[:|｜]?\s*([^\s(（|｜:：]{1,12})',
        r'颱風.{0,18}?名稱\s*[:|｜]?\s*([^\s(（|｜:：]{1,12})',
        r'名稱\s*[:|｜]?\s*([^\s(（|｜:：]{1,12})\s*(?:颱風|台風)'
    )
    for pat in name_patterns:
        m=re.search(pat,p,re.I)
        if m:
            name=_clean_typhoon_name(m.group(1))
            if name: break

    sea=(re.search(r'海上\s*(20\d{2}-\d{2}-\d{2})\s*\d{2}:\d{2}',p) or [None,None])[1]
    land=(re.search(r'陸上\s*(20\d{2}-\d{2}-\d{2})\s*\d{2}:\d{2}',p) or [None,None])[1]
    releases=[m.group(1) or m.group(2) for m in re.finditer(r'解除時間[^2]*(?:陸上\s*)?(20\d{2}-\d{2}-\d{2})\s*\d{2}:\d{2}|海上\s*(20\d{2}-\d{2}-\d{2})\s*\d{2}:\d{2}',p)]
    if not name or not (sea or land): return None
    start=land or sea; end=sorted([x for x in releases if x])[-1] if releases else start
    return name,start,end

def sync_cwa(existing, full, years_back):
    today=dt.date.today(); first=today.year-years_back if full else today.year-1; fresh=[]
    for year in range(first,today.year+1):
        def one(i):
            tid=f'{year}{i:02d}'; url=f'{CWA}/TDB/public/typhoon_detail?typhoon_id={tid}'
            try:
                t=parse_typhoon(get(url,8),tid)
                if not t:return []
                name,start,end=t; a=dt.date.fromisoformat(start); b=dt.date.fromisoformat(end); out=[]
                while a<=b:
                    out.append({'type':'typhoon','date':a.isoformat(),'name':f'{name}颱風','detail':'中央氣象署颱風警報期間','source':'中央氣象署颱風資料庫','sourceUrl':url}); a+=dt.timedelta(days=1)
                return out
            except Exception:return []
        with concurrent.futures.ThreadPoolExecutor(max_workers=7) as ex:
            for rows in ex.map(one,range(1,36)): fresh.extend(rows)
    floor=f'{first}-01-01'; keep=[e for e in existing if e.get('type')!='typhoon' or e.get('date','')<floor]
    return keep+fresh,floor

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--full',action='store_true'); ap.add_argument('--years-back',type=int,default=10); args=ap.parse_args()
    try: payload=json.loads(OUT.read_text(encoding='utf-8'))
    except: payload={'events':[]}
    events=payload.get('events') or []
    # v4.4 self-heals stores produced by an earlier bad charset decode. If any
    # replacement/mojibake marker is present, rebuild the requested history once.
    sample=' '.join(str(e.get(k,'')) for e in events for k in ('region','name','detail','source'))
    polluted=('\ufffd' in sample)
    bootstrap=not events or (payload.get('coverage') or {}).get('status')=='bootstrap_pending'
    full=args.full or bootstrap or polluted
    if polluted:
        print('Detected mojibake in existing event store; forcing clean full rebuild')
        events=[]
    events,dgpa_floor=sync_dgpa(events,full,args.years_back)
    events,cwa_floor=sync_cwa(events,full,args.years_back)

    # Remove any bad historical typhoon labels already present in the store so
    # incremental runs self-heal instead of retaining an old false positive.
    clean_events=[]
    for e in events:
        if e.get('type')=='typhoon':
            raw_name=str(e.get('name') or '')
            base=re.sub(r'(?:颱風|台風)$','',raw_name).strip()
            if not _clean_typhoon_name(base):
                continue
        clean_events.append(e)
    events=clean_events

    dedup={}
    for e in events:
        key=(e.get('type'),e.get('date'),e.get('region',''),e.get('name',''))
        dedup[key]=e
    events=sorted(dedup.values(),key=lambda e:(e.get('date',''),e.get('type',''),e.get('region',''),e.get('name','')))
    now=dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat()
    out={'schemaVersion':1,'updatedAt':now,'coverage':{'status':'ready','dgpaFrom':dgpa_floor,'cwaFrom':cwa_floor,'through':dt.date.today().isoformat(),'mode':'full' if full else 'incremental'},'events':events}
    OUT.parent.mkdir(parents=True,exist_ok=True); OUT.write_text(json.dumps(out,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
    print(f'Wrote {len(events)} events to {OUT}; full={full}')
if __name__=='__main__': main()
