"""Download only workbook publications and their OA PDF locations; checkpoint every result."""
import os, json, time, hashlib, threading, urllib.request, urllib.error, urllib.parse
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl
from pypdf import PdfReader
import io

ROOT=Path('/Users/maura/Desktop/KnowldegeBase')
KEY=os.environ.get('OPENALEX_API_KEY','')
assert KEY, 'Missing locally configured API key'
STATE=ROOT/'download_records'; STATE.mkdir(exist_ok=True)
LOG=STATE/'attempts.jsonl'; lock=threading.Lock(); api_blocked=threading.Event()
def sheets(name):
    w=openpyxl.load_workbook(ROOT/'reconciled_master_corrected.xlsx',read_only=True,data_only=True)
    s=w[name];s.reset_dimensions(); rows=iter(s.values); headers=next(rows)
    data=[dict(zip(headers,r)) for r in rows];w.close();return data
master={r['ID']:r for r in sheets('Master')}
records=sheets('Source_records')
tasks={}
for r in records:
    sid=r['canonical_ID']; wid=r.get('OpenAlex_ID'); doi=r.get('DOI')
    if wid: tasks[(sid,wid)]={'ID':sid,'work':wid,'doi':doi}
    elif sid not in {x[0] for x in tasks} and doi and str(doi).startswith('10.'):
        tasks.setdefault((sid,'doi:'+doi),{'ID':sid,'work':None,'doi':doi})
# DOI lookup only when no source row already supplies a work ID for the study.
have_ids={sid for (sid,wid) in tasks if not wid.startswith('doi:')}
tasks={k:v for k,v in tasks.items() if not k[1].startswith('doi:') or k[0] not in have_ids}
previous=[]
if LOG.exists():
    for line in LOG.read_text().splitlines():
        try: previous.append(json.loads(line))
        except ValueError: pass
done={(r['ID'],r.get('url')):r for r in previous if r.get('result') in ('downloaded','duplicate_content')}
def log(**r):
    r['time']=time.strftime('%Y-%m-%d %H:%M:%S')
    with lock:
        with LOG.open('a') as f:f.write(json.dumps(r,ensure_ascii=False)+'\n')
    return r
def fetch(url, api=False):
    if api and api_blocked.is_set():return None,'api_paused'
    target=url
    if api:target += ('&' if '?' in target else '?')+urllib.parse.urlencode({'api_key':KEY})
    for attempt in range(2):
        try:
            req=urllib.request.Request(target,headers={'User-Agent':'RDM literature collection/1.0'})
            with urllib.request.urlopen(req,timeout=35) as response:
                body=response.read(80*1024*1024+1)
                if len(body)>80*1024*1024:return None,'file_over_80MB'
                return body,200
        except urllib.error.HTTPError as e:
            if api and e.code in (401,402,403,429):api_blocked.set()
            if e.code in (500,502,503,504) and attempt==0:time.sleep(1);continue
            return None,e.code
        except Exception:
            if attempt==0:time.sleep(1);continue
            return None,'network_error'
def run(task):
    sid=task['ID'];wid=task['work']; folder=ROOT/sid
    # Preserve already-collected studies, including manual folders referenced by aliases.
    if master[sid].get('Download') in ('manual_GS','manual_OA'):
        return log(ID=sid,work=wid,result='preserved_manual_collection')
    key=wid or task['doi']; cache=STATE/(hashlib.sha256(key.encode()).hexdigest()+'.json')
    if cache.exists():
        try:meta=json.loads(cache.read_text())
        except ValueError:meta=None
    else:meta=None
    if meta is None:
        lookup=wid or ('https://doi.org/'+task['doi'])
        body,status=fetch('https://api.openalex.org/works/'+urllib.parse.quote(lookup,safe=':/'),True)
        if body is None:return log(ID=sid,work=wid,result='metadata_unavailable',status=status)
        try:meta=json.loads(body)
        except ValueError:return log(ID=sid,work=wid,result='invalid_metadata')
        cache.write_text(json.dumps(meta))
    wid=str(meta.get('id','')).rsplit('/',1)[-1]
    if not wid.startswith('W'):return log(ID=sid,work=wid,result='missing_work_id')
    urls=[]
    if meta.get('has_content',{}).get('pdf'):
        urls.append(('https://content.openalex.org/works/'+wid+'.pdf','openalex_main',True))
    for loc in meta.get('locations') or []:
        url=loc.get('pdf_url')
        if loc.get('is_oa') and url and url.startswith(('https://','http://')):
            if url not in {u[0] for u in urls}:urls.append((url,loc.get('version') or 'oa_copy',False))
    if not urls:return log(ID=sid,work=wid,result='no_pdf_link')
    hashes={}
    if folder.exists():
        for p in folder.rglob('*.pdf'):
            hashes[hashlib.sha256(p.read_bytes()).hexdigest()]=str(p)
    results=[]
    for url,version,api in urls:
        prior=done.get((sid,url))
        if prior and prior.get('file') and Path(prior['file']).exists():continue
        name=f'{sid}_{wid}_{version}_{hashlib.sha256(url.encode()).hexdigest()[:10]}.pdf'
        dest=folder/name
        if dest.exists():continue
        body,status=fetch(url,api)
        if body is None:
            results.append(log(ID=sid,work=wid,url=url,result='unavailable',status=status));continue
        if b'%PDF-' not in body[:1024]:
            results.append(log(ID=sid,work=wid,url=url,result='not_pdf',status=status));continue
        try:
            pages=len(PdfReader(io.BytesIO(body),strict=False).pages)
            if not pages:raise ValueError()
        except Exception:
            results.append(log(ID=sid,work=wid,url=url,result='invalid_pdf',status=status));continue
        digest=hashlib.sha256(body).hexdigest()
        if digest in hashes:
            results.append(log(ID=sid,work=wid,url=url,result='duplicate_content',file=hashes[digest],sha256=digest));continue
        folder.mkdir(exist_ok=True)
        with dest.open('xb') as f:f.write(body)
        hashes[digest]=str(dest)
        results.append(log(ID=sid,work=wid,url=url,result='downloaded',file=str(dest),bytes=len(body),pages=pages,sha256=digest))
    return {'ID':sid,'result':'checked','saved':sum(r['result']=='downloaded' for r in results)}
print(f'Checking {len(tasks)} work/DOI records; existing manual collections preserved.',flush=True)
with ThreadPoolExecutor(max_workers=4) as pool:
    futures={pool.submit(run,t):t for t in tasks.values()}
    saved=0
    for n,future in enumerate(as_completed(futures),1):
        try:r=future.result()
        except Exception as e:r=log(ID=futures[future]['ID'],result='worker_error',error_type=type(e).__name__)
        saved+=r.get('saved',0)
        if n%10==0 or n==len(futures):print(f'Checked {n}/{len(futures)}; saved {saved} new PDFs; API paused={api_blocked.is_set()}',flush=True)
print('Finished. Detailed results are checkpointed in download_records/attempts.jsonl.',flush=True)
