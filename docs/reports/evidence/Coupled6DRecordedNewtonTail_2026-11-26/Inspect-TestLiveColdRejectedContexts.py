"""ДЛЯ ТЕСТОВ. Читает только законченные бинарные записи живого923; не подменяет окончательную проверку пути и завершения процесса."""
from pathlib import Path
from runpy import run_path
import numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold923RecordedNewtonCold96_20261126';reference=base/'Cold920AffineRecoveryCold96_20261126'
read=run_path(str(here/'Read-TestColdRawCapture.py'));R=read['TestReader'];frames,meta=read['TestHistory'](root/'Native/Stock.RawHistory.bin',live=True)
actual=(root/'Native/Stock.RawHistory.bin').read_bytes()[:meta['complete_bytes']]
assert (reference/'Native/Stock.RawHistory.bin').read_bytes().startswith(actual)
known={}
for frame in frames:
    for i,field in enumerate(frame['fields']):known.setdefault(field.tobytes(),set()).add((i+1,float(frame['lambdas'][i])))
finished=[];partial=[]
for path in sorted((root/'Native').glob('NewtonRejected.*.bin')):
    r=R(path.read_bytes());r.position=0
    try:
        m,x,h,rr,b,d=[r.VectorTest() for _ in range(6)]
        A=np.array([r.VectorTest() for _ in range(61)]);end=r.VectorTest();n=int(r.VectorTest()[0])
        for _ in range(n):r.VectorTest();r.VectorTest();r.VectorTest()
    except EOFError:
        partial.append(path.name);continue
    assert r.position==len(r.raw)
    row=int(np.argmax(abs(rr[:56])))+1
    item=(path.name,int(m[1]),float(end[1]),row,float(m[5]),sorted(known.get(h.tobytes(),set())))
    finished.append(item)
print('LIVE ONLY;COMPLETE RAW PREFIX EXACT920;FRAMES',len(frames),'STATES',len(frames[-1]['fields']),
    'COMPLETE REJECTED RECORDS',len(finished),'PARTIAL FILES',partial)
for item in finished:
    if item[2]>.01 or item in finished[-4:]:print('RECORD;ITER;NORM;ROW;TARGET;ACCEPTED MATCHES',item)
