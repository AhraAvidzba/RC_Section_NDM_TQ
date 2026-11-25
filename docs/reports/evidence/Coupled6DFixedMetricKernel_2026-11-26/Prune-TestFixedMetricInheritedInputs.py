"""ДЛЯ ТЕСТОВ. Исключает из нового архива только ненужные унаследованные входы предшествующих экспериментов; используемые байты сохраняются точно."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
import tarfile,json
here=Path(__file__).resolve().parent;out=Path.cwd()/'docs/reports/evidence/Coupled6DFixedMetricKernel_2026-11-26'
archive=out/'RawEvidence.tar.xz';allowed={'Reverse1.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json','ArcCallAllowance.bin','OperationalArcRatio.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin','ArcScanFields.bin'}
with tarfile.open(archive,'r:xz') as z:
    pins=json.loads(z.extractfile('RawFiles.json').read());raw={};excluded={}
    for name,digest in pins.items():
        content=z.extractfile(name).read();assert sha256(content).hexdigest().upper()==digest
        parts=name.split('/')
        if len(parts)==3 and parts[1]=='Input' and parts[2] not in allowed:excluded[name]=dict(sha256=digest,bytes=len(content))
        else:raw[name]=content
assert excluded and all('/Input/' in name for name in excluded)
fresh={name:sha256(data).hexdigest().upper() for name,data in raw.items()};temporary=out/'RawEvidencePruned.tmp.tar.xz'
assert not temporary.exists()
with tarfile.open(temporary,'w:xz',preset=6) as z:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(fresh,indent=2,sort_keys=True)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);z.addfile(info,BytesIO(data))
with tarfile.open(temporary,'r:xz') as z:
    assert {m.name for m in z.getmembers()}==set(raw)|{'RawFiles.json'}
    for name,digest in fresh.items():assert sha256(z.extractfile(name).read()).hexdigest().upper()==digest
temporary.replace(archive)
summary=json.loads((out/'Summary.json').read_bytes());summary.update(raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),unused_inherited_previous_experiment_inputs_excluded=excluded)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for name in (Path(__file__).name,'Record-TestFixedMetricKernel.py'):(out/name).write_bytes((here/name).read_bytes())
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file() and p.name!='Files.json'}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('RETAINED',len(raw),'USED RAW FILES;OMITTED',len(excluded),'UNUSED INHERITED INPUTS;ARCHIVE BYTES',archive.stat().st_size)
