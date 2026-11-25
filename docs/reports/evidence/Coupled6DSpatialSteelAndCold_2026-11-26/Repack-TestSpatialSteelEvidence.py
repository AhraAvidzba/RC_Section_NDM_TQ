"""ДЛЯ ТЕСТОВ. Сжимает те же доказательства большим словарём XZ и повторно проверяет каждый исходный байт."""
from pathlib import Path
from hashlib import sha256
import ctypes,json,tarfile
repo=Path.cwd();out=repo/'docs/reports/evidence/Coupled6DSpatialSteelAndCold_2026-11-26'
class TestMemory(ctypes.Structure):
    """Поля GlobalMemoryStatusEx для резерва физической памяти и commit перед большим словарём."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(name,ctypes.c_ulonglong) for name in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory)
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
file=out/'RawEvidence.tar.xz';temp=out/'RawEvidence.repacked.tar.xz';assert not temp.exists()
before=dict(bytes=file.stat().st_size,sha256=sha256(file.read_bytes()).hexdigest().upper())
with tarfile.open(file,'r:xz') as source,tarfile.open(temp,'w:xz',preset=9) as destination:
    names=source.getnames();pins=json.load(source.extractfile('RawFiles.json'))
    for name in names:
        member=source.getmember(name);assert member.isfile()
        destination.addfile(member,source.extractfile(member))
with tarfile.open(temp,'r:xz') as check:
    assert check.getnames()==names
    for name,digest in pins.items():assert sha256(check.extractfile(name).read()).hexdigest().upper()==digest,name
    assert json.load(check.extractfile('RawFiles.json'))==pins
assert temp.stat().st_size<file.stat().st_size
temp.replace(file)
proof=dict(before=before,after=dict(bytes=file.stat().st_size,sha256=sha256(file.read_bytes()).hexdigest().upper()),
           raw_files=len(pins),all_raw_bytes_exact=True,free_physical_before=memory.free_physical,free_commit_before=memory.free_page,
           xz_preset=9,working_code_or_books_changed=False)
(out/'CompressionProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
(out/Path(__file__).name).write_bytes(Path(__file__).read_bytes())
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file() and p.name!='Files.json'}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('SAME',len(pins),'RAW FILES;ARCHIVE BYTES',before['bytes'],'TO',file.stat().st_size,'ALL SHA256 VERIFIED')
