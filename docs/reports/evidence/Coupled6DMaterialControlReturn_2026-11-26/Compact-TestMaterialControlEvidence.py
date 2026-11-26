"""ДЛЯ ТЕСТОВ. Устраняет побайтовые дубликаты архива стандартными TAR-ссылками, проверяя каждый логический файл до замены."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
import json,tarfile,shutil,lzma
repo=Path.cwd();out=repo/'docs/reports/evidence/Coupled6DMaterialControlReturn_2026-11-26'
source=out/'RawEvidence.tar.xz';target=out/'RawEvidence.compact.tar.xz'
if target.exists():
    assert target.resolve().is_relative_to(out.resolve())
    residue=target.read_bytes()
    assert not residue or (len(residue)==32 and lzma.decompress(residue)==b'')
    target.unlink() # Пустой XZ неудачной инициализации Python3.12; исходный архив не затрагивается.
before_hash=sha256(source.read_bytes()).hexdigest().upper();before_size=source.stat().st_size
canonical={};actual={};aliases={};expected=None
with tarfile.open(source,'r|xz') as original,lzma.open(target,'wb',preset=6) as packed,tarfile.open(fileobj=packed,mode='w|') as compact:
    for info in original:
        assert info.isfile() and not Path(info.name).is_absolute() and '..' not in Path(info.name).parts
        data=original.extractfile(info).read();digest=sha256(data).hexdigest().upper()
        if info.name=='RawFiles.json':expected=json.loads(data)
        else:actual[info.name]=digest
        entry=tarfile.TarInfo(info.name)
        if info.name!='RawFiles.json' and digest in canonical:
            entry.type=tarfile.LNKTYPE;entry.linkname=canonical[digest];aliases[info.name]=entry.linkname
            compact.addfile(entry)
        else:
            entry.size=len(data);compact.addfile(entry,BytesIO(data));canonical[digest]=info.name
assert actual==expected and aliases
checked={};read_pins=None
with tarfile.open(target,'r|xz') as compact:
    for info in compact:
        if info.islnk():
            assert info.linkname in checked;digest=checked[info.linkname]
        else:
            assert info.isfile();data=compact.extractfile(info).read();digest=sha256(data).hexdigest().upper()
            if info.name=='RawFiles.json':read_pins=json.loads(data);continue
        assert digest==expected[info.name],info.name;checked[info.name]=digest
assert checked==expected==read_pins
backup=Path('D:/CodexScratch/RC_Section_NDM_TQ/EvidenceArchiveLayouts/MaterialControlReturnFullCopy_20261126.tar.xz')
assert backup.resolve().is_relative_to(Path('D:/CodexScratch/RC_Section_NDM_TQ').resolve()) and not backup.exists()
backup.parent.mkdir(parents=True,exist_ok=True)
shutil.move(source,backup);target.replace(source)
proof=dict(original_full_copy_archive_sha256=before_hash,original_full_copy_archive_bytes=before_size,
    compact_archive_sha256=sha256(source.read_bytes()).hexdigest().upper(),compact_archive_bytes=source.stat().st_size,
    logical_files=len(actual),hardlink_aliases=len(aliases),all_logical_sha256_preserved_and_read_back=True,
    original_retained_at=str(backup),tar_links_only_to_preceding_verified_files=True)
(out/'ArchiveCompactionProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
(out/Path(__file__).name).write_bytes(Path(__file__).read_bytes())
summary=json.loads((out/'Summary.json').read_bytes());summary.update(raw_archive_bytes=source.stat().st_size,
    raw_archive_hardlink_aliases=len(aliases),all_raw_hashes_read_back=True)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
hashes={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file() and p.name!='Files.json'}
(out/'Files.json').write_text(json.dumps(hashes,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('COMPACT ARCHIVE',source.stat().st_size,'BYTES;ALL',len(actual),'LOGICAL HASHES VERIFIED;',len(aliases),'DUPLICATES LINKED')
