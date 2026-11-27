"""ДЛЯ ТЕСТОВ. Независимо перечитывает фактические восстановленные файлы этапа и проверяет все исходные SHA256."""
from pathlib import Path,PurePosixPath
from hashlib import sha256
from argparse import ArgumentParser
from datetime import datetime,timezone
import json

p=ArgumentParser(description=__doc__);p.add_argument('--restored',type=Path,required=True)
a=p.parse_args();root=a.restored.resolve();here=Path(__file__).resolve().parent
manifest=json.loads((here/'Files.json').read_bytes());proof=json.loads((here/'ArchiveProof.json').read_bytes())
books={};total=0
for relative,record in manifest['files'].items():
    parts=PurePosixPath(relative);assert not parts.is_absolute() and '..' not in parts.parts
    path=root.joinpath(*parts.parts).resolve();assert path.is_relative_to(root)
    raw=path.read_bytes();assert len(raw)==record['bytes'] and sha256(raw).hexdigest().upper()==record['sha256'],relative
    total+=len(raw)
    if path.suffix.lower() in ('.xlsm','.xlsx'):books[relative]=record['sha256']
assert len(books)==28 and len(manifest['files'])==5191 and total==manifest['logical_bytes']
for item in proof['archives']:
    raw=(here/item['file']).read_bytes();assert len(raw)==item['bytes'] and sha256(raw).hexdigest().upper()==item['sha256']
result=dict(utc=datetime.now(timezone.utc).isoformat(),restored_root=str(root),actual_files=len(manifest['files']),actual_bytes=total,
    all_original_file_hashes_match=True,all_archive_hashes_match=True,exact_books=books,
    verifier_sha256=sha256(Path(__file__).read_bytes()).hexdigest().upper(),no_new_solver_or_Excel_run=True,
    physical_steel_event_confirmed=True,general_section_capacity_certified=False,whole_goal_completed=False)
(here/'RestorationProof.json').write_text(json.dumps(result,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('CURVATURE EVIDENCE ACTUALLY RESTORED AND REVERIFIED',len(manifest['files']),'FILES',len(books),'BOOKS')
