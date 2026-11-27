"""ДЛЯ ТЕСТОВ. Проверяет фактически восстановленные файлы, CRC книг и точную рабочую выдачу."""
from pathlib import Path
from hashlib import sha256
from zipfile import ZipFile
import json

repo=Path.cwd();base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
evidence=repo/'docs/reports/evidence/Coupled6DASDCoordinates_2026-11-27'
restored=base/'RestoredASDCoordinates_20261127'
manifest=json.loads((evidence/'Files.json').read_bytes())
books=[]
for relative,record in manifest['files'].items():
    path=restored/relative
    assert path.stat().st_size==record['bytes'] and sha256(path.read_bytes()).hexdigest().upper()==record['sha256'],relative
    if path.suffix.lower()=='.xlsm':
        with ZipFile(path) as archive:assert archive.testzip() is None,relative
        books.append(relative)
assert len(books)==manifest['exact_books']==17
assert len(manifest['files'])==2107
summary=json.loads((evidence/'Summary.json').read_bytes())
for name,record in summary['delivery']['output'].items():
    assert (repo/'workbook/output'/name).read_bytes()==(restored/'PublicOutput'/name).read_bytes(),name
    assert sha256((repo/'workbook/output'/name).read_bytes()).hexdigest().upper()==record['after_sha256'],name
for relative,digest in summary['working_sources'].items():
    assert (repo/relative).read_bytes()==(restored/'WorkingSource'/relative).read_bytes(),relative
    assert sha256((repo/relative).read_bytes()).hexdigest().upper()==digest,relative
proof=dict(actual_files=len(manifest['files']),exact_books=books,all_original_file_hashes_match=True,
           all_restored_book_zip_crc_verified=True,actual_public_output_exact_restored=True,
           working_sources_exact_restored=True,restored_destination=str(restored),whole_goal_completed=False)
(evidence/'RestorationProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
resume_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json'
resume=json.loads(resume_path.read_bytes())
resume['asd_coordinate_api_20261127']['restoration_proof']=proof
resume_path.write_text(json.dumps(resume,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('RESTORATION VERIFIED;2107 FILES;17 BOOKS CRC PASS;ACTUAL WORKING OUTPUT AND SOURCES EXACT')
