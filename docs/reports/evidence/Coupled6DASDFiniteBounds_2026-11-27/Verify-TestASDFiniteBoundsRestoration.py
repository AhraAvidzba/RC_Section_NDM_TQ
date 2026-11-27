"""ДЛЯ ТЕСТОВ. Проверяет восстановление, обязательный состав, CRC книг и при необходимости рабочую выдачу."""
from pathlib import Path
from argparse import ArgumentParser
from hashlib import sha256
from zipfile import ZipFile
import json

parser=ArgumentParser(description=__doc__)
parser.add_argument('--evidence',type=Path,default=Path(__file__).resolve().parent)
parser.add_argument('--restored',type=Path,required=True)
parser.add_argument('--verify-working',action='store_true')
args=parser.parse_args();repo=Path.cwd();evidence=args.evidence;restored=args.restored
manifest=json.loads((evidence/'Files.json').read_bytes());books=[]
assert manifest['roots']==['Cold1128ASDPathPhysicalLimits_20261127','Cold1129ASDFiniteBounds_20261127','Cold1130ASDFiniteBoundsOutput_20261127']
assert len(manifest['files'])==976 and manifest['exact_books']==8
for relative,record in manifest['files'].items():
    path=restored/relative
    assert path.stat().st_size==record['bytes'] and sha256(path.read_bytes()).hexdigest().upper()==record['sha256'],relative
    if path.suffix.lower()=='.xlsm':
        with ZipFile(path) as archive:assert archive.testzip() is None,relative
        books.append(relative)
assert len(books)==8
summary=json.loads((evidence/'Summary.json').read_bytes())
for point in range(75):
    data=json.loads((restored/f'Cold1128ASDPathPhysicalLimits_20261127/Native/Point{point:02d}.json').read_bytes())
    assert data['point']==point and len(data['coordinates'])==141 and data['engine_version']=='3.8.0'
for folder in manifest['roots'][1:]:
    report=(restored/folder/'Native/RealColdReplay.txt').read_text('utf-8-sig')
    assert 'TOTAL_COUPLED_ASD_COORDINATES_TEST: passed=3031; failed=0' in report
    assert 'SOURCE_UNCHANGED: True' in report and 'RUNTIME ERROR' not in report
    closure=json.loads((restored/folder/'Native/RealColdReplay.txt.ProcessClosure.json').read_bytes())
    assert closure['Passed'] and closure['ChildHasExited'] and closure['ChildExitCode']==0
if args.verify_working:
    for name,record in summary['delivery']['output'].items():
        assert (repo/'workbook/output'/name).read_bytes()==(restored/'PublicOutput'/name).read_bytes(),name
        assert sha256((repo/'workbook/output'/name).read_bytes()).hexdigest().upper()==record['after_sha256'],name
    for relative in list(summary['working_sources'])+['AGENTS.md']:
        assert (repo/relative).read_bytes()==(restored/'WorkingSource'/relative).read_bytes(),relative
    assert sha256((repo/'AGENTS.md').read_bytes()).hexdigest().upper()==summary['project_rules_sha256']
proof=dict(actual_files=976,exact_books=books,all_original_file_hashes_match=True,all_restored_book_zip_crc_verified=True,
           mandatory_histories_and_native_suites_present=True,actual_working_output_and_source_verified=args.verify_working,
           restored_destination=str(restored),whole_goal_completed=False)
(evidence/'RestorationProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
if args.verify_working:
    path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';resume=json.loads(path.read_bytes())
    resume['asd_finite_bounds_20261127']['restoration_proof']=proof
    path.write_text(json.dumps(resume,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('FINITE ASD LIMIT EVIDENCE VERIFIED;976 FILES;8 BOOKS CRC PASS;75 HISTORIES;TWO NATIVE SUITES')
