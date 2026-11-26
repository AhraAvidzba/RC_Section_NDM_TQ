"""ДЛЯ ТЕСТОВ. Закрывает этап943–946 воспроизводимым архивом и подготовленным журналом запрошенной пользователем паузы."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import json,re,tarfile,ctypes
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={943:'Cold943AffineControlPredictor_20261126',944:'Cold944AffineControlReturn_20261126',
    945:'Cold945AcceptedAffinePredictor_20261126',946:'Cold946AcceptedAffineReturn_20261126'}
roots={n:base/name for n,name in names.items()}
statefile=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statefile.read_bytes())
assert not state['active_jobs'] and state['active_exec_session'] is None
proofs={str(n):json.loads((root/('AffineControlPredictorProof.json' if n in (943,945) else 'OriginalControlReturnProof.json')).read_bytes()) for n,root in roots.items()}
for n,p in proofs.items():
    assert p['native']['own_process_closed'] and p['native']['failed']==0 and not p['physical_capacity_certified']
    if int(n) in (944,946):assert p['original_equilibria_verified']==2 and all(c['steps']==3 for c in p['cases'])
prior=repo/'docs/reports/evidence/Coupled6DMaterialControlReturn_2026-11-26'
priorindex=json.loads((prior/'Files.json').read_bytes());baseline=prior/'RawEvidence.tar.xz'
assert sha256(baseline.read_bytes()).hexdigest().upper()==priorindex[baseline.name]
with tarfile.open(baseline,'r:xz') as archive:actual_book=archive.extractfile('941/ActualCompiledWorkbook.xlsm').read()
book_hash=sha256(actual_book).hexdigest().upper()
for root in roots.values():assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==book_hash
class TestMemory(ctypes.Structure):
    """Проверяет свободную память перед упаковкой завершённых доказательств."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
m=TestMemory();m.length=ctypes.sizeof(m);assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m))
assert m.free_physical>3*1024**3 and m.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DSecantControlReturn_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','docs/regression/CodeUpdate.json','CompiledSourceReuse.json','PointTraceSourceComposition.json',
    'CompilePointClosure.json','CompileClosure.json','Run.ndjson','Run-TestRealColdReplay.ps1','Guard-IsolatedOffice.ps1',
    'tools/build_workbook/Run-Audit02ModeTests.ps1','tools/build_workbook/Run-FocusedKaniWatchdog.ps1',
    'SourceLowMaterialRoot.bin','SourceHighMaterialRoot.bin','SourceAffineMatrices.bin','SourceAffinePredictorProof.json','AffinePredictorInput.json']
raw={}
for n,root in roots.items():
    key='AffineControlPredictorProof.json' if n in (943,945) else 'OriginalControlReturnProof.json'
    files=fixed+[key]
    files += [p.relative_to(root).as_posix() for folder in ('Native','Input','src','tests') for p in (root/folder).rglob('*')
        if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{n}/{name}']=(root/name).read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};canonical={};aliases=0
archivefile=out/'RawEvidence.tar.xz'
with tarfile.open(archivefile,'w:xz',preset=6) as archive:
    for name,data in raw.items():
        info=tarfile.TarInfo(name);digest=pins[name]
        if digest in canonical:
            info.type=tarfile.LNKTYPE;info.linkname=canonical[digest];archive.addfile(info);aliases+=1
        else:
            canonical[digest]=name;info.size=len(data);archive.addfile(info,BytesIO(data))
    data=(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode();info=tarfile.TarInfo('RawFiles.json');info.size=len(data);archive.addfile(info,BytesIO(data))
checked={};readpins=None
with tarfile.open(archivefile,'r|xz') as archive:
    for info in archive:
        if info.islnk():assert info.linkname in checked;digest=checked[info.linkname]
        else:
            data=archive.extractfile(info).read();digest=sha256(data).hexdigest().upper()
            if info.name=='RawFiles.json':readpins=json.loads(data);continue
        assert digest==pins[info.name];checked[info.name]=digest
assert checked==pins==readpins
helpers=set('''Stage-TestAffineControlPredictor.py Check-TestAffineControlPredictor.py
Stage-TestAffineControlReturn.py Check-TestAffineControlReturn.py Stage-TestAcceptedAffinePredictor.py
Check-TestAcceptedAffinePredictor.py Stage-TestAcceptedAffineReturn.py Check-TestAcceptedAffineReturn.py
Read-TestRecordedYieldCoordinate.py Set-TestActiveStage.py Set-TestTerminalStage.py'''.split())
helpers.add(Path(__file__).name);todo=list(helpers)
while todo:
    p=here/todo.pop();assert p.is_file()
    for dep in re.findall(r'[A-Za-z][A-Za-z0-9_-]*\.(?:py|ps1|bas\.txt|cls\.txt)',p.read_text('utf-8-sig')):
        if dep not in helpers and (here/dep).is_file():helpers.add(dep);todo.append(dep)
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
reference=dict(archive=baseline.relative_to(repo).as_posix(),archive_sha256=priorindex[baseline.name],
    compiled_workbook_member='941/ActualCompiledWorkbook.xlsm',compiled_workbook_sha256=book_hash,
    all943_to946_books_exact_reference=True,source_audit257_inherited_from941=True,actual_project_recompiled=False,
    reproduction_requires_parent_evidence_chain=True)
(out/'CompiledBookReference.json').write_text(json.dumps(reference,sort_keys=True,indent=2)+'\n',encoding='utf-8')
summary=dict(proofs=proofs,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),raw_archive_bytes=archivefile.stat().st_size,
    exact_duplicate_hardlinks=aliases,all_logical_raw_hashes_read_back=True,compiled_book_reference=reference,
    scope='read-only predictors and original61 solve from immutable history25; not a fresh capacity search',
    high_material938_steps=17,return946_steps=3,combined_high_material_and_original_return_steps=20,
    earlier_failed_primary_and_required_half_check_effort_not_included=True,
    user_requested_pause_after_current_stage=True,all_numerical_stages_terminal=True,
    physical_capacity_certified=False,cold_gain_claimed=False,working_release_changed=False,full_acceptance_not_run=True)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
index={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(index,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='paused',pause_requested_explicitly=True,
    pause_request_revoked=False,pause_requested_after_logical_stage=True,active_exec_session=None,active_jobs=[],
    all_started_stages_terminal=True,latest_secant_control_return_evidence=out.relative_to(repo).as_posix(),
    latest_secant946_proof=proofs['946'],goal_tool_pause_after_git_checkpoint_requested=True,
    previous_goal_turn_classification='progress:943-944 original61 solve from automatic secant of two material equilibria;945-946 replace lower material root by accepted25 and verify same root in3 steps. No physical limit/cold gain. User requests pause after current stage.',
    next_action='PAUSED at explicit user request after946. On explicit resume: generic bounded material high-target selection and secant initialisation in working return path, preserve accumulated effort and half/uncertainty/direction guards, verify actual material-history acceptance and fresh cold. Do not start947 before resume.')
data=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for p in (statefile,here/'Continuation666_Resume20261123.json'):p.write_bytes(data)
print('CLOSED943-946',len(raw),'RAW;',archivefile.stat().st_size,'BYTES;',aliases,'EXACT DUPLICATES;PAUSE CHECKPOINT PREPARED')
