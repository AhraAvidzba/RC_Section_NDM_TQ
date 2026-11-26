"""ДЛЯ ТЕСТОВ. Архивирует положительные и отрицательные внешние контроли и точные тестовые книги; численную остановку не выдаёт за разрушение."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from runpy import run_path
from datetime import datetime
import json,tarfile,re,ctypes

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={988:'Cold988FixedForceExternal_20261126',989:'Cold989FixedForceAbove_20261126',
       990:'Cold990FixedForcePeakFromPrior_20261126',991:'Cold991FixedForcePeakHybr_20261126',
       1000:'Cold1000ASDHardShearCold_20261126',1001:'Cold1001ASDNativeUpdate_20261126',1002:'Cold1002ASDExactFloors_20261126'}
roots={n:base/name for n,name in names.items()}
keys={n:('FixedForceExternalProof.json' if n<1000 else 'ASDHardShearColdProof.json') for n in names}
proofs={str(n):json.loads((root/keys[n]).read_bytes()) for n,root in roots.items()}
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statepath.read_bytes())
assert not state['active_jobs'] and state['active_exec_session'] is None and state['goal_status']=='active'
class TestMemory(ctypes.Structure):
    """ДЛЯ ТЕСТОВ. Проверяет запас перед упаковкой больших двоичных доказательств."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory)
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
for n in names:
    proof=proofs[str(n)];assert not proof['physical_capacity_certified'] and proof['source_unchanged']
    if n<1000:assert proof['own_excel_closed'] and proof['accepted_history_unchanged_bitexact']
    else:assert proof['native']['own_process_closed'] and proof['all_six_resultants_and56plus4_equations_verified']
parent_book=(base/'Cold954RefinedEventApproach_20261126/docs/regression/Diagnostic.xlsm').read_bytes()
parent_sha=sha256(parent_book).hexdigest().upper()
assert parent_sha=='FB807B2694BC0334D24A57B32FD97B0CFED81FD0798FA34E8359C762805406E9'
prior=repo/'docs/reports/evidence/Coupled6DDetectedEventApproach_2026-11-26'
prior_index=json.loads((prior/'Files.json').read_bytes())
assert sha256((prior/'RawEvidence.tar.xz').read_bytes()).hexdigest().upper()==prior_index['RawEvidence.tar.xz']
segment=run_path(str(here/'Rebuild-TestBookSegments.py'));raw={};publication={};reconstruction={}
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def TestCellContents(data):
    """Сопоставляет все координаты, типы, значения и формулы ячеек с уже опубликованной книгой, сохраняя буквальные строки."""
    result=[]
    for cell in ET.fromstring(data).findall('.//s:c',ns):
        formula=cell.find('s:f',ns);value=cell.find('s:v',ns);inline=cell.find('s:is',ns)
        result.append((cell.attrib['r'],cell.get('t'),None if formula is None else (formula.text,tuple(sorted(formula.attrib.items()))),
                       None if value is None else value.text,None if inline is None else ''.join(inline.itertext())))
    return result

for n,root in roots.items():
    book=(root/'docs/regression/Diagnostic.xlsm').read_bytes()
    manifest=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    assert sha256(book).hexdigest().upper()==manifest['WorkbookSHA256']
    with ZipFile(BytesIO(parent_book)) as parent,ZipFile(BytesIO(book)) as current:
        assert set(parent.namelist())==set(current.namelist())
        sheets=[name for name in current.namelist() if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',name)]
        assert len(sheets)==5
        for name in sheets:assert TestCellContents(parent.read(name))==TestCellContents(current.read(name)),(n,name)
        for name in current.namelist():
            if name=='xl/sharedStrings.xml' or name.startswith('xl/externalLinks/') or name=='xl/connections.xml':assert parent.read(name)==current.read(name)
    patch,blobs=segment['TestMakeSegments'](parent_book,book)
    for kind in ('parent','segment','final_digest'):
        bad_patch=json.loads(json.dumps(patch));bad_blobs=blobs.copy();bad_parent=parent_book
        if kind=='parent':bad_parent+=b'x'
        if kind=='final_digest':bad_patch['target_sha256']='0'*64
        if kind=='segment':
            item=next(item for item in bad_patch['segments'] if 'blob' in item);bad_blobs[item['blob']]+=b'x'
        try:segment['TestRebuildSegments'](bad_parent,bad_patch,bad_blobs)
        except AssertionError:pass
        else:raise AssertionError('Corrupt book accepted: '+kind)
    for digest,data in blobs.items():raw['BookSegmentBlobs/'+digest]=data
    raw[f'{n}/ActualSavedBookSegments.json']=(json.dumps(patch,indent=2,sort_keys=True)+'\n').encode()
    reconstruction[str(n)]=dict(parent_book_sha256=parent_sha,actual_saved_book_sha256=patch['target_sha256'],
        all_saved_zip_bytes_reconstructed_and_verified=True,saved_book_bytes=len(book),
        parent_archive='docs/reports/evidence/Coupled6DDetectedEventApproach_2026-11-26/RawEvidence.tar.xz',
        parent_archive_sha256=prior_index['RawEvidence.tar.xz'],parent_book_member='954/ActualCompiledWorkbook.xlsm',
        manifest_member=f'{n}/ActualSavedBookSegments.json',blobs_prefix='BookSegmentBlobs/',
        reconstruction_tool='Rebuild-TestBookSegments.py',corrupt_parent_segment_and_final_digest_rejected=True)
    publication[str(n)]=dict(all5_sheet_cell_values_formulas_types_and_addresses_exact_public954=True,
        shared_strings_external_links_and_connections_unchanged=True,no_new_workbook_members=True)
    prep=json.loads((root/'Preparation.json').read_bytes())
    files=set(prep['inputs'])|{'Preparation.json',keys[n],'docs/regression/CodeUpdate.json','PointTraceSourceComposition.json',
        'CompilePoint.log','CompilePointClosure.json','CompileClosure.json','CompiledSourceReuse.json','Run.ndjson'}
    if n==1002:files.discard('CompilePoint.log')
    files.update(p.relative_to(root).as_posix() for p in (root/'Native').rglob('*') if p.is_file())
    for name in sorted(files):
        path=root/name;assert path.is_file(),path
        data=path.read_bytes()
        if name in prep['inputs']:assert sha256(data).hexdigest().upper()==prep['inputs'][name],name
        raw[f'{n}/{name}']=data
out=repo/'docs/reports/evidence/Coupled6DExternalForceAndASD_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};canonical={};aliases=0
archivefile=out/'RawEvidence.tar.xz'
with tarfile.open(archivefile,'w:xz',preset=6) as archive:
    for name,data in raw.items():
        info=tarfile.TarInfo(name);digest=pins[name]
        if digest in canonical:info.type=tarfile.LNKTYPE;info.linkname=canonical[digest];archive.addfile(info);aliases+=1
        else:canonical[digest]=name;info.size=len(data);archive.addfile(info,BytesIO(data))
    data=(json.dumps(pins,indent=2,sort_keys=True)+'\n').encode();info=tarfile.TarInfo('RawFiles.json');info.size=len(data);archive.addfile(info,BytesIO(data))
assert archivefile.stat().st_size<100*1024**2,'Split archive before publication'
checked={};actual_pins=None;parts={}
with tarfile.open(archivefile,'r|xz') as archive:
    for info in archive:
        if info.islnk():
            assert info.linkname in checked;digest=checked[info.linkname];data=raw[info.name]
        else:
            data=archive.extractfile(info).read();digest=sha256(data).hexdigest().upper()
            if info.name=='RawFiles.json':actual_pins=json.loads(data);continue
        assert digest==pins[info.name];checked[info.name]=digest
        if info.name.startswith('BookSegmentBlobs/') or info.name.endswith('ActualSavedBookSegments.json'):parts[info.name]=data
assert checked==pins==actual_pins
for n,root in roots.items():
    patch=json.loads(parts[f'{n}/ActualSavedBookSegments.json'])
    blobs={item['blob']:parts['BookSegmentBlobs/'+item['blob']] for item in patch['segments'] if 'blob' in item}
    assert segment['TestRebuildSegments'](parent_book,patch,blobs)==(root/'docs/regression/Diagnostic.xlsm').read_bytes()
helpers=set('''Record-TestExternalComparisons.py Rebuild-TestBookSegments.py Stage-TestFixedForceComparison.py
Stage-TestFixedForceAbove.py Stage-TestFixedForceCase.py Run-TestFixedForceExternal.py Check-TestFixedForceExternal.py
TestFixedForceComparison.bas.txt Stage-TestASDHardShearCold.py TestASDHardShearCold.bas.txt Check-TestASDHardShearCold.py
Stage-TestASDNativeUpdate.py Stage-TestASDExactFloors.py Replay-TestASDNativeMaterial.py Check-TestASDNativeStepSource.py
Finalize-TestExisting257Project.py Set-TestActiveStage.py Set-TestTerminalStage.py'''.split());todo=list(helpers)
while todo:
    path=here/todo.pop();assert path.is_file(),path
    for dep in re.findall(r'[A-Za-z][A-Za-z0-9_-]*\.(?:py|ps1|bas\.txt|cls\.txt)',path.read_text('utf-8-sig')):
        if dep not in helpers and (here/dep).is_file():helpers.add(dep);todo.append(dep)
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
release={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release==json.loads((prior/'Summary.json').read_bytes())['release_hashes_unchanged']
summary=dict(proofs=proofs,book_reconstruction=reconstruction,publication_scope_verification=publication,
    raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),raw_archive_bytes=archivefile.stat().st_size,
    exact_duplicate_hardlinks=aliases,all_logical_raw_hashes_read_back=True,all_seven_actual_books_rebuilt_bitexact=True,
    release_hashes_unchanged=release,all_started_stages_terminal=True,full_acceptance_not_run=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
index={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(index,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
             latest_external_force_and_ASD_evidence=out.relative_to(repo).as_posix(),latest_external_force_and_ASD_proofs=proofs)
data=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for path in (statepath,here/'Continuation666_Resume20261123.json'):path.write_bytes(data)
print('ARCHIVED988-991/1000-1002',len(raw),'RAW;',archivefile.stat().st_size,'BYTES;',aliases,'DUPLICATES;SEVEN EXACT BOOKS REBUILT;GOAL ACTIVE')
