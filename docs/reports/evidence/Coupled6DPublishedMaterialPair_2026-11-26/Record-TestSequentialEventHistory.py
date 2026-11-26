"""ДЛЯ ТЕСТОВ. Сохраняет955–957 с точным восстановлением фактических книг из опубликованной954 и проверенных ZIP-сегментов."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
from runpy import run_path
from xml.etree import ElementTree as ET
import json,tarfile,re,ctypes

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={n:base/name for n,name in {955:'Cold955SequentialEventHistory_20261126',956:'Cold956RefinedSequentialHistory_20261126',957:'Cold957FourPartEventHistory_20261126'}.items()}
keys={955:'SequentialEventHistoryProof.json',956:'RefinedSequentialHistoryProof.json',957:'FourPartEventHistoryProof.json'}
proofs={str(n):json.loads((root/keys[n]).read_bytes()) for n,root in roots.items()}
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statepath.read_bytes())
assert not state['active_jobs'] and state['active_exec_session'] is None and state['goal_status']=='active'
class TestMemory(ctypes.Structure):
    """Ограничивает упаковку при недостаточном запасе физической и общей выделяемой памяти."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory)
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
for proof in proofs.values():
    assert proof['native']['own_process_closed'] and proof['native']['failed']==0 and not proof['physical_capacity_certified']
    assert proof['original25_field_concrete_state_energy_stress_and_steel_stress_restored_bitexact_twice']
parent_book=(base/'Cold954RefinedEventApproach_20261126/docs/regression/Diagnostic.xlsm').read_bytes()
parent_sha=sha256(parent_book).hexdigest().upper()
assert parent_sha=='FB807B2694BC0334D24A57B32FD97B0CFED81FD0798FA34E8359C762805406E9'
prior=repo/'docs/reports/evidence/Coupled6DDetectedEventApproach_2026-11-26'
prior_index=json.loads((prior/'Files.json').read_bytes())
assert sha256((prior/'RawEvidence.tar.xz').read_bytes()).hexdigest().upper()==prior_index['RawEvidence.tar.xz']
segment=run_path(str(here/'Rebuild-TestBookSegments.py'));raw={};publication={};reconstruction={}
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def TestCellContents(data):
    """Сравнивает все фактические значения, формулы, типы и координаты ячеек; оформление не маскируется как пользовательские данные."""
    items=[]
    for cell in ET.fromstring(data).findall('.//s:c',ns):
        formula=cell.find('s:f',ns);value=cell.find('s:v',ns);inline=cell.find('s:is',ns)
        items.append((cell.attrib['r'],cell.get('t'),None if formula is None else (formula.text,tuple(sorted(formula.attrib.items()))),
            None if value is None else value.text,None if inline is None else ''.join(inline.itertext())))
    return items

fixed=['Preparation.json','docs/regression/CodeUpdate.json','CompiledSourceReuse.json','PointTraceSourceComposition.json',
    'CompilePoint.log','CompilePointClosure.json','CompileClosure.json','Run.ndjson','ParentCodeUpdate.json','ParentCompileClosure.json',
    'Run-TestRealColdReplay.ps1','Guard-IsolatedOffice.ps1','Compile-TestExternalPoint.ps1',
    'tools/build_workbook/Run-Audit02ModeTests.ps1','tools/build_workbook/Run-FocusedKaniWatchdog.ps1',
    'tools/build_workbook/Update-IsolatedOptimizerCode.ps1','SourceDetectedEventProof.json']
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
        raw[f'{n}/ActualCompiledVBA.bin']=current.read('xl/vbaProject.bin')
    assert (root/'Input/Reverse1.bin').read_bytes()==(base/'Cold954RefinedEventApproach_20261126/Input/Reverse1.bin').read_bytes()
    patch,blobs=segment['TestMakeSegments'](parent_book,book)
    for bad_case in ('parent','segment','final_digest'):
        bad_patch=json.loads(json.dumps(patch));bad_blobs=blobs.copy();bad_parent=parent_book
        if bad_case=='parent':bad_parent=parent_book+b'x'
        if bad_case=='final_digest':bad_patch['target_sha256']='0'*64
        if bad_case=='segment':
            item=next(item for item in bad_patch['segments'] if 'blob' in item)
            bad_blobs[item['blob']]=blobs[item['blob']]+b'x'
        try:segment['TestRebuildSegments'](bad_parent,bad_patch,bad_blobs)
        except AssertionError:pass
        else:raise AssertionError('Повреждение книги не отклонено: '+bad_case)
    for digest,data in blobs.items():raw['BookSegmentBlobs/'+digest]=data
    raw[f'{n}/ActualSavedBookSegments.json']=(json.dumps(patch,sort_keys=True,indent=2)+'\n').encode()
    reconstruction[str(n)]=dict(parent_book_sha256=parent_sha,actual_saved_book_sha256=patch['target_sha256'],
        all_saved_zip_bytes_reconstructed_and_verified=True,new_segment_bytes=sum(map(len,blobs.values())),
        saved_book_bytes=len(book),parent_archive='docs/reports/evidence/Coupled6DDetectedEventApproach_2026-11-26/RawEvidence.tar.xz',
        parent_archive_sha256=prior_index['RawEvidence.tar.xz'],parent_book_member='954/ActualCompiledWorkbook.xlsm',
        manifest_member=f'{n}/ActualSavedBookSegments.json',blobs_prefix='BookSegmentBlobs/',
        reconstruction_tool='Rebuild-TestBookSegments.py',corrupt_parent_segment_and_final_digest_all_rejected=True)
    publication[str(n)]=dict(all5_sheet_cell_values_formulas_types_and_addresses_exact_public954=True,
        shared_strings_external_links_and_connections_unchanged=True,geometry_materials_and_original25_input_bitexact_public954=True,
        no_new_workbook_members=True)
    files=fixed+[keys[n]]
    for folder in ('Native','Input','src','tests','SequentialEventCompilationAncestor','RefinedSequentialCompilationAncestor','FourPartCompilationAncestor'):
        files += [p.relative_to(root).as_posix() for p in (root/folder).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{n}/{name}']=(root/name).read_bytes()
out=repo/'docs/reports/evidence/Coupled6DSequentialEventHistory_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};canonical={};aliases=0
archivefile=out/'RawEvidence.tar.xz'
with tarfile.open(archivefile,'w:xz',preset=6) as archive:
    for name,data in raw.items():
        info=tarfile.TarInfo(name);digest=pins[name]
        if digest in canonical:info.type=tarfile.LNKTYPE;info.linkname=canonical[digest];archive.addfile(info);aliases+=1
        else:canonical[digest]=name;info.size=len(data);archive.addfile(info,BytesIO(data))
    data=(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode();info=tarfile.TarInfo('RawFiles.json');info.size=len(data);archive.addfile(info,BytesIO(data))
checked={};actual_pins=None;actual_parts={}
with tarfile.open(archivefile,'r|xz') as archive:
    for info in archive:
        if info.islnk():assert info.linkname in checked;digest=checked[info.linkname]
        else:
            data=archive.extractfile(info).read();digest=sha256(data).hexdigest().upper()
            if info.name=='RawFiles.json':actual_pins=json.loads(data);continue
            if info.name.startswith('BookSegmentBlobs/') or info.name.endswith('ActualSavedBookSegments.json'):actual_parts[info.name]=data
        assert digest==pins[info.name];checked[info.name]=digest
        if info.islnk() and (info.name.startswith('BookSegmentBlobs/') or info.name.endswith('ActualSavedBookSegments.json')):
            actual_parts[info.name]=raw[info.name]
assert checked==pins==actual_pins
for n,root in roots.items():
    patch=json.loads(actual_parts[f'{n}/ActualSavedBookSegments.json'])
    blobs={item['blob']:actual_parts['BookSegmentBlobs/'+item['blob']] for item in patch['segments'] if 'blob' in item}
    restored=segment['TestRebuildSegments'](parent_book,patch,blobs)
    assert restored==(root/'docs/regression/Diagnostic.xlsm').read_bytes()
helpers=set('''Stage-TestSequentialEventHistory.py Check-TestSequentialEventHistory.py TestRunSequentialEventHistory.bas.txt
Stage-TestRefinedSequentialHistory.py Check-TestRefinedSequentialHistory.py Stage-TestFourPartEventHistory.py
Check-TestFourPartEventHistory.py Rebuild-TestBookSegments.py Finalize-TestExisting257Project.py
Set-TestActiveStage.py Set-TestTerminalStage.py'''.split());helpers.add(Path(__file__).name);todo=list(helpers)
while todo:
    p=here/todo.pop();assert p.is_file()
    for dep in re.findall(r'[A-Za-z][A-Za-z0-9_-]*\.(?:py|ps1|bas\.txt|cls\.txt)',p.read_text('utf-8-sig')):
        if dep not in helpers and (here/dep).is_file():helpers.add(dep);todo.append(dep)
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
release={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release==json.loads((prior/'Summary.json').read_bytes())['release_hashes_unchanged']
summary=dict(proofs=proofs,book_reconstruction=reconstruction,publication_scope_verification=publication,
    raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),raw_archive_bytes=archivefile.stat().st_size,
    exact_duplicate_hardlinks=aliases,all_logical_raw_hashes_read_back=True,all3_saved_books_rebuilt_bitexact_from_archived_segments=True,
    release_hashes_unchanged=release,all_numerical_stages_terminal=True,full_acceptance_not_run=True,
    physical_capacity_certified=False,cold_gain_claimed=False,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
index={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(index,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    latest_sequential_history_evidence=out.relative_to(repo).as_posix(),latest_sequential_history955=proofs['955'],
    latest_refined_sequential956=proofs['956'],latest_four_part957=proofs['957'],
    previous_goal_turn_classification='progress:955 real two-substep history endpoint verified and original25 restored twice;956 tighter second equilibrium leaves2.022% endpoint defect and4.12e-6 lambda difference;957 four-step refinement independently verified and archived. No geometric/force/steel guards relaxed, production acceptance or physical capacity claimed.')
data=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for target in (statepath,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED955-957',len(raw),'RAW;',archivefile.stat().st_size,'BYTES;',aliases,'EXACT DUPLICATES;ALL3 ACTUAL BOOKS REBUILT BITEXACT;GOAL ACTIVE')
