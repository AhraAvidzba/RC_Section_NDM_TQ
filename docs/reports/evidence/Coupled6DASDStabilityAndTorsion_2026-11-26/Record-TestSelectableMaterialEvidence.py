"""ДЛЯ ТЕСТОВ. Архивирует свежую сборку CDP/ASD, направленные проверки и незавершённую визуальную приёмку без общего PASS."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from datetime import datetime
import json,tarfile,re,shutil,ctypes,subprocess
from openpyxl import load_workbook

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={1017:base/'Cold1017SelectableMaterialBook_20261126',1018:base/'Cold1018SelectableMaterialCore_20261126'}
out=repo/'docs/reports/evidence/Coupled6DSelectableMaterials_2026-11-26';assert not out.exists()
fresh=json.loads((roots[1017]/'FreshSourceProof.json').read_bytes())
core=json.loads((roots[1018]/'SelectableMaterialCoreProof.json').read_bytes())
replay=json.loads((roots[1018]/'Native/OpenSeesSelectableCoreMaterialReplay.json').read_bytes())
regression=json.loads((roots[1018]/'MaterialRegressionProof.json').read_bytes())
validation=json.loads((roots[1018]/'Native/ValidationAPI/ValidationAPIProof.json').read_text('utf-8-sig'))
assert fresh['changed_core_modules_strictly_match_pinned_sources'] and core['all_eight_working_core_cases_completed']
assert replay['material_tensor_match'] and not replay['section_equilibrium_solved_by_OpenSees']
assert regression['native']['passed']==39 and regression['native']['own_process_closed']
assert validation['NativeAPIValuesPassed'] and not validation['NativeDropdownUIPassed']
memory=ctypes.create_string_buffer(64);ctypes.c_ulong.from_buffer(memory).value=64
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(memory)
free_ram=ctypes.c_ulonglong.from_buffer(memory,16).value;free_commit=ctypes.c_ulonglong.from_buffer(memory,32).value
assert free_ram>6500*1024**2 and free_commit>6*1024**3

raw={};book_pins={}
for n,root in roots.items():
    prep=json.loads((root/'Preparation.json').read_bytes());files=set(prep['inputs'])
    files.update(p.relative_to(root).as_posix() for p in root.iterdir() if p.is_file() and p.suffix in ('.json','.log','.ndjson','.ps1','.py','.txt'))
    files.update(p.relative_to(root).as_posix() for p in (root/'Native').rglob('*') if p.is_file())
    if n==1017:
        files.update(p.relative_to(root).as_posix() for p in (root/'workbook/output').iterdir() if p.is_file())
        closure=json.loads((root/'BuildClosureAfterExit.json').read_bytes());assert closure['ExactOwnProcessClosed']
        book_name='workbook/output/SelectableLaws.xlsm';expected=fresh['saved_book_sha256']
    else:
        files.add('docs/regression/CodeUpdate.json');book_name='docs/regression/Diagnostic.xlsm'
        expected=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))['WorkbookSHA256']
    files.add(book_name)
    for name in sorted(files):
        path=root/name;assert path.is_file(),path;data=path.read_bytes()
        if name in prep['inputs']:assert sha256(data).hexdigest().upper()==prep['inputs'][name],(n,name)
        if name=='workbook/output/RC_Section_NDM.xlsm':
            assert data==subprocess.check_output(['git','show','HEAD:workbook/output/RC_Section_NDM.xlsm'])
        raw[f'{n}/{name}']=data
    assert sha256(raw[f'{n}/{book_name}']).hexdigest().upper()==expected
    book_pins[str(n)]=dict(member=f'{n}/{book_name}',sha256=expected,exact_saved_book_included=True)

ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
def TestWorkbookCells(data):
    """Сверяет все заполненные ячейки; раскрывает общие формулы и строки, не сопоставляя пересчитанный кэш формул как пользовательский ввод."""
    book=load_workbook(BytesIO(data),read_only=True,data_only=False)
    try:
        return {sheet.title:[(cell.coordinate,cell.data_type,cell.value)
            for row in sheet for cell in row if cell.value is not None] for sheet in book}
    finally:book.close()

source=raw['1018/docs/regression/Diagnostic.xlsm'];source_cells=TestWorkbookCells(source)
assert source_cells==TestWorkbookCells(raw['1017/workbook/output/SelectableLaws.xlsm'])
clone_checks={}
for name,data in raw.items():
    if name.startswith('1018/Native/') and name.endswith('.xlsm'):
        assert TestWorkbookCells(data)==source_cells,name
        clone_checks[name]='all_populated_cell_values_expanded_formulas_types_and_addresses_equal_fresh_source'
out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};canonical={};links=0
archivefile=out/'RawEvidence.tar.xz'
with tarfile.open(archivefile,'w:xz',preset=3) as archive:
    for name,data in raw.items():
        info=tarfile.TarInfo(name);digest=pins[name]
        if digest in canonical:info.type=tarfile.LNKTYPE;info.linkname=canonical[digest];archive.addfile(info);links+=1
        else:canonical[digest]=name;info.size=len(data);archive.addfile(info,BytesIO(data))
    data=(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode();info=tarfile.TarInfo('RawFiles.json');info.size=len(data);archive.addfile(info,BytesIO(data))
assert archivefile.stat().st_size<100*1024**2
checked={}
with tarfile.open(archivefile,'r:xz') as archive:
    index=json.load(archive.extractfile('RawFiles.json'));assert index==pins
    for name,digest in pins.items():
        actual=archive.extractfile(name).read();assert sha256(actual).hexdigest().upper()==digest,name
        checked[name]=digest
assert checked==pins
helpers=set('''Record-TestSelectableMaterialEvidence.py Stage-TestSelectableMaterialBook.py Check-TestSelectableBookSource.py
Stage-TestSelectableMaterialCore.py Check-TestSelectableCoreSource.py Check-TestSelectableMaterialCore.py
Replay-TestSelectableCoreASD.py Run-TestSelectableMaterialUI.ps1 Run-TestSelectableValidationAPI.ps1
Stage-TestSelectableMaterialRegressions.py Fix-TestSelectableMaterialRegressionHarness.py Fix-TestSelectableCoreHarness.py
TestSelectableMaterialCore.bas.txt'''.split())
while True:
    expanded=set(helpers)
    for name in helpers:
        path=here/name;assert path.is_file(),name
        text=path.read_text('utf-8-sig')
        expanded.update(p.name for p in here.iterdir() if p.is_file() and p.suffix in ('.py','.ps1','.txt') and p.name in text)
    if expanded==helpers:break
    helpers=expanded
for name in sorted(helpers):shutil.copyfile(here/name,out/name)
summary=dict(created_local=datetime.now().astimezone().isoformat(),fresh_source_proof=fresh,working_core_proof=core,
    ASD_material_replay={k:v for k,v in replay.items() if k!='observed_OpenSees_concrete_stresses'},
    nonlinear_CDP_regression=regression,native_API_validation=validation,native_visual_dropdown_acceptance_pending=True,
    visual_environment_failure='GPU suspended with monitor off; native accessibility selection not confirmed; own UI session COM busy.',
    book_hashes=book_pins,clone_cell_checks=clone_checks,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),
    exact_duplicate_hardlinks=links,all_raw_hashes_read_back=True,raw_archive_bytes=archivefile.stat().st_size,
    published_workbook_not_replaced=True,main_not_merged=True,full_acceptance_not_run=True,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8')
files={p.name:sha256(p.read_bytes()).hexdigest().upper() for p in out.iterdir() if p.is_file()}
(out/'Files.json').write_text(json.dumps(files,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('SELECTABLE MATERIAL EVIDENCE ARCHIVED',len(raw),'RAW FILES',archivefile.stat().st_size,'BYTES; VISUAL DROPDOWN ACCEPTANCE PENDING')
