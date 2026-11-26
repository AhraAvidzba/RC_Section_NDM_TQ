"""ДЛЯ ТЕСТОВ. Повторяет управление кривизной для текущего CDP в точной скомпилированной книге ASD-проверки."""
from pathlib import Path
from hashlib import sha256
import json,shutil,struct

base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold1011PureBendingASDCurvature_20261126'
root=base/'Cold1012PureBendingCDPCurvature_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
root.mkdir()
for directory in ('src','tests','tools','Input','workbook','docs/regression'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp'):(root/directory).mkdir()
for name in ('Guard-IsolatedOffice.ps1','Run-TestPureBendingCurvature.ps1'):
    shutil.copyfile(parent/name,root/name)
reference=json.loads((root/'Input/PureBendingReferenceProof.json').read_bytes())
(root/'Input/CurvatureOptions.bin').write_bytes(struct.pack('<i3d',3,1.,reference['Mx_ultimate_Nmm'],reference['deformation_plane'][1]))
reuse=dict(parent=str(parent),actual_compiled_book_bitexact=True,all257_components_bitexact=True,
    changed_only_model_input_from_ASD_to_CDP=True,source_book_sha256=compiled['WorkbookSHA256'])
(root/'CompiledSourceReuse.json').write_text(json.dumps(reuse,indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=json.loads((parent/'Preparation.json').read_bytes());prep['inputs']={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
    if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
prep['directed_scope'].update(case='PureBendingCDPCurvatureControl',actual_import_this_pass=[],material='current CDP',
    exact_compiled_book_reused=True,all257_components_exact1011=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1012 CDP; EXACT COMPILED BOOK1011; MODEL INPUT ONLY')
