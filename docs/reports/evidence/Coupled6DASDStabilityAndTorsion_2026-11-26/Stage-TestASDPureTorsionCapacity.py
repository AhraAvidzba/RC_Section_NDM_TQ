"""ДЛЯ ТЕСТОВ. Проверяет автоматическое арматурное ограничение при чистом кручении тем же ASD-ядром с исходной точностью."""
from pathlib import Path
from hashlib import sha256
import json,shutil,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1030ASDScale1M_20261126';root=base/'Cold1035ASDPureTorsion_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
root.mkdir()
for directory in ('src','tests','tools','workbook','Input'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestASDBendingMainCapacity.ps1'):shutil.copyfile(parent/name,root/name)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'ParentCodeUpdate.json')
name='tests/modTestSelectableMaterialCore.bas';code=(root/name).read_text('utf-8-sig')
needle='Private mMainCapacityBaseTest As Double';assert code.count(needle)==1
code=code.replace(needle,needle+'\nPrivate mMainCapacityComponentTest As Long',1)
needle='    If testCount <> 1 Or mMainCapacityBaseTest <= 0# Then Err.Raise 5, , "Нужен положительный исходный момент контрольного сочетания."';assert code.count(needle)==1
code=code.replace(needle,needle+'''
    testInput = FreeFile
    Open folder & "\\Input\\CapacityComponent.bin" For Binary Access Read As #testInput
    Get #testInput, , testCount: Get #testInput, , mMainCapacityComponentTest: Close #testInput
    If testCount <> 1 Or (mMainCapacityComponentTest <> 2 And mMainCapacityComponentTest <> 4) Then Err.Raise 5, , "Для контроля выберите компоненту Mx или T."''',1)
needle='loads.Initialize 0#, mMainCapacityBaseTest, 0#, 0#, 0#: loads.SetSpatialComponents 0#, 0#, 0#';assert code.count(needle)==1
code=code.replace(needle,'''loads.Initialize 0#, IIf(mMainCapacityComponentTest = 2, mMainCapacityBaseTest, 0#), 0#, 0#, 0#
    loads.SetSpatialComponents IIf(mMainCapacityComponentTest = 4, mMainCapacityBaseTest, 0#), 0#, 0#
    Dim pathKey As String
    If mMainCapacityComponentTest = 2 Then pathKey = ChrW$(&H3BB) & "*Mx" Else pathKey = ChrW$(&H3BB) & "*T"''',1)
needle='analysis.SolveCapacity(loads, ChrW$(&H3BB) & "*Mx", spec, 1#, 10#, 0.001, 30, trace)';assert code.count(needle)==1
code=code.replace(needle,'analysis.SolveCapacity(loads, pathKey, spec, 1#, 100#, 0.001, 30, trace)',1)
needle='magic = "NDMMAINHIST00001"';assert code.count(needle)==1;code=code.replace(needle,'magic = "NDMMAINHIST00002"',1)
needle='0.001, 0.0000005, 0.0000005, mMainCapacityBaseTest)';assert code.count(needle)==1
code=code.replace(needle,'0.001, 0.0000005, 0.0000005, mMainCapacityBaseTest, mMainCapacityComponentTest)',1)
(root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
(root/'Input/CapacityBase.bin').write_bytes(struct.pack('<id',1,100000.))
(root/'Input/CapacityComponent.bin').write_bytes(struct.pack('<ii',1,4))
compiler=(base/'Cold1027ASDCapacityScaling_20261126/Compile-TestASDCapacityScaling.ps1').read_bytes()
(root/'Compile-TestASDPureTorsion.ps1').write_bytes(compiler)
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='ASD_pure_torsion_capacity',base_T_Nmm=100000.,load_path='lambda*T',maximum_lambda=100.,
    actual_import=['modTestSelectableMaterialCore'],all255_kernel_components_exact_parent=True,
    original_absolute_tolerance_MPa=.0000005,original_relative_tolerance=.0000005,physical_capacity_not_assumed=True),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1035; PURE TORSION LAMBDA*T;BASE0.1KNM;ORIGINAL MATERIALS/TOLERANCES;KERNEL UNCHANGED')
