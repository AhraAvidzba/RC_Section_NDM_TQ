"""ДЛЯ ТЕСТОВ. Параметризует исходный момент для проверки, зависит ли автоматический поиск от масштаба lambda, сохраняя физические материалы."""
from pathlib import Path
from hashlib import sha256
import json,shutil,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1026ASDEarlyDeformation_20261126';root=base/'Cold1027ASDCapacityScaling_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
independent=json.loads((parent/'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes());assert independent['all56plus4_equations_and_six_resultants_verified']
root.mkdir()
for directory in ('src','tests','tools','workbook','Input'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestASDBendingMainCapacity.ps1'):shutil.copyfile(parent/name,root/name)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'ParentCodeUpdate.json')
module='tests/modTestSelectableMaterialCore.bas';code=(root/module).read_text('utf-8-sig')
needle='Private mMainCapacityFolderTest As String';assert code.count(needle)==1
code=code.replace(needle,needle+'\nPrivate mMainCapacityBaseTest As Double',1)
needle='    mMainCapacityFolderTest = folder';assert code.count(needle)==1
code=code.replace(needle,needle+'''
    Dim testInput As Integer, testCount As Long: testInput = FreeFile
    Open folder & "\\Input\\CapacityBase.bin" For Binary Access Read As #testInput
    Get #testInput, , testCount: Get #testInput, , mMainCapacityBaseTest: Close #testInput
    If testCount <> 1 Or mMainCapacityBaseTest <= 0# Then Err.Raise 5, , "Нужен положительный исходный момент контрольного сочетания."''',1)
needle='loads.Initialize 0#, 1000000#, 0#, 0#, 0#';assert code.count(needle)==1
code=code.replace(needle,'loads.Initialize 0#, mMainCapacityBaseTest, 0#, 0#, 0#',1)
needle='0.0000005, 0.0000005, 1000000#)';assert code.count(needle)==1
code=code.replace(needle,'0.0000005, 0.0000005, mMainCapacityBaseTest)',1)
(root/module).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
(root/'Input/CapacityBase.bin').write_bytes(struct.pack('<id',1,1000000.))
compiler=(parent/'Compile-TestASDEarlyDeformation.ps1').read_text('utf-8-sig')
compiler=compiler.replace("-ModuleNames @('CCoupledLoadContinuation','CCoupledSectionAnalysis','modTestSelectableMaterialCore')","-ModuleNames 'modTestSelectableMaterialCore'")
(root/'Compile-TestASDCapacityScaling.ps1').write_bytes(compiler.replace('\n','\r\n').encode('utf-8-sig'))
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='ASD_capacity_load_multiplier_scaling',actual_import=['modTestSelectableMaterialCore'],
    base_Mx_Nmm=1000000.,all255_kernel_components_exact1026=True,reference_steel_limit_Mx_kNm=4.657850580869645),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1027; ONLY TEST DRIVER PARAMETERIZED;ALL KERNEL COMPONENTS EXACT1026;COLD REPEAT FIRST')
