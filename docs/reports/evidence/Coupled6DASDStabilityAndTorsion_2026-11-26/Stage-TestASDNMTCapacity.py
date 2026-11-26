"""ДЛЯ ТЕСТОВ. Проверяет совместное N+Mx+My+T с точной записью постоянной и переменной частей LoadPath."""
from pathlib import Path
from hashlib import sha256
import json,shutil,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1035ASDPureTorsion_20261126';root=base/'Cold1037ASDNMTCapacity_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
root.mkdir()
for directory in ('src','tests','tools','workbook','Input'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestASDBendingMainCapacity.ps1'):shutil.copyfile(parent/name,root/name)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'ParentCodeUpdate.json')
name='tests/modTestSelectableMaterialCore.bas';code=(root/name).read_text('utf-8-sig')
code=code.replace('Private mMainCapacityBaseTest As Double','Private mMainCapacityLoadTest(1 To 6) As Double')
code=code.replace('Private mMainCapacityComponentTest As Long','Private mMainCapacityPathTest As Long')
code=code.replace('TestRunASDBendingMainCapacity','TestRunASDSpatialMainCapacity')
start=code.index('    Dim testInput As Integer, testCount As Long: testInput = FreeFile')
end=code.index('    Dim geometry As CGeometryRegion',start)
code=code[:start]+'''    Dim testInput As Integer, testCount As Long, testIndex As Long: testInput = FreeFile
    Open folder & "\\Input\\CapacityLoad.bin" For Binary Access Read As #testInput
    Get #testInput, , testCount
    If testCount <> 6 Then Err.Raise 5, , "Для контроля передайте шесть внутренних усилий."
    For testIndex = 1 To 6: Get #testInput, , mMainCapacityLoadTest(testIndex): Next testIndex
    Get #testInput, , mMainCapacityPathTest: Close #testInput
'''+code[end:]
start=code.index('    loads.Initialize 0#, IIf(mMainCapacityComponentTest = 2,')
end=code.index('    Dim capacity As CCapacityResult',start)
code=code[:start]+'''    loads.Initialize mMainCapacityLoadTest(1), mMainCapacityLoadTest(2), mMainCapacityLoadTest(3), 0#, 0#
    loads.SetSpatialComponents mMainCapacityLoadTest(4), mMainCapacityLoadTest(5), mMainCapacityLoadTest(6)
    Dim pathKey As String, pathNames As Variant
    pathNames = Array("Auto", "N", "Mx", "My", "T", "Qx", "Qy", "All6")
    If mMainCapacityPathTest < 0 Or mMainCapacityPathTest > 7 Then Err.Raise 5, , "Неизвестная траектория контрольного расчёта."
    If mMainCapacityPathTest = 0 Then pathKey = "Auto" Else pathKey = ChrW$(&H3BB) & "*" & pathNames(mMainCapacityPathTest)
'''+code[end:]
code=code.replace("' Вызывает штатный SolveCapacity рабочего ASD на чистом изгибе; исход поиска сохраняется отдельно от успешного исполнения теста.",
    "' Прослеживает шесть усилий по выбранной траектории; успешное исполнение макроса не заменяет физическую проверку.")
needle='Public Sub TestCaptureMainCapacity(ByVal controlled As CCoupledDeformationContinuation, ByVal operators As CCoupledElasticState)'
assert code.count(needle)==1
code=code.replace(needle,needle[:-1]+', ByVal loadPath As CLoadPathVector)',1)
code=code.replace('magic = "NDMMAINHIST00002"','magic = "NDMMAINHIST00003"',1)
needle='    TestWriteVector handle, Array(operators.DegreeCount, operators.PointCount, operators.ConcretePointCount, area, length, 0.001, 0.0000005, 0.0000005, mMainCapacityBaseTest, mMainCapacityComponentTest)'
assert code.count(needle)==1
code=code.replace(needle,'''    TestWriteVector handle, Array(operators.DegreeCount, operators.PointCount, operators.ConcretePointCount, area, length, 0.001, 0.0000005, 0.0000005)
    Dim offset() As Double, direction() As Double: loadPath.Components offset, direction
    TestWriteVector handle, offset: TestWriteVector handle, direction''',1)
assert 'mMainCapacityBaseTest' not in code and 'mMainCapacityComponentTest' not in code and max(map(len,code.splitlines()))<1024
(root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
name='src/Solver/CCoupledSectionAnalysis.cls';code=(root/name).read_text('utf-8-sig')
needle='modTestSelectableMaterialCore.TestCaptureMainCapacity controlled, mOperators';assert code.count(needle)==1
code=code.replace(needle,needle+', path',1)
(root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
runner=(root/'Run-TestASDBendingMainCapacity.ps1').read_text('utf-8-sig').replace('modTestSelectableMaterialCore.TestRunASDBendingMainCapacity','modTestSelectableMaterialCore.TestRunASDSpatialMainCapacity')
(root/'Run-TestASDSpatialMainCapacity.ps1').write_bytes(runner.replace('\n','\r\n').encode('utf-8-sig'))
compiler=(parent/'Compile-TestASDPureTorsion.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestSelectableMaterialCore'","-ModuleNames @('CCoupledSectionAnalysis','modTestSelectableMaterialCore')")
(root/'Compile-TestASDNMTCapacity.ps1').write_bytes(compiler.replace('\n','\r\n').encode('utf-8-sig'))
load=[28000.,510000.,180000.,1500000.,0.,0.]
(root/'Input/CapacityLoad.bin').write_bytes(struct.pack('<i6di',6,*load,7))
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='ASD_NMT_explicit_All6',target_N_Nmm=load,positive_N_means='tension',load_path='lambda*All6',
    maximum_lambda=100.,observation_hook_only_in_analysis=True,no_ready_equilibrium_input=True,
    actual_import=['CCoupledSectionAnalysis','modTestSelectableMaterialCore'],all_original_materials_and_tolerances_unchanged=True),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1037; N28KN TENSION/MX0.51/MY0.18/T1.5KNM;EXPLICIT ALL6;SAME MATERIALS AND TOLERANCES')
