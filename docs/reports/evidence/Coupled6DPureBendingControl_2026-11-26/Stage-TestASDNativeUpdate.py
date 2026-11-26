"""ДЛЯ ТЕСТОВ. Отделяет однократное материальное обновление ASD от замыкания при тех же таблицах, нагрузках, стали и холодном драйвере."""
from pathlib import Path
from hashlib import sha256
import json,shutil

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1000ASDHardShearCold_20261126';root=base/'Cold1001ASDNativeUpdate_20261126'
assert not root.exists()
proof=json.loads((parent/'ASDHardShearColdProof.json').read_bytes());assert proof['all_six_resultants_and56plus4_equations_verified']
root.mkdir()
for directory in ('src','tests','tools','Input','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
files={
    'docs/regression/Diagnostic.xlsm':(parent/'docs/regression/Diagnostic.xlsm').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'CompiledSourceReuse.json':(json.dumps(dict(parent=str(parent)),indent=2)+'\n').encode(),
    'Guard-IsolatedOffice.ps1':(parent/'Guard-IsolatedOffice.ps1').read_bytes(),
    'Run-TestASDHardShearCold.ps1':(parent/'Run-TestASDHardShearCold.ps1').read_bytes(),
    Path(__file__).name:Path(__file__).read_bytes(),
}
flow=(parent/'src/Materials/CCoupledASDFlow.cls').read_text('utf-8-sig')
flow+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Вычисляет ровно один спектральный материальный шаг без замыкания; состояние вызывающего не меняется.
Friend Function TryNativeStepTest(ByRef strain() As Double, ByRef oldStrain() As Double, ByRef oldEffective() As Double, _
        ByVal oldT As Double, ByVal oldC As Double, ByRef newT As Double, ByRef newC As Double, _
        ByRef effective() As Double, ByRef nominal() As Double, ByRef plastic() As Double, _
        ByRef damageT As Double, ByRef damageC As Double) As Boolean
    On Error GoTo ArithmeticFailed
    SingleStep strain, oldStrain, oldEffective, oldT, oldC, newT, newC, effective, nominal, plastic, damageT, damageC
    TryNativeStepTest = True
    Exit Function
ArithmeticFailed:
    If Err.Number <> 6 And Err.Number <> 11 Then Err.Raise Err.Number, Err.Source, Err.Description
End Function
'''
response=(parent/'src/Materials/CCoupledConcreteResponse.cls').read_text('utf-8-sig')
anchor='Private mASDFlow As CCoupledASDFlow';assert response.count(anchor)==1
response=response.replace(anchor,anchor+'\nPrivate mNativeASDUpdateTest As Boolean \' Только отдельная исследовательская проба однократного обновления.',1)
anchor='    Set mASDFlow = flow';assert response.count(anchor)==1
response=response.replace(anchor,anchor+'\n    mNativeASDUpdateTest = False',1)
anchor='    If mASDFlow Is Nothing Then LawIdentifier = "Barcelona-CDP" Else LawIdentifier = mASDFlow.LawIdentifier'
assert response.count(anchor)==1
response=response.replace(anchor,anchor+'\n    If mNativeASDUpdateTest Then LawIdentifier = "ASD-NativeStep-Cdf0-Test"',1)
anchor='    Dim priorStrain(1 To 6) As Double, priorEffective(1 To 6) As Double, effective(1 To 6) As Double, i As Long'
assert response.count(anchor)==1
response=response.replace(anchor,'    If mNativeASDUpdateTest Then Exit Function \' Замкнутая аналитическая касательная не относится к однократному тестовому обновлению.\n'+anchor,1)
old='''    If Not mASDFlow.TryStep(strain, priorStrain, priorEffective, committed.Tension, committed.Compression, _
            updated.Tension, updated.Compression, effective, apparent, plastic, updated.TensionDamage, updated.CompressionDamage) Then Exit Function'''
assert response.count(old)==1
new='''    If mNativeASDUpdateTest Then
        If Not mASDFlow.TryNativeStepTest(strain, priorStrain, priorEffective, committed.Tension, committed.Compression, _
                updated.Tension, updated.Compression, effective, apparent, plastic, updated.TensionDamage, updated.CompressionDamage) Then Exit Function
    Else
'''+old+'''
    End If'''
response=response.replace(old,new,1)
response+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подключает однократное обновление только для сравнения с внешним материальным алгоритмом; рабочая инициализация сохраняется.
Friend Sub InitializeASDNativeUpdateTest(ByVal pointCount As Long, ByVal flow As CCoupledASDFlow)
    InitializeASD pointCount, flow
    mNativeASDUpdateTest = True
End Sub
'''
test=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
start=test.index('Public Function TestRunASDHardShearCold(');end=test.index('Public Function TestStartFixedForceComparison(',start)
body=test[start:end];assert body.count('.InitializeASD operators.PointCount, flow')==3
body=body.replace('.InitializeASD operators.PointCount, flow','.InitializeASDNativeUpdateTest operators.PointCount, flow')
body=body.replace('замыкание материального шага НДМ сохраняется явно.','используется однократное тестовое материальное обновление.')
body=body.replace('closed_ASD_integration_not_identical_OpenSees_single_update','native_ASD_single_update_test_not_closed')
test=test[:start]+body+test[end:]
for name,code in [('src/Materials/CCoupledASDFlow.cls',flow),('src/Materials/CCoupledConcreteResponse.cls',response),('tests/modTestCoupledElasticState.bas',test)]:
    assert max(map(len,code.splitlines()))<1024
    files[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
compiler=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledASDFlow','CCoupledConcreteResponse','modTestCoupledElasticState')")
files['Compile-TestExternalPoint.ps1']=compiler.encode('utf-8-sig')
contract=json.loads((parent/'Input/MaterialContract.json').read_bytes())
contract.update(integration_mode='native_single_update_test',integration_difference='Single spectral update is tested; material-tensor agreement with OpenSees is not assumed from naming',
                analytic_closed_tangent_disabled=True,original_finite_difference_tangent_fallback=True)
files['Input/MaterialContract.json']=(json.dumps(contract,indent=2,sort_keys=True)+'\n').encode()
for name,data in files.items():(root/name).write_bytes(data)
prep=json.loads((parent/'Preparation.json').read_bytes())
prep['inputs']={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
                if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
prep['directed_scope'].update(case='ASDNativeSingleUpdateCold',actual_import_this_pass=['CCoupledASDFlow','CCoupledConcreteResponse','modTestCoupledElasticState'],
    all256_numerical_components_exact988=False,other254_actual_components_exact1000=True,
    test_only_native_update=True,working_closed_update_preserved=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('PREPARED1001 NATIVE ASD UPDATE TEST;CLOSED WORKING UPDATE RETAINED;SAME COLD DRIVER/TABLES/STEEL/TOLERANCE')
