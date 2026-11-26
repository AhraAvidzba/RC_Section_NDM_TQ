"""ДЛЯ ТЕСТОВ. Проверяет ранний переход ASD к деформационному продолжению и сохраняет полный принятый путь для независимого контроля."""
from pathlib import Path
from hashlib import sha256
import json,shutil,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1025ASDGeometricDivision_20261126';root=base/'Cold1026ASDEarlyDeformation_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
root.mkdir()
for directory in ('src','tests','tools','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Input','Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestASDBendingMainCapacity.ps1'):shutil.copyfile(parent/name,root/name)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'ParentCodeUpdate.json')
module='tests/modTestSelectableMaterialCore.bas';code=(root/module).read_text('utf-8-sig')
assert code.count('Option Explicit')==1
code=code.replace('Option Explicit','Option Explicit\n\nPrivate mMainCapacityFolderTest As String',1)
needle='    On Error GoTo Failed\n    Dim geometry As CGeometryRegion'
assert code.count(needle)==1
code=code.replace(needle,'    On Error GoTo Failed\n    mMainCapacityFolderTest = folder\n    Dim geometry As CGeometryRegion',1)
code+='\n'+(here/'TestCaptureMainCapacity.bas.txt').read_text('utf-8-sig')+'\n'
legacy=(base/'Cold1016PureBendingSteelLimitFixed_20261126/tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
for name,kind in (('TestWriteHardPreparedOperators','Sub'),('TestHardPreparedVector','Function'),('TestOnsetNumber','Function')):
    found=re.search(r'^Private '+kind+' '+name+r'\b.*?^End '+kind,legacy,re.M|re.S);assert found,name
    code+="\n' -------------------- ДЛЯ ТЕСТОВ --------------------\n"+found.group()+'\n'
(root/module).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
name='src/Solver/CCoupledLoadContinuation.cls';code=(root/name).read_text('utf-8-sig')
needle='Optional ByVal preferDeformationOnStall As Boolean = False)';assert code.count(needle)==1
code=code.replace(needle,'Optional ByVal preferDeformationOnStall As Boolean = False, Optional ByVal preferDeformationAfterInitial As Boolean = False)',1)
needle='        If Not Advance(target) Then Exit Sub';assert code.count(needle)==1
code=code.replace(needle,needle+'''
        If preferDeformationAfterInitial And target < maximumLambda Then
            If SavedFieldCount >= 2 And PreviousAcceptedLambda < mLambda Then
                mMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkCapacity, _
                    "Начальная силовая точка подтверждена; поиск продолжается при управлении деформацией.", ProgressDetails()
                Exit Sub
            End If
        End If''',1)
(root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
name='src/Solver/CCoupledSectionAnalysis.cls';code=(root/name).read_text('utf-8-sig')
needle='preferDeformationOnStall:=True';assert code.count(needle)==1
code=code.replace(needle,needle+', preferDeformationAfterInitial:=mCalibratedASD',1)
needle='            Set acceptedTrace = controlled.SavedTrace';assert code.count(needle)==1
code=code.replace(needle,needle+'\n            modTestSelectableMaterialCore.TestCaptureMainCapacity controlled, mOperators',1)
(root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
compiler=(parent/'Compile-TestASDGeometricContinuation.ps1').read_text('utf-8-sig')
compiler=compiler.replace("-ModuleNames 'CCoupledDeformationContinuation'","-ModuleNames @('CCoupledLoadContinuation','CCoupledSectionAnalysis','modTestSelectableMaterialCore')")
(root/'Compile-TestASDEarlyDeformation.ps1').write_bytes(compiler.replace('\n','\r\n').encode('utf-8-sig'))
contract=json.loads((base/'Cold1018SelectableMaterialCore_20261126/Native/SelectableCoreMaterialReplayInput.json').read_bytes())['material_contract']
(root/'Input/MaterialContract.json').write_text(json.dumps(contract,sort_keys=True,indent=2)+'\n',encoding='utf-8')
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='ASD_early_deformation_and_exact_history_capture',probe_limit=96,
    materials_and_equilibrium_tolerances_unchanged=True,no_future_equilibrium_input=True,
    initial_target_lambda=1,base_Mx_Nmm=1000000,private_observer_does_not_modify_material_history=True),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1026; EARLY ASD GEOMETRIC CONTINUATION;EXACT ACCEPTED FIELDS/OPERATORS;NO FUTURE ROOTS')
