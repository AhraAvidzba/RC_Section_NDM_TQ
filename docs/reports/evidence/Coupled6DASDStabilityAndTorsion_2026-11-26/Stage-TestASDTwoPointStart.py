"""ДЛЯ ТЕСТОВ. Передаёт ASD деформационному продолжению после первой пары силовых равновесий, сохраняя проверку масштаба нагрузки."""
from pathlib import Path
from hashlib import sha256
import json,shutil

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1028ASDCapacityDoubleInput_20261126';root=base/'Cold1029ASDTwoPointStart_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
native=(parent/'Native/RealColdReplay.txt').read_text('utf-8-sig');assert 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in native
root.mkdir()
for directory in ('src','tests','tools','workbook','Input'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestASDBendingMainCapacity.ps1'):shutil.copyfile(parent/name,root/name)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'ParentCodeUpdate.json')
name='src/Solver/CCoupledLoadContinuation.cls';code=(root/name).read_text('utf-8-sig')
needle='Private mPreferDeformationOnStall As Boolean';assert code.count(needle)==1
code=code.replace(needle,needle+'\nPrivate mPreferDeformationAfterInitial As Boolean',1)
needle='    mPreferDeformationOnStall = False';assert code.count(needle)==1
code=code.replace(needle,needle+': mPreferDeformationAfterInitial = False',1)
needle='    mPreferDeformationOnStall = preferDeformationOnStall';assert code.count(needle)==1
code=code.replace(needle,needle+': mPreferDeformationAfterInitial = preferDeformationAfterInitial',1)
needle='            reached = candidate\n            triedExactTarget = False';assert code.count(needle)==1
code=code.replace(needle,'''            reached = candidate
            If mPreferDeformationAfterInitial And Not offsetStage Then
                If SavedFieldCount >= 2 And mPreviousLambda < mLambda Then
                    mMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkCapacity, _
                        "Начальная силовая пара подтверждена; поиск продолжается при управлении деформацией.", ProgressDetails()
                    Exit Function
                End If
            End If
            triedExactTarget = False''',1)
(root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
original=base/'Cold1027ASDCapacityScaling_20261126';compiler=(original/'Compile-TestASDCapacityScaling.ps1').read_text('utf-8-sig')
compiler=compiler.replace("-ModuleNames 'modTestSelectableMaterialCore'","-ModuleNames 'CCoupledLoadContinuation'")
(root/'Compile-TestASDTwoPointStart.ps1').write_bytes(compiler.replace('\n','\r\n').encode('utf-8-sig'))
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='ASD_initial_two_point_control_handoff',base_Mx_Nmm=2000000.,
    numerical_material_models_unchanged=True,control_probe_budget=96,all_equilibrium_tolerances_unchanged=True,
    regression_trigger='same physical ray at doubled base moment stalled before steel'),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1029; EARLY HANDOFF AFTER FIRST FORCE PAIR; SAME DOUBLE-BASE REGRESSION')
