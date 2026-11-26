"""ДЛЯ ТЕСТОВ. Проверяет ограниченную адаптацию геометрического шага ASD и откладывание локальных пиков без их выдачи за предел."""
from pathlib import Path
from hashlib import sha256
import json,shutil

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1021ASDBendingMainCapacityRetry_20261126';root=base/'Cold1023ASDGeometricContinuation_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
root.mkdir()
for directory in ('src','tests','tools','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1','Run-TestASDBendingMainCapacity.ps1'):shutil.copyfile(parent/name,root/name)
shutil.copyfile(parent/'docs/regression/Diagnostic.xlsm',root/'docs/regression/Diagnostic.xlsm')
shutil.copyfile(parent/'docs/regression/CodeUpdate.json',root/'ParentCodeUpdate.json')
compile_script=(parent/'Compile-TestSelectableMaterialCore.ps1').read_text('utf-8-sig')
compile_script=compile_script.replace("-ModuleNames 'modTestSelectableMaterialCore'","-ModuleNames 'CCoupledDeformationContinuation'")
(root/'Compile-TestASDGeometricContinuation.ps1').write_bytes(compile_script.replace('\n','\r\n').encode('utf-8-sig'))
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
assert 'Private Const PROBE_LIMIT As Long = 96' in code
needle='                ElseIf IsMaterialGrowthControl() Then\n                    smoothTrial = FitsMaterialGrowthSecant(history, left, current, trial)'
assert code.count(needle)==1
code=code.replace(needle,'''                ElseIf UsesASDGeometricControl() Then
                    smoothTrial = (mSolver.IterationCount <= GeomMax(1, mMaxIterations \\ 2)) ' Малое число итераций разрешает рост численного шага; допуск равновесия сохраняется.
                ElseIf IsMaterialGrowthControl() Then
                    smoothTrial = FitsMaterialGrowthSecant(history, left, current, trial)''',1)
needle='If smoothTrial And (mFollowingLocalPeak Or IsMaterialGrowthControl()) Then'
assert code.count(needle)==1
code=code.replace(needle,'If smoothTrial And (mFollowingLocalPeak Or IsMaterialGrowthControl() Or UsesASDGeometricControl()) Then',1)
needle='            If mUsesTransverseControl Then stepCeiling = TRANSVERSE_STEP_RATIO * nominal'
assert code.count(needle)==1
code=code.replace(needle,'''            If UsesASDGeometricControl() And smoothSteps >= 2 Then _
                stepCeiling = GeomMax(nominal, 0.25 * Abs(current.Coordinate)) ' Относительный потолок не зависит от размерности выбранной меры.
'''+needle,1)
needle='            If IsMaterialGrowthControl() And smoothSteps >= 2 Then'
assert code.count(needle)==1
code=code.replace(needle,'            If (IsMaterialGrowthControl() Or UsesASDGeometricControl()) And smoothSteps >= 2 Then',1)
needle='    If Not SelectPeakAnchor(history, left, middle) Then Exit Function'
assert code.count(needle)==1
code=code.replace(needle,'''    If keepRight And UsesASDGeometricControl() Then Exit Function ' Локальный спад железобетона проходит как непроверенный; максимум не объявляется, критерий стали остаётся обязательным.
'''+needle,1)
needle="' Разрешает адаптацию только для той же точной пространственной меры"
assert code.count(needle)==1
index=code.index(needle)
code=code[:index]+'''\
' Разрешает геометрическую адаптацию ASD при неизменной осевой, изгибной или крутильной мере.
' Материальные, сдвиговые и заменённые проекции сохраняют свои контроллеры.
Private Function UsesASDGeometricControl() As Boolean
    UsesASDGeometricControl = mSolver.HasPrincipalBoundaries And Not mUsesEnvelopeControl And _
        Not mUsesYieldControl And Not mUsesTransverseControl And Not mUsesFieldControl
End Function

'''+code[index:]
assert max(map(len,code.splitlines()))<1024
(root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent)),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='ASD_private_geometric_adaptation',probe_limit=96,material_and_equilibrium_criteria_unchanged=True,
    refined_local_peaks_not_certified=True,relative_step_ceiling=.25,reference_ASD_strain_limit_kNm=4.658444129476362),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
        if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1023; ASD GEOMETRIC ADAPTATION;96 PROBES;NO RELAXED EQUILIBRIUM CRITERIA;NO CERTIFIED PEAK FROM UNREFINED DIP')
