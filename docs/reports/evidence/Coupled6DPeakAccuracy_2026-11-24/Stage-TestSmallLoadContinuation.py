"""ДЛЯ ТЕСТОВ. Разделяет точность максимума и допустимый шаг продолжения; сохраняет все конечные проверки предела."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold713SmallLoadContinuation_20261124';parent=base/'Cold712PriorDilationCapacity_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'SmallLoadCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='src/Solver/CCoupledDeformationContinuation.cls';code=(root/path).read_text('utf-8-sig')
anchor='''        ElseIf ((changedControl Or mUnverifiedDecline) And Not mUsesEnvelopeControl) Or mFollowingLocalPeak Or _
''';assert code.count(anchor)==1
code=code.replace(anchor,'''        ElseIf SelectCumulativePeak(history, trial, left, current) Then
            mSolver.DiscardStep
            If VerifyPeak(history, left, current, trial, reinforced) Then
                If Not reinforced Then Exit Sub
                ResumeAfterLocalPeak history, current, trial
                increment = nominal * 0.25: minimum = increment * MIN_STEP_RATIO
            Else
                StopSearch "Суммарный спад найден, но его максимум не подтверждён повтором плеч."
                Exit Sub
            End If
'''+anchor,1)
anchor='''                (mUsesEnvelopeControl And trial.Lambda >= current.Lambda - LoadResolution(current.Lambda)) Then''';assert code.count(anchor)==1
code=code.replace(anchor,'''                (mUsesEnvelopeControl And trial.Lambda >= current.Lambda - LoadResolution(current.Lambda)) Or _
                AdvancesWithSmallLoadChange(history, current, trial) Then''',1)
anchor='            If accuracy <= tolerance Then Exit For\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''            If mDesignHistoryIndex = 0 Then Exit For ' До подтверждения расчётной нагрузки сохраняется грубый локальный пик; конечная приёмка по-прежнему требует исходной точности.
''',1)
code+='\n'+(here/'CandidateSmallLoadContinuation.cls.txt').read_text('utf-8-sig')
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
anchor='    branch.Trace solver, response(2), offset, direction, history, previousFactor, initialFactor, initialActual, 3#, 0.001, 30, ABS_TOL, REL_TOL, trace\n';assert code.count(anchor)==1
code=code.replace(anchor,'''    If capacityOnly Then
        mStockProbeRootTest = evidenceRoot: Set mStockProbeObserverTest = Nothing
        mStockRawHistoryCountTest = 0: mStockHRequestCountTest = 0
    End If
'''+anchor+'''    mStockProbeRootTest = vbNullString: Set mStockProbeObserverTest = Nothing
''',1)
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledDeformationContinuation','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
for name in (Path(__file__).name,'CandidateSmallLoadContinuation.cls.txt'):values[name]=(here/name).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='SmallLoadSpatialContinuation',actual_import_this_pass=['CCoupledDeformationContinuation','modTestCoupledElasticState'],
    small_lambda_change_requires_verified_forward_spatial_increment=True,cumulative_peak_requires_replayed_shoulders=True,
    early_local_refinement_deferred_until_design_reached=True,final_precision_and_material_checks_unchanged=True,
    original96_probe_limit=True,original30_primary_iterations=True,checkpoint_observer_enabled=True,
    physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('713 SMALL LOAD CONTINUATION; CUMULATIVE PEAK; ORIGINAL FINAL PRECISION AND96 PROBES')
