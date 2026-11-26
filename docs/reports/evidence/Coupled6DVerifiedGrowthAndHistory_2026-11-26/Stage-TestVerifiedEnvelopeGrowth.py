"""ДЛЯ ТЕСТОВ. Ускоряет восстановление H-шага армированной ветви только с обязательной строгой полной/половинной проверкой увеличенной пробы."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold881SpatialSteelLimits_20261126';root=base/'Cold887VerifiedEnvelopeGrowth_20261126'
proof=json.loads((base/'Cold886Cold56LargerHPairs_20261126/Cold56LargerHPairsProof.json').read_bytes())
assert [(r['ratio'],r['solved']) for r in proof['cases']]==[(1,True),(2,True),(4,True),(8,True),(16,False),(32,False)]
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'VerifiedGrowthCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
def replace(code,old,new):
    """Проверяет точный контракт общего контроллера, исключая замену в другой ветви."""
    assert code.count(old)==1,(old[:100],code.count(old));return code.replace(old,new,1)
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
code=replace(code,'Private Const MATERIAL_STEP_GROWTH As Double = 1.25','''Private Const MATERIAL_STEP_GROWTH As Double = 1.25
Private Const VERIFIED_ENVELOPE_STEP_GROWTH As Double = 2# ' Увеличенная H-проба армированной ветви обязана пройти полную/половинную проверку, иначе общий цикл уменьшает шаг.''')
anchor='Private mEnvelopePrevious As Double, mEnvelopeCurrent As Double, mEnvelopeLastIncrement As Double'
code=replace(code,anchor,anchor+'\nPrivate mVerifiedEnvelopeGrowth As Long, mEnvelopeGrowthAttempts As Long')
anchor='    mEnvelopePrevious = 0#: mEnvelopeCurrent = 0#: mEnvelopeLastIncrement = 0#: mEnvelopeHistoryIndex = 0'
code=replace(code,anchor,anchor+'\n    mVerifiedEnvelopeGrowth = 0: mEnvelopeGrowthAttempts = 0')
code=replace(code,'            stepCeiling = nominal\n            If IsMaterialGrowthControl() Or mLocalPeak > 0# Then',
'''            Dim materialGrowth As Double: materialGrowth = MATERIAL_STEP_GROWTH
            If mUsesEnvelopeControl And reinforced Then materialGrowth = VERIFIED_ENVELOPE_STEP_GROWTH
            stepCeiling = nominal
            If IsMaterialGrowthControl() Or mLocalPeak > 0# Then''')
code=replace(code,'                    stepCeiling = MATERIAL_STEP_GROWTH * nominal','                    stepCeiling = materialGrowth * nominal')
code=replace(code,'                increment = GeomMin(stepCeiling, increment * MATERIAL_STEP_GROWTH)',
    '                increment = GeomMin(stepCeiling, increment * materialGrowth)')
anchor='    ProbeEnvelope = mSolver.TryEnvelopeStep(mOffset, mDirection, previous, acceptedLambda, seed, initialLambda, _'
code=replace(code,anchor,'''    Dim verifyGrowth As Boolean
    If Not mReinforcement Is Nothing Then
        If mReinforcement.HasSteel And after > before Then
            verifyGrowth = (point.Coordinate - after > MATERIAL_STEP_GROWTH * (after - before) + 4# * mEnvelopeResolution)
        End If
    End If
    If verifyGrowth Then
        If mProbeCount >= PROBE_LIMIT Then
            mProbeDetails = "Не осталось второй пробы общего бюджета для половинной проверки увеличенного H-шага."
            Exit Function
        End If
        mEnvelopeGrowthAttempts = mEnvelopeGrowthAttempts + 1
    End If
'''+anchor)
code=replace(code,'        point.Coordinate, point.Lambda, projection, mAbsTol, mRelTol, mMaxIterations, probes, True)',
    '        point.Coordinate, point.Lambda, projection, mAbsTol, mRelTol, mMaxIterations, probes, True, verifyGrowth)')
code=replace(code,'    If ProbeEnvelope Then\n        mProjection = projection',
    '    If ProbeEnvelope Then\n        If verifyGrowth Then mVerifiedEnvelopeGrowth = mVerifiedEnvelopeGrowth + 1\n        mProjection = projection')
anchor='    If Len(mSteelEventDetails) > 0 Then Details = Details & " | " & mSteelEventDetails'
code=replace(code,anchor,'''    If mEnvelopeGrowthAttempts > 0 Then Details = Details & "; увеличенных H-шагов с проверенной половиной: " & _
        CStr(mVerifiedEnvelopeGrowth) & " из " & CStr(mEnvelopeGrowthAttempts)
'''+anchor)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledDeformationContinuation'",1)
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
              'SourceLargerHPairProof.json':(base/'Cold886Cold56LargerHPairs_20261126/Cold56LargerHPairsProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='VerifiedEnvelopeGrowth',actual_import_this_pass=['CCoupledDeformationContinuation'],actual_component_count=256,
    reinforced_H_growth_proposal=2.,other_controls_and_concrete_only_growth_unchanged=True,
    amplified_H_probes_require_existing_full_half_gate=True,extra_half_counted_in_global96=True,
    material_limits_force_tolerances_and_primary30_unchanged=True,physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'H GROWTH PROPOSAL2 WITH MANDATORY PAIR;OTHER CONTROLS1.25;SAME96/30/MATERIAL/FORCE')
