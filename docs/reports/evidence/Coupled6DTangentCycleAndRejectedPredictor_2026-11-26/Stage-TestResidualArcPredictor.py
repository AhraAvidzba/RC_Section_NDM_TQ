"""ДЛЯ ТЕСТОВ. Выбирает прогноз дуги по фактической невязке принятой секущей и касательной; полную/половинную приёмку сохраняет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import ast,json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold901ConsistentControlRefinement_20261126';source=base/'Cold900Cold32OutgoingJacobian_20261126'
root=base/'Cold908ResidualArcPredictor_20261126'
assert json.loads((base/'Cold907Cold32TangentConeIterations_20261126/Cold32TangentConeIterationsProof.json').read_bytes())['all12_available']
assert json.loads((base/'Cold905ConsistentRefinementColdStress_20261126/VerifiedColdStressBridgeProof.json').read_bytes())['entire_raw_trajectory_bit_identical897_and898']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ResidualPredictorCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
code+='\n\n'+(here/'CandidateArcPredictorDefects.cls.txt').read_text('utf-8-sig')
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
def TestReplace(old,new):
    """Проверяет точный исходный контракт перед заменой одного участка общего класса."""
    global code
    assert code.count(old)==1,old[:100];code=code.replace(old,new,1)
TestReplace('Private Const FIELD_ERROR_LIMIT As Double = 0.005','Private Const FIELD_ERROR_LIMIT As Double = 0.005\nPrivate Const PREDICTOR_PROBE_RATIO As Double = 0.03125\nPrivate Const PREDICTOR_IMPROVEMENT_RATIO As Double = 0.25')
TestReplace('Private mUsesTangent As Boolean, mPredictionVerification As Variant','Private mUsesTangent As Boolean, mPredictionVerification As Variant\nPrivate mPredictorComparison As Variant')
TestReplace('    mJointError = -1#: mStepError = -1#: mUsesTangent = False: mPredictionVerification = Empty','    mJointError = -1#: mStepError = -1#: mUsesTangent = False: mPredictionVerification = Empty: mPredictorComparison = Empty')
old='''                Dim tangentScale As Double: tangentScale = Sgn(orientation) * spacing / tangentLength
                For i = 1 To UBound(movement): movement(i) = tangentScale * tangent(i): Next i
                dl = tangentScale
                Dim destination() As Double: destination = accepted
                For i = 1 To UBound(destination): destination(i) = accepted(i) + movement(i): Next i
                If Not mSolver.SelectSpatialDirectionControl(accepted, destination, movement, projection, coordinate, spacing) Then Exit Function
                mUsesTangent = True
                mPredictionVerification = Array(Array(1# / 1024#, dl, tangentLength, orientation, alpha), probe, tangent, movement)'''
new='''                Dim tangentScale As Double: tangentScale = Sgn(orientation) * spacing / tangentLength
                Dim proposed() As Double: proposed = movement
                For i = 1 To UBound(proposed): proposed(i) = tangentScale * tangent(i): Next i
                Dim secantRadius As Double, tangentRadius As Double, sampleRadius As Double, quality As Variant, useTangent As Boolean
                secantRadius = Sqr(spacing * spacing + alpha * alpha * dl * dl)
                tangentRadius = Sqr(spacing * spacing + alpha * alpha * tangentScale * tangentScale)
                sampleRadius = PREDICTOR_PROBE_RATIO * GeomMin(secantRadius, tangentRadius)
                useTangent = True
                If mSolver.TryReadArcPredictorDefects(mBase, mDirection, accepted, factor(4), movement, dl, proposed, tangentScale, _
                        sampleRadius / secantRadius, sampleRadius / tangentRadius, mAbsolute, mRelative, quality) Then
                    useTangent = Not (CDbl(quality(0)) + CDbl(quality(2)) < CDbl(quality(1)) And _
                        CDbl(quality(0)) < PREDICTOR_IMPROVEMENT_RATIO * CDbl(quality(1))) ' Выбор прогноза требует различимого четырёхкратного выигрыша; это не изменение силового допуска приёмки.
                    mPredictorComparison = Array(Abs(CLng(useTangent)), quality(0), quality(1), quality(2), quality(3), quality(4))
                End If
                If useTangent Then
                    movement = proposed: dl = tangentScale
                    Dim destination() As Double: destination = accepted
                    For i = 1 To UBound(destination): destination(i) = accepted(i) + movement(i): Next i
                    If Not mSolver.SelectSpatialDirectionControl(accepted, destination, movement, projection, coordinate, spacing) Then Exit Function
                    mUsesTangent = True
                    mPredictionVerification = Array(Array(1# / 1024#, dl, tangentLength, orientation, alpha), probe, tangent, movement)
                End If'''
TestReplace(old,new)
anchor='    mDetails = mDetails & "; совместная ошибка " & CStr(mJointError) & "; исходящая касательная " & CStr(mUsesTangent)'
TestReplace(anchor,anchor+'''
    If Not IsEmpty(mPredictorComparison) Then mDetails = mDetails & "; касательный прогноз выбран по невязке: " & CStr(CBool(mPredictorComparison(0))) & _
        "; дефект секущей, МПа: " & CStr(mPredictorComparison(1)) & "; дефект касательной, МПа: " & CStr(mPredictorComparison(2))''')
code+='''

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Копирует сравнение прогнозов текущей пробы; не вызывает материал и не меняет историю.
Friend Property Get TestPredictorComparison() As Variant
    TestPredictorComparison = mPredictorComparison
End Property
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
tree=ast.parse((here/'Stage-TestArc56Diagnostics.py').read_text('utf-8-sig'))
entry,=[ast.literal_eval(n.value) for n in tree.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='entry' for t in n.targets)]
entry=entry.replace('Arc56','Arc32').replace('ARC56','ARC32').replace('Count <> 56','Count <> 32').replace('история56','история32')
anchor='        TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest'
assert entry.count(anchor)==1
entry=entry.replace(anchor,'''        selected = arc.TestPredictorComparison
        If IsEmpty(selected) Then
            TestWriteReverseBinaryVector handle, Array(0#)
        Else
            TestWriteReverseBinaryVector handle, Array(1#)
            TestWriteReverseBinaryVector handle, selected
        End If
'''+anchor,1)
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig');anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
values[name]=code.replace(anchor,anchor+"' FOR TESTS: baseline entry point only; existing methods retained.\n\n\n"+entry,1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledArcContinuation','CCoupledNewtonSolver','modTestCoupledElasticState') -EntryPointNames 'TestRunArc32Diagnostics'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig');anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected arc32 residual predictor setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunArc32Diagnostics"))
        Save-Progress
''',1);values[name]=runner.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/Reverse1.bin':(source/'Input/Reverse1.bin').read_bytes(),
 'Input/SourceColdRawHistory.bin':(base/'Cold904ConsistentRefinementCold96_20261126/Native/Stock.RawHistory.bin').read_bytes(),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():
    if name.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ActualResidualArcPredictorSelection',actual_import_this_pass=['CCoupledArcContinuation','CCoupledNewtonSolver','modTestCoupledElasticState'],
 actual_component_count=256,source_cold_history32_exact=True,only_two_working_classes_changed=True,
 predictor_probe_ratio=1/32,predictor_required_improvement_ratio=.25,comparison_uses_existing_equilibrium_tolerance=True,
 no_material_or_force_acceptance_changes=True,primary30_and_global96_unchanged=True,
 extra_three_readonly_residual_evaluations_inside_one_predictor_preparation=True,
 no_material_history_acceptance=True,physical_capacity_certified=False,no_working_or_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'READONLY RESIDUAL PREDICTOR SELECTION;SAME PHYSICAL/FULL-HALF ACCEPTANCE;NO LIMIT CLAIM')
