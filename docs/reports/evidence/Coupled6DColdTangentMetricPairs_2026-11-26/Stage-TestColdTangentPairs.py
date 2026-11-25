"""ДЛЯ ТЕСТОВ. Проверяет управляемые полные/половинные пробы с настоящей историей19 и касательным прогнозом спада нагрузки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
from argparse import ArgumentParser

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold858BalancedInitialStressReplay_20261125';tangents=base/'Cold864Cold19BranchTangents_20261125';verified=base/'Cold859SpatialDeclineStressReplay_20261125'
p=ArgumentParser(description=__doc__);p.add_argument('--reference-metric',action='store_true');args=p.parse_args()
root=base/('Cold866Cold19ReferenceMetricPairs_20261126' if args.reference_metric else 'Cold865Cold19TangentPairs_20261125');proof=json.loads((tangents/'ColdBranchTangentsProof.json').read_bytes())
assert proof['native']['own_process_closed'] and proof['exact19_history_and_frozen_stresses_unchanged']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ColdTangentPairsCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
cases=[(proof['tangents'][1],.125),(proof['tangents'][1],.0625),(proof['tangents'][2],.125)]
if args.reference_metric:cases=[(proof['tangents'][1],ratio) for ratio in (.125,.0625,.03125)]
data=vector([len(cases)]);labels=[]
for item,ratio in cases:
    assert item['same_length_lambda_increment']<0 and item['cosine_with_incoming_secant']<-.99
    data+=vector([item['level'],ratio,item['same_length_lambda_increment']])+vector(item['same_length_field_movement'])
    labels.append(dict(level=item['level'],ratio=ratio,negative_predicted_dlambda=item['same_length_lambda_increment']))
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
entry=(here/'CandidateTestColdTangentPairs.bas.txt').read_text('utf-8-sig');assert 'Public Function TestRunColdTangentPairs(' not in code
if args.reference_metric:
    anchor='    Dim alpha As Double: alpha = GeomMax(GeomMin(slopes(1), slopes(2)), GeomMin(GeomMax(slopes(1), slopes(2)), slopes(3)))\n';assert entry.count(anchor)==1
    entry=entry.replace(anchor,anchor+'''
    Dim medianScale As Double: medianScale = alpha
    Dim linearDeformation() As Double, linearRecovered() As Double, linearStress() As Double, rawLinear As Variant
    If Not mReverseOperatorsTest.Solve(mReverseDirectionTest(1), mReverseDirectionTest(2), mReverseDirectionTest(3), _
            mReverseDirectionTest(4), mReverseDirectionTest(5), mReverseDirectionTest(6), linearDeformation, linearRecovered, linearStress, rawLinear) Then _
        Err.Raise 5, , "Не подготовлен постоянный упругий масштаб направления LoadPath."
    Dim linearField() As Double, zeroField() As Double
    ReDim linearField(1 To 60): ReDim zeroField(1 To 60)
    For i = 1 To 60: linearField(i) = rawLinear(i): Next i
    alpha = mReverseSolverTest.SpatialFieldDistance(zeroField, linearField)
    If alpha <= 0# Then Err.Raise 5, , "Упругий масштаб дуги должен быть положительным."
    Dim referenceHandle As Integer: referenceHandle = FreeFile
    Open mPolishRootTest & "\\Native\\ColdReferenceMetric.bin" For Binary Access Write As #referenceHandle
    Dim referenceMagic As String * 16: referenceMagic = "NDMCOLDREFMET001"
    Put #referenceHandle, , referenceMagic
    TestWriteReverseBinaryVector referenceHandle, Array(alpha, medianScale, mReverseModulusTest)
    TestWriteReverseBinaryVector referenceHandle, mReverseDirectionTest
    TestWriteReverseBinaryVector referenceHandle, linearField
    TestWriteReverseBinaryVector referenceHandle, linearRecovered
    Close #referenceHandle
''',1)
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+"' FOR TESTS: baseline entry point only; existing methods retained.\n\n\n"+entry,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');anchor="-ModuleNames 'modTestCoupledElasticState'";assert compiler.count(anchor)==1
values['Compile-TestExternalPoint.ps1']=compiler.replace(anchor,anchor+" -EntryPointNames 'TestRunColdTangentPairs'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';script=(root/name).read_text('utf-8-sig');anchor='        else { $result = [string]$excel.Run($macroName) }';assert script.count(anchor)==1
script=script.replace(anchor,anchor+'''
        # ДЛЯ ТЕСТОВ: перед контролируемыми пробами восстановлена настоящая история19, код материала не меняется.
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected setup macro for tangent pairs.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunColdTangentPairs"))
        Save-Progress
''',1)
values[name]=script.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/Reverse1.bin':(verified/'Input/Reverse1.bin').read_bytes(),'Input/ColdTangentPredictors.bin':data,
 'Input/ColdTangentPredictors.json':(json.dumps(labels,indent=2)+'\n').encode(),
 'SourceColdTangentProof.json':(tangents/'ColdBranchTangentsProof.json').read_bytes(),
 'CandidateTestColdTangentPairs.bas.txt':(here/'CandidateTestColdTangentPairs.bas.txt').read_bytes(),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Cold19TangentControlledPairs',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
 only_short_test_entry_added=True,all_working_methods_and_materials_exact858=True,
 source_history19_exact=True,previous_fields_not_replaced_by_artificial_secant=True,
 predictions_from_verified_native_tangents_no_future_field=True,negative_dlambda_predictor_allowed=True,
 three_independent_pairs_same_history=True,original30pairbudget_and_strict_accuracy_scale1e_minus4=True,
 restoration_requires_zero_iterations_same_lambda_bit_identical_field=True,
 all_root_admissibility_and_half_checks_independent_and_pending=True,
 reference_metric_from_existing_prepared_elastic_solve=args.reference_metric,
 fixed_metric_for_given_section_and_load_direction=args.reference_metric,
 no_branch_or_capacity_acceptance=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'THREE NATIVE FULL/HALF PAIRS;TRUE HISTORY19;SIGNED TANGENT PREDICTOR;SAME MATERIALS/TOLERANCES;FIXED REFERENCE METRIC',args.reference_metric)
