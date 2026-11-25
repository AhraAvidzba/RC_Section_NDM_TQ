"""ДЛЯ ТЕСТОВ. Повторно читает напряжения при нулевом приращении после принятия последнего холодного состояния без новой компиляции."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
compiled=base/'Cold858BalancedInitialStressReplay_20261125';verified=base/'Cold859SpatialDeclineStressReplay_20261125'
root=base/'Cold862ColdAcceptedStressContinuity_20261125'
code=(compiled/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub',code)
assert method[method.index('    mPtcSuppressTest = True\n'):]=='''    mPtcSuppressTest = True
    TestExportFastFrozenReplay
    mPtcSuppressTest = False
    mPolishExecutedTest = True
End Sub'''
scan,=re.findall(r'(?ms)^Private Sub TestReadUnresolvedArcScan\(.*?^End Sub',code)
assert 'recordTrial:=False' in scan and 'mReverseSolverTest.AcceptedRevision <> revision' in scan
assert not re.search(r'\b(?:AcceptStep|TryStep|TryDeformationStep)\b',scan)
proof=json.loads((verified/'VerifiedColdStressReplayProof.json').read_bytes());assert proof['source_states']==19
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];data=read(verified/'Input/Reverse1.bin')
original=read(compiled/'Input/Reverse1.bin')
for key in ('calibration','options','offset','direction'):np.testing.assert_array_equal(data[key],original[key])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,compiled,1e-6,1,1e-6,0)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
accepted=data['fields'][-1];delta=accepted-data['fields'][-2]
predictions=[accepted,accepted+.125*delta,accepted+delta]
prep=json.loads((root/'Preparation.json').read_bytes());values={
 'Input/Reverse1.bin':(verified/'Input/Reverse1.bin').read_bytes(),
 'Input/ArcScanFields.bin':vector([len(predictions)])+b''.join(vector(f) for f in predictions),
 'SourceColdStressProof.json':(verified/'VerifiedColdStressReplayProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
name='tools/build_workbook/Run-Audit02ModeTests.ps1';script=(root/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert script.count(anchor)==1
extra='''
        # ДЛЯ ТЕСТОВ: исходный экспорт уже закончен; вызывается только читающая процедура неизменённой частной книги.
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected setup macro for read-only scan.'}
        $lines.Add('READONLY_PRIVATE_SCAN_STARTED: modTestCoupledElasticState.TestReadUnresolvedArcScan')
        Save-Progress
        $null = $excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadUnresolvedArcScan")
        $lines.Add('READONLY_PRIVATE_SCAN_COMPLETED: no new equilibrium search')
        Save-Progress
'''
values[name]=script.replace(anchor,anchor+extra,1).replace('\n','\r\n').encode('utf-8-sig')
for name,raw in values.items():(root/name).write_bytes(raw);prep['inputs'][name]=sha256(raw).hexdigest().upper()
compilation=json.loads((compiled/'docs/regression/CodeUpdate.json').read_bytes())
prep['directed_scope']=dict(case='AcceptedCold19PostCommitStressContinuity',actual_import_this_pass=compilation['Modules'],
 actual_imports_this_replay=[],immutable_compiled_book_reused=True,compiled_source_root=compiled.name,
 verified_stress_root=verified.name,source_states=19,only_PS_runner_calls_existing_readonly_private_scan=True,
 no_workbook_or_VBA_edit=True,three_predictions_including_exact_zero_increment=True,
 all_queries_recordTrial_false_and_history_unchanged=True,no_new_NDS_or_capacity_search=True,
 physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(compiled/'docs/regression/Diagnostic.xlsm').read_bytes()
print(root,'IMMUTABLE COMPILED858;EXACT COLD19;READONLY POSTCOMMIT ZERO/FORWARD STRESS QUERIES')
