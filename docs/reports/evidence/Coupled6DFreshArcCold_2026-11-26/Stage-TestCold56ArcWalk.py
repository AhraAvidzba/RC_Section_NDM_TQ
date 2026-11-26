"""ДЛЯ ТЕСТОВ. Продолжает точную холодную историю56 существующей рабочей дугой с её штатным адаптивным отношением и новым отдельным бюджетом96."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold888StrictEnvelopeGrowth_20261126';source=base/'Cold893Arc56Diagnostics_20261126';root=base/'Cold894Cold56ArcWalk_20261126'
proof=json.loads((source/'Arc56DiagnosticsProof.json').read_bytes());assert all(c['solved'] for c in proof['cases'])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
code=(root/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig');method,=re.findall(r'(?ms)^Public Sub TestRunFixedMetricArcWalk\(.*?^End Sub',code)
assert 'For attempt = 1 To 32' in method and 'totalCalls + 3 > 96' in method and 'attemptsAtRevision >= 6' in method
assert 'ratio = GeomMin(arc.NextStepRatioLimit, 0.8 * Sqr(0.005 / GeomMax(arc.StepError, 0.00000001)))' in method
prep=json.loads((root/'Preparation.json').read_bytes());name='tools/build_workbook/Run-Audit02ModeTests.ps1'
runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig');anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected cold56 arc-walk setup macro.'}
        $null=$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunFixedMetricArcWalk")
        $lines.Add('TOTAL_COLD56_ARC_WALK_EXECUTION_TEST: passed=1; failed=0; procedure completed; numerical outcomes recorded separately; not capacity')
        Save-Progress
''',1)
values={name:runner.replace('\n','\r\n').encode('utf-8-sig'),'Input/Reverse1.bin':(source/'Input/Reverse1.bin').read_bytes(),
        'Input/SourceColdRawHistory.bin':(source/'Input/SourceColdRawHistory.bin').read_bytes(),
        'SourceArc56Proof.json':(source/'Arc56DiagnosticsProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExactCold56FixedMetricArcWalk',actual_import_this_pass=json.loads((parent/'docs/regression/CodeUpdate.json').read_bytes())['Modules'],
    actual_imports_this_replay=[],actual_component_count=256,all_working_and_test_VBA_exact888=True,
    source_history56_exact=True,source_history_replayed_from_zero=True,continuation_research_not_cold_from_zero=True,
    separate_follow_on_budget96=True,attempts32=True,max6_ratios_at_revision=True,primary30_unchanged=True,
    every_accepted_step_has_verified_half_pair=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
print(root,'EXACT COMPILED888;NO IMPORT/COMPILE;FOLLOW-ON ARC96 NOT A COLD RUN')
