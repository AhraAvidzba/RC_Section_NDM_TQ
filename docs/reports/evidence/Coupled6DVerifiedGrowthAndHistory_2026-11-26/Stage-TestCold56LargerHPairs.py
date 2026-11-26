"""ДЛЯ ТЕСТОВ. Проверяет шесть увеличенных полных/половинных H-проб на фактическом холодном состоянии56 без изменения рабочего алгоритма."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import ast,json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold881SpatialSteelLimits_20261126';source=base/'Cold880CriticalHCold96_20261126';root=base/'Cold886Cold56LargerHPairs_20261126'
proof=json.loads((source/'CriticalHCold96Proof.json').read_bytes());assert proof['accepted_states']==56 and proof['cold']['native']['own_process_closed']
stress=json.loads((base/'Cold882CriticalHColdStress_20261126/VerifiedColdStressReplayProof.json').read_bytes());assert stress['source_states']==56
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'Cold56PairCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];original=read(source/'Input/Reverse1.bin');prior=read(parent/'Input/Reverse1.bin')
for key in ('calibration','options','offset','direction'):np.testing.assert_array_equal(original[key],prior[key])
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](source/'Native/Stock.RawHistory.bin');frame=frames[-1];assert len(frame['fields'])==56 and raw['incomplete_tail_bytes']==0
values['Input/Reverse1.bin']=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frame)
values['Input/SourceColdRawHistory.bin']=(source/'Native/Stock.RawHistory.bin').read_bytes()
tree=ast.parse((here/'Stage-TestVerifiedHCornerPairs.py').read_text('utf-8-sig'))
entry,=[ast.literal_eval(n.value) for n in tree.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='entry' for t in n.targets)]
entry=entry.replace('TestRunVerifiedHPairs','TestRunCold56LargerHPairs').replace('VerifiedHPairs','Cold56LargerHPairs').replace('VerifiedH.Case','Cold56H.Case')
entry=entry.replace('Проверяет три независимых общих H-шага','Проверяет шесть увеличенных независимых общих H-шагов на холодной истории56')
assert entry.count('For index = 1 To 3')==1;entry=entry.replace('For index = 1 To 3','For index = 1 To 6',1)
assert entry.count('ratio = 0.25 ^ (index - 1)')==1;entry=entry.replace('ratio = 0.25 ^ (index - 1)','ratio = 2# ^ (index - 1)',1)
old='factor + ratio * (factor - previousFactor), target, returned, projection, absTol, relTol, 30, probes, False)'
assert entry.count(old)==1;entry=entry.replace(old,old[:-1]+', True)',1)
entry=entry.replace('TOTAL_VERIFIED_H_PAIRS_TEST','TOTAL_COLD56_LARGER_H_PAIRS_TEST')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig');anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+"' FOR TESTS: baseline entry point only; existing methods retained.\n\n\n"+entry,1);values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunCold56LargerHPairs'",1)
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig');anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected cold56 H-pair setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunCold56LargerHPairs"))
        Save-Progress
''',1);values[name]=runner.replace('\n','\r\n').encode('utf-8-sig')
for name,data in {'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
                  'SourceColdProof.json':(source/'CriticalHCold96Proof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}.items():values[name]=data
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Cold56LargerHPairs',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all_working_sources_exact881=True,source_cold_states=56,source_cold_history_sha256=raw['sha256'],
    pair_H_step_ratios=[1,2,4,8,16,32],every_pair_budget30=True,every_pair_original_strict_tolerances=True,
    accepted_material_history_not_advanced=True,physical_capacity_certified=False,no_working_or_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'SIX INDEPENDENT STRICT H PAIRS ON ACTUAL COLD56;SAME30/TOLERANCES;NO WORKING CHANGE')
