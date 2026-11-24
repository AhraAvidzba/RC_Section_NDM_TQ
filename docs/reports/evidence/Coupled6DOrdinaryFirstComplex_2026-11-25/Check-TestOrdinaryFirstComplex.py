"""ДЛЯ ТЕСТОВ. Проверяет парный НДС, холодный исход741 и фактически повторно открытую книгу без заявления о пределе."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from zipfile import ZipFile
import json,re

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold741OrdinaryFirstComplex_20261125';parent=base/'Cold739PseudoTimeProbe_20261125'
helpers=run_path(str(here/'Check-TestPersistentHEvidence.py'));read=run_path(str(here/'Read-TestColdRawCapture.py'))
direct,text=helpers['TestNative'](root,'PriorBranchHardShearDirect')
paired,ptext=helpers['TestNative'](parent,'PriorBranchHardShearDirect')
assert direct['passed']==paired['passed']==19 and direct['failed']==paired['failed']==0
assert 'direct=OK; capacity=N/A' in text and 'direct=OK; capacity=N/A' in ptext
for component in ('N','Mx','My','T','Qx','Qy'):
    pattern=r'OK: direct '+component+r' actual=([^ ]+) expected=([^\r\n]+)'
    assert re.findall(pattern,text)==re.findall(pattern,ptext),component
native,stock=helpers['TestNative'](root,'HBackendStock')
assert native['passed']==28 and native['failed']==0 and not native['timeout600']
saved,=re.findall(r'(?m)^RESULTS_SAVE_REOPEN: .*?equal=True; file=([^\r\n]+)',stock)
data=Path(saved).read_bytes();(root/'SavedWorkbook.xlsm').write_bytes(data)
with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book: compiled=book.read('xl/vbaProject.bin')
result=helpers['TestStock'](repo,root,compiled)
assert result['direct']=='OK' and result['capacity']=='NumFail' and result['code']=='NUMERICAL_FAILURE'
assert result['limit']==''
frames,raw=read['TestHistory'](root/'Native/Stock.RawHistory.bin');frame=frames[-1]
assert raw['incomplete_tail_bytes']==0 and len(frame['fields'])==12
assert abs(result['lower']-max(frame['lambdas']))<1e-13
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls',
             'tests/modTestCoupledElasticState.bas'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
name='src/Materials/CCoupledConcreteResponse.cls'
assert (root/name).read_bytes()==(base/'Cold711PriorDilationCurve_20261124'/name).read_bytes()
proof=dict(material_regressions=json.loads((root/'OrdinaryMaterialRegressionProof.json').read_bytes()),
    direct741=direct,direct739=paired,direct741_seconds=float(re.findall(r'TIMING_SECONDS: ([^ |]+)',text)[0]),
    direct739_seconds=float(re.findall(r'TIMING_SECONDS: ([^ |]+)',ptext)[0]),
    direct_all_six_printed_resultants_exact_between_pair=True,single_sequential_timing_pair_not_statistical_benchmark=True,
    cold741=result,raw_history=raw,accepted_states=12,raw_frames=len(frames),
    ordinary_material_exact711=True,newton_controller_test_module_exact739=True,
    finite_concrete_criterion_not_resolved=True,independent_tensor_reintegration_pending=True,
    no_physical_limit=True,no_working_promotion=True,whole_goal_completed=False)
(root/'OrdinaryFirstComplexProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('741 COLD BOTH',result['seconds'],'s; DIRECT OK, CAPACITY NUMFAIL; LOWER',result['lower'],
      '; REOPEN VERIFIED255; DIRECT PAIR19/0; NO PHYSICAL LIMIT')
