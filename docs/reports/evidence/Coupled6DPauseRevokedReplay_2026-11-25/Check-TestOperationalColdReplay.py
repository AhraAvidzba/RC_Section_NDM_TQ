"""ДЛЯ ТЕСТОВ. Независимо собирает все60 холодных равновесий785 и сопоставляет точность принятых корней с внутренним дуговым допуском."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np
from argparse import ArgumentParser

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__)
p.add_argument('--root',type=Path,default=base/'Cold788OperationalColdReplay_20261125')
p.add_argument('--source',type=Path,default=base/'Cold785OperationalArcController_20261125')
p.add_argument('--source-history-input',type=Path,default=Path('Input/Source785RawHistory.bin'))
args=p.parse_args();root=args.root;source_root=args.source
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/args.source_history_input)
expected=json.loads((source_root/'OperationalColdProof.json').read_bytes())['accepted_states']
assert raw['incomplete_tail_bytes']==0 and len(source['fields'])==expected
for key in ('fields','fractions','lambdas'):assert np.array_equal(source[key],frames[-1][key]),key
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls','src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls'):
    assert (root/name).read_bytes()==(source_root/name).read_bytes(),name
validate=run_path(str(here/'Validate-TestHardPrepared.py'))
mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
assert len(fields)==expected;np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
rows=validate['TestRows'](root/'Native/EnvelopeEnergy.Accepted.csv');reports=[r for r in rows if r[0]=='R']
operators=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
g,loads,w,v,coords,axis,b,gradient=operators
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
metric=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(coords[:,2]))
metric[56:,:]=0;metric[:,56:]=0
resolution=[]
for i,row in enumerate(reports):
    native_norm,user_bound,steel=map(float,row[2:5]);assert native_norm<=user_bound and steel<=1
    assert int(row[1])==i+1
    step=0.
    if i:
        delta=source['fields'][i]-source['fields'][i-1];step=float(np.sqrt(max(0.,delta@metric@delta))*source['calibration'][0])
    resolution.append(dict(state=i+1,factor=float(source['lambdas'][i]),native_weak_norm_MPa=native_norm,
        original_user_bound_MPa=user_bound,arc_inner_bound_MPa=.0001*user_bound,
        native_to_inner_tolerance=native_norm/(.0001*user_bound),previous_step_tensor_MPa=step,
        residual_to_step_tensor=(native_norm/step if step else None)))
proof=dict(native=native,source_name=source_root.name,accepted_states=expected,source_cold_exact=True,numerical_material_residual_methods_exact_source=True,
    all_accepted_states_independent_weak_and_six_force_equilibria=mechanics,
    resolution=resolution,states_above_arc_inner_tolerance=sum(x['native_to_inner_tolerance']>1 for x in resolution),
    maximum_native_to_inner_tolerance=max(x['native_to_inner_tolerance'] for x in resolution),
    resolution_comparison_does_not_prove_cause_of_arc_failure=True,
    source_capacity_NumFail_not_reclassified=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'OperationalColdReplayProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'COLD EQUILIBRIA VERIFIED',expected,'ABOVE INNER TOLERANCE',proof['states_above_arc_inner_tolerance'],'MAX RATIO',proof['maximum_native_to_inner_tolerance'],'NO PHYSICAL LIMIT')
for x in resolution[-5:]:print('TAIL',x)
