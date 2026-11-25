"""ДЛЯ ТЕСТОВ. Проверяет происхождение H-границы, автоматические равновесия и фактический рост материальной координаты после неё."""
from pathlib import Path
from runpy import run_path
import json,numpy as np,sys

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold842HReentryWalk_20261125';parent=base/'Cold841HGuideMaterialEvent_20261125'
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];source=read(parent/'Input/Reverse1.bin')
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=Reader((parent/'Native/HGuideEvent.bin').read_bytes())
record,info,state,face,half,candidate,restore,effort=[r.VectorTest() for _ in range(8)]
assert record[0]==1 and restore[0]==1
expected=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](source,
    dict(fields=np.vstack([source['fields'],candidate]),fractions=np.ones(71),lambdas=np.append(source['lambdas'],record[8])))
assert (root/'Input/Reverse1.bin').read_bytes()==expected
saved=sys.argv;sys.argv=['Check-TestAutomaticLoadingWalk.py','--root',str(root),'--initial-input',str(root/'Input/Reverse1.bin'),'--h-reentry-prefix']
try:run_path(str(here/'Check-TestAutomaticLoadingWalk.py'),run_name='__main__')
finally:sys.argv=saved
walk=json.loads((root/'AutomaticLoadingWalkProof.json').read_bytes())
assert walk['initial_source_states']==71 and walk['new_accepted_equilibria']==8
validate=run_path(str(here/'Validate-TestHardPrepared.py'));rows=validate['TestRows'](root/'Native/LoadingWalk.Accepted.Sequence.csv')
count=len(walk['all_equilibria']['states']);stored=np.zeros((count,3,43,2))
for row in rows:
    if row[0]=='S':stored[int(row[1])-1,int(row[2])+1,int(row[3])-1]=np.array(row[10:12],float)
coords=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/LoadingWalk.Accepted.Sequence.Operators.csv',60,43)[4]
calibration=read(root/'Input/Reverse1.bin')['calibration'];gt,gc,length=calibration[6:9]
h=np.einsum('nspc,p,c->n',stored[:,:,:25],coords[:25,2],np.array([gt/length,gc/length]))/(3*sum(coords[:25,2]))
increments=np.diff(h[70:]);all_growth=bool(np.all(increments>0))
proof=dict(walk=walk,exact70_plus_verified841_face_prefix=True,source_H=h[70],new_H_levels=h[71:].tolist(),
    new_H_increments=increments.tolist(),all_new_states_have_strictly_growing_H=all_growth,
    maximum_new_lambda=max(x['lambda'] for x in walk['all_equilibria']['states'][71:]),
    first_loading_departure_kind=next(a['kind'] for a in walk['attempts'] if a['solved']),
    no_new_material_parameters_or_working_solver_changes=True,
    not_cold_capacity_search=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'HReentryWalkProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('H REENTRY NEW8 VERIFIED;H GROWS',all_growth,'FIRST KIND',proof['first_loading_departure_kind'],'MAX NEW LAMBDA',proof['maximum_new_lambda'])
assert all_growth,'A later segment has frozen H; do not present it as a loading capacity branch.'
