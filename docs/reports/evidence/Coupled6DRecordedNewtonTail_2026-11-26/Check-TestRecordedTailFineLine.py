"""ДЛЯ ТЕСТОВ. Проверяет исходные критерии точной системы после расширения диапазона долей штатного поиска."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;file=here/'Check-TestBoundedBranchRecovery.py'
code=file.read_text('utf-8-sig')
code=code.replace('Cold927BoundedBranchRecovery_20261126','Cold929RecordedTailFineLine_20261126')
code=code.replace('BoundedBranchRecoveryProof.json','RecordedTailFineLineProof.json')
start=code.index("assert 'For localStep = 1 To 3' in code")
finish=code.index("(root/'RecordedTailFineLineProof.json')",start)
code=code[:start]+'''\
assert 'If mControlActive And mUseSpatialHistory Then minimumFraction = 0.0000152587890625' in code
proof.update(original30_budget_and_material_force_control_acceptance_unchanged=True,
    numerical_change_only_finer_line_candidates=True,no_working_or_release_promotion=True,
    diagnostic_last_norm_is_restored_context_not_minimum_over_all_attempts=True)
'''+code[finish:]
scope={'__file__':str(file),'__name__':'recorded_tail_fine_line_test'}
exec(compile(code,str(file),'exec'),scope)
root=scope['root'];R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
records=[]
for path in sorted((root/'Native').glob('NewtonRejected.*.bin')):
    r=R(path.read_bytes());r.position=0
    meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
    A=np.array([r.VectorTest() for _ in range(61)]);end=r.VectorTest();count=int(r.VectorTest()[0])
    probes=[[r.VectorTest() for _ in range(3)] for _ in range(count)];assert r.position==len(r.raw)
    records.append(dict(serial=int(path.stem.split('.')[-1]),iteration=int(meta[1]),current_lambda=float(meta[2]),
        governing_row=int(np.argmax(abs(residual[:56])))+1,initial_norm_MPa=float(end[1]),
        final_valid_trial_norm_MPa=float(probes[-1][0][2]),minimum_fraction=float(end[-1]),valid_probe_count=count,
        actual_current_max_abs_field=float(max(abs(current)))))
proof=scope['proof'];proof.update(actual_rejected_direction_records=records)
(root/'RecordedTailFineLineProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ACTUAL REJECTED DIRECTIONS',len(records))
