"""ДЛЯ ТЕСТОВ. Связывает реальные отклонённые системы холодного923 с точными принятыми полями и проверяет фактические направления."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold923RecordedNewtonCold96_20261126';prior=base/'Cold920AffineRecoveryCold96_20261126'
cold=json.loads((root/'RecordedNewtonCold96Proof.json').read_bytes())
assert cold['entire_raw_trajectory_and_intermediate_frames_bit_identical920'] and cold['all257_saved_sources_verified']
read=run_path(str(here/'Read-TestColdRawCapture.py'));R=read['TestReader'];frames,history=read['TestHistory'](root/'Native/Stock.RawHistory.bin')
assert history['incomplete_tail_bytes']==0
matches={}
for frame in frames:
    for i,field in enumerate(frame['fields']):
        prefix=sha256(frame['fields'][:i+1].tobytes()+frame['fractions'][:i+1].tobytes()+frame['lambdas'][:i+1].tobytes()).hexdigest().upper()
        matches.setdefault(field.tobytes(),set()).add((i+1,float(frame['lambdas'][i]),prefix))
records=[];eps=np.finfo(float).eps
files=sorted((root/'Native').glob('NewtonRejected.*.bin'));assert files
for serial,path in enumerate(files,1):
    assert path.name==f'NewtonRejected.{serial:05d}.bin'
    r=R(path.read_bytes());r.position=0
    meta,field,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
    assert [len(x) for x in (meta,field,accepted,residual,rhs,delta)]==[9,60,60,60,61,61]
    A=np.array([r.VectorTest() for _ in range(61)]);end=r.VectorTest();n=int(r.VectorTest()[0]);assert len(end)==6
    np.testing.assert_array_equal(rhs[:60],-residual)
    linear=A@delta;denominator=abs(A)@abs(delta)+abs(rhs)
    backward=float(max(abs(linear-rhs)/np.maximum(denominator,np.finfo(float).tiny)))
    assert backward<=64*61*eps
    physical=int(meta[3])-1;assert physical==56 and meta[4]==30000
    governing=int(np.argmax(abs(residual[:physical])))
    full_control=original_control=-meta[5]*meta[4]+A[-1,:60]@field+A[-1,-1]*meta[2]
    control_round=64*eps*(abs(meta[5]*meta[4])+abs(A[-1,:60])@abs(field)+abs(A[-1,-1]*meta[2]))
    assert abs(full_control+rhs[-1])<=control_round+np.finfo(float).tiny
    probes=[]
    for _ in range(n):
        info,trial,actual=[r.VectorTest() for _ in range(3)]
        assert len(info)==6 and info[5]==0 and len(trial)==len(actual)==60
        f=info[0];assert f>0
        roundscale=abs(field)+abs(f*delta[:60])
        assert np.all(abs(trial-(field+f*delta[:60]))<=8*eps*np.maximum(roundscale,np.finfo(float).tiny))
        assert abs(info[1]-(meta[2]+f*delta[60]))<=8*eps*(abs(meta[2])+abs(f*delta[60]))
        observed=(actual-residual)/f;error=observed-linear[:60];row=int(np.argmax(abs(error[:physical])))
        probes.append(dict(fraction=float(f),lambda_value=float(info[1]),norm_MPa=float(info[2]),bound_MPa=float(info[3]),
            governing_actual_direction_MPa=float(observed[governing]),governing_predicted_direction_MPa=float(linear[governing]),
            governing_direction_opposes_prediction=bool(observed[governing]*linear[governing]<0),
            maximum_physical_direction_error_MPa=float(max(abs(error[:physical]))),largest_error_row=row+1))
    assert r.position==len(r.raw)
    matched=sorted(matches.get(accepted.tobytes(),set()))
    records.append(dict(serial=serial,sha256=sha256(path.read_bytes()).hexdigest().upper(),accepted_revision=int(meta[0]),
        iteration=int(meta[1]),lambda_value=float(meta[2]),accepted_history_prefix_matches=[dict(state=i,lambda_value=f,raw_prefix_sha256=h) for i,f,h in matched],
        accepted_field_zero=bool(np.all(accepted==0)),accepted_field_not_found_in_saved_history=not bool(matched),
        control_target=float(meta[5]),arc_load_projection=float(meta[6]),consistent=bool(meta[7]),refined=bool(meta[8]),
        norm_MPa=float(end[1]),bound_MPa=float(end[2]),governing_row=governing+1,
        governing_residual_MPa=float(residual[governing]),linear_backward_error=backward,valid_probes=n,probes=probes))
row55=[x for x in records if x['governing_row']==55]
large=[x for x in row55 if x['norm_MPa']>.01]
proof=dict(entire_cold_history_exact920=True,saved257_sources_verified=True,records=records,
    recorded_rejected_directions=len(records),row55_rejected_directions=len(row55),large_row55_rejected_directions=len(large),
    accepted_field_matches_are_byte_exact=True,field_match_alone_not_a_full_material_history_certificate=True,
    all_recorded_linear_systems_independently_verified=True,all_valid_probes_read_without_material_reevaluation=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'ColdRejectedContextsProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('REJECTED DIRECTIONS',len(records),'ROW55',len(row55),'LARGE ROW55',len(large))
for item in large[-12:]:
    p=item['probes'][-1] if item['probes'] else {}
    print(item['serial'],'ITER',item['iteration'],'BASE NORM',item['norm_MPa'],'ACCEPTED',[(x['state'],x['lambda_value']) for x in item['accepted_history_prefix_matches']],
        'TARGET',item['control_target'],'REFINED',item['refined'],'LAST DIRECTION',p.get('governing_actual_direction_MPa'))
