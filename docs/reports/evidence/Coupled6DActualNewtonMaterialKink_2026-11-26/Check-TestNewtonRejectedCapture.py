"""ДЛЯ ТЕСТОВ. Проверяет неизменность восьми проб и реальные системы/направления в отклонённых шагах Newton."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold911NewtonRejectedCapture_20261126';parent=base/'Cold909ResidualArcPredictorBaseline_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['CCoupledNewtonSolver','modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==254
oldproof=json.loads((parent/'ResidualArcPredictorBaselineProof.json').read_bytes())
assert oldproof['all_available_roots_independently_equilibrated']
exactfiles=[]
for path in (parent/'Native').glob('*'):
    if path.name=='Arc32Diagnostics.bin' or path.name.startswith('Arc32.Case') or path.name.endswith('.Operators.csv'):
        assert path.read_bytes()==(root/'Native'/path.name).read_bytes(),path.name
        exactfiles.append(path.name)
assert 'Arc32Diagnostics.bin' in exactfiles
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
known=[('cold'+str(i+1),field) for i,field in enumerate(source['fields'])]
r=R((root/'Native/Arc32Diagnostics.bin').read_bytes())
assert r.raw[:16]==b'NDMARC32DIAG0001'
for case in range(1,9):
    meta=r.VectorTest();packet=[r.VectorTest() for _ in range(14)]
    known.extend([('case'+str(case)+'packet'+str(i),field) for i,field in enumerate(packet) if len(field)==60])
    for _ in range(3):
        header=r.VectorTest()
        for _ in range(int(header[1])):r.VectorTest()
    comparison=r.VectorTest();assert comparison.tolist()==[0.]
    r.VectorTest()
    if meta[3]:known.append(('case'+str(case)+'returned',r.VectorTest()));r.VectorTest()
assert r.position==len(r.raw)
records=[];eps=np.finfo(float).eps
files=sorted((root/'Native').glob('NewtonRejected.*.bin'));assert files
for serial,path in enumerate(files,1):
    assert path.name==f'NewtonRejected.{serial:05d}.bin'
    r=R(path.read_bytes());r.position=0
    meta,field,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
    assert len(meta)==9 and len(field)==len(accepted)==len(residual)==60 and len(rhs)==len(delta)==61
    A=np.array([r.VectorTest() for _ in range(61)])
    end=r.VectorTest();count=int(r.VectorTest()[0]);probes=[]
    np.testing.assert_array_equal(rhs[:60],-residual)
    linear=A@delta;den=np.abs(A)@np.abs(delta)+np.abs(rhs)
    backward=float(max(abs(linear-rhs)/np.maximum(den,np.finfo(float).tiny)))
    assert backward<=64*61*eps,backward # Только проверка округления линейной системы, не допуск физического равновесия.
    physical=int(meta[3])-1;assert physical==56 and meta[4]==30000 and end[4] in (0,1)
    governing=int(np.argmax(abs(residual[:physical])))
    match=[name for name,f in known if np.array_equal(f,accepted)]
    for _ in range(count):
        info,trial,actual=r.VectorTest(),r.VectorTest(),r.VectorTest()
        assert len(info)==6 and len(trial)==len(actual)==60 and info[0]>0 and info[5]==0
        f=info[0];roundscale=abs(field)+abs(f*delta[:60])
        assert np.all(abs(trial-(field+f*delta[:60]))<=4*eps*np.maximum(roundscale,np.finfo(float).tiny))
        assert abs(info[1]-(meta[2]+f*delta[60]))<=4*eps*(abs(meta[2])+abs(f*delta[60]))
        observed=(actual-residual)/f
        error=observed-linear[:60]
        row=int(np.argmax(abs(error[:physical])))
        assert info[2]>info[3] and info[2]>=end[1]*(1-0.0001*f)
        probes.append(dict(fraction=float(f),lambda_value=float(info[1]),norm_MPa=float(info[2]),bound_MPa=float(info[3]),
            maximum_physical_direction_error_MPa=float(max(abs(error[:physical]))),largest_error_row=row+1,
            governing_actual_direction_MPa=float(observed[governing]),governing_predicted_direction_MPa=float(linear[governing]),
            governing_direction_opposes_prediction=bool(observed[governing]*linear[governing]<0),
            control_direction_error_MPa=float((info[4]+rhs[-1])/f-linear[-1])))
    assert r.position==len(r.raw)
    records.append(dict(serial=serial,bytes=path.stat().st_size,sha256=sha256(path.read_bytes()).hexdigest().upper(),
        accepted_revision=int(meta[0]),iteration=int(meta[1]),lambda_value=float(meta[2]),
        accepted_exact_matches=match,consistent=bool(meta[7]),refined=bool(meta[8]),
        norm_MPa=float(end[1]),bound_MPa=float(end[2]),governing_row=governing+1,
        governing_residual_MPa=float(residual[governing]),linear_backward_error=backward,
        valid_line_probes=count,probes=probes))
proof=dict(native=native,actual_components=256,unchanged_other_components=254,
    baseline8_cases_and_all_root_exports_bit_identical=True,exact_native_files=exactfiles,
    independently_verified_roots_in_baseline_proof=True,records=records,
    recorded_rejected_directions=len(records),no_additional_material_evaluations=True,
    only_valid_probe_residuals_recorded=True,invalid_probe_residuals_not_used=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'NewtonRejectedCaptureProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ACTUAL REJECTED DIRECTIONS',len(records),'BASELINE FILES EXACT',len(exactfiles),'MAX LINEAR BACKWARD ERROR',max(x['linear_backward_error'] for x in records))
for item in records:
    if item['probes']:
        p=item['probes'][-1]
        print(item['serial'],'row',item['governing_row'],'norm',item['norm_MPa'],'refined',item['refined'],'last error',p['maximum_physical_direction_error_MPa'],'opposite',p['governing_direction_opposes_prediction'])
