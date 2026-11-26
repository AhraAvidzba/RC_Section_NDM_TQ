"""ДЛЯ ТЕСТОВ. Проверяет прочитанные касательные21/+1 и отделяет линейные прогнозы от равновесных решений."""
from pathlib import Path
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold970PrefixSurfaceMatrices_20261126';parent=base/'Cold969PrefixSpatialFloor_20261126'
code=(here/'Check-TestSurfaceMatrices.py').read_text('utf-8-sig')
changes={
    "root=base/'Cold931SurfaceMatrices_20261126';parent=base/'Cold925RecordedTailSystem_20261126'":
        "root=base/'Cold970PrefixSurfaceMatrices_20261126';parent=base/'Cold969PrefixSpatialFloor_20261126'",
    "audit['imported_this_pass']==['modTestCoupledElasticState']":"audit['imported_this_pass']==['CCoupledNewtonSolver']",
}
for old,new in changes.items():assert code.count(old)==1;code=code.replace(old,new,1)
env={'__file__':str(here/'Check-TestSurfaceMatrices.py'),'__name__':'prefix_surface_matrices_test'}
exec(compile(code,str(here/'Check-TestSurfaceMatrices.py'),'exec'),env)
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
for path in (parent/'tests').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
proof=env['proof'];matrices=env['matrices'];original=env['original'];delta=env['old_delta'];baseline=-env['rhs']
source=env['source'];R=env['R'];r=R((root/'SourceAcceptedSurfaceStates.bin').read_bytes());states={}
for _ in range(129):
    point,side=r.VectorTest();states[int(point),int(side)]=r.VectorTest()
assert r.position==len(r.raw)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
helper=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig')
begin=helper.index('def TestYield(');end=helper.index('def TestAssemble(',begin)
exec(compile(helper[begin:end],str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
scans=[]
for item in matrices:
    level=item['level'];entry=dict(level=level)
    if level:
        measured=(item['residual']-baseline)/level;linear=original@delta
        error=float(max(abs((measured-linear)[:56])));size=float(max(abs(linear[:56])))
        entry.update(physical_directional_difference_MPa=error,relative_directional_difference=error/max(size,1e-300))
    entry['matrix_change_relative_to_initial']=float(np.linalg.norm(item['matrix']-original)/np.linalg.norm(original))
    entry['point21_plus_trial_loading_value_MPa']=TestYield(21,1,item['seed'])
    entry['unverified_corrected_fields']=[]
    for prediction in env['predictions']:
        if prediction['level']!=level:continue
        predicted=prediction['field']+prediction['correction'][:60]
        entry['unverified_corrected_fields'].append(dict(mode=prediction['mode'],
            point21_plus_trial_loading_value_MPa=TestYield(21,1,predicted),
            predicted_lambda=float(prediction['lambda_value']+prediction['correction'][-1]),
            equilibrium_and_branch_membership_not_verified=True))
    scans.append(entry)
proof.update(all257_actual_sources_and_saved_book_bitexact969=True,no_new_numerical_code_or_compilation=True,
    actual_source_record968=2,point=21,station_z_over_L=.001,source_matrix_refined=True,
    original25_input_bitexact969=True,directional_scans=scans,
    point21_plus_loading_value_at_accepted25_MPa=TestYield(21,1,source['fields'][-1]),
    directional_differences_cross_loading_surface_not_full_smooth_jacobian_proof=True,cold_gain_claimed=False)
(root/'SurfaceMatricesProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('EXISTING SURFACE ENTRY;EXACT969 BOOK;BASE MATRIX/RESIDUAL VERIFIED;21/+1 SCANS',scans)
