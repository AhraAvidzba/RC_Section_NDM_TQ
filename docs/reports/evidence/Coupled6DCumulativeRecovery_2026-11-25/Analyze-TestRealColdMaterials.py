"""ДЛЯ ТЕСТОВ. Читает фактические главные деформации и напряжения742, не вводя порог разрушения бетона."""
from pathlib import Path
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold742RealColdFrozenReplay_20261125')
mod=run_path(str(here/'Validate-TestHardPrepared.py'));rows=mod['TestRows'](root/'Native/EnvelopeEnergy.Accepted.csv')
operators=mod['TestRows'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv')
dimension=next(r for r in rows if r[0]=='DIM');dofs,physical,points,concrete=map(int,dimension[1:5]);half=float(dimension[5])
b=np.zeros((points,6,dofs));gradient=b.copy();axis=np.zeros((points,6));areas=np.zeros(points)
for kind,*r in operators:
    if kind=='B':p,d=int(r[0])-1,int(r[1])-1;b[p,:,d]=np.asarray(r[2:8],float);gradient[p,:,d]=np.asarray(r[8:],float)
    elif kind=='A':axis[int(r[0])-1]=np.asarray(r[1:],float)
    elif kind=='P':areas[int(r[0])-1]=float(r[3])
law=np.asarray(next(r[1:] for r in rows if r[0]=='LAW'),float)
frames=[r for r in rows if r[0]=='F'];state=[]
for frame in frames:
    number=int(frame[1]);field=np.asarray(frame[4:],float)
    center=np.einsum('pcd,d->pc',b,field);slope=np.einsum('pcd,d->pc',gradient,field)
    strains=np.asarray([center+s*half*slope for s in (-1,0,1)])
    tensor=np.zeros((3,concrete,3,3));e=strains[:,:concrete]
    tensor[:,:,0,0]=e[:,:,0];tensor[:,:,1,1]=e[:,:,1];tensor[:,:,2,2]=e[:,:,2]
    for component,i,j in ((3,0,1),(4,0,2),(5,1,2)):tensor[:,:,i,j]=tensor[:,:,j,i]=.5*e[:,:,component]
    principal=np.linalg.eigvalsh(tensor)
    stress=np.asarray([list(map(float,r[4:])) for r in rows if r[0]=='S' and int(r[1])==number]).reshape(3,points,9)
    bars=np.einsum('spc,pc->sp',strains[:,concrete:],axis[concrete:])
    native=next(r for r in rows if r[0]=='R' and int(r[1])==number)
    state.append(dict(index=number,lambda_value=float(frame[3]),
        minimum_concrete_principal_total_strain=float(principal.min()),maximum_concrete_principal_total_strain=float(principal.max()),
        maximum_concrete_normal_stress_MPa=float(abs(stress[:,:concrete,:3]).max()),
        maximum_tension_envelope_energy_fraction=float(stress[:,:concrete,6].max()),
        maximum_compression_envelope_energy_fraction=float(stress[:,:concrete,7].max()),
        minimum_all_bar_axial_strain=float(bars.min()),maximum_all_bar_axial_strain=float(bars.max()),
        maximum_abs_all_bar_stress_MPa=float(abs(stress[:,concrete:,8]).max()),
        native_maximum_steel_deformation_utilization=float(native[4])))
    energy=law[6]/law[8]*stress[:,:concrete,6]+law[7]/law[8]*stress[:,:concrete,7]
    state[-1]['reconstructed_envelope_coordinate_MPa']=float(np.sum(energy*areas[None,:concrete])/(3*sum(areas[:concrete])))
proof=dict(states=state,derived_from_actual742_fields_and_operators=True,
    engineering_shear_divided_by2_in_strain_tensor=True,energy_fractions_are_not_damage_thresholds=True,
    principal_strain_extremes_are_not_a_calibrated_multiaxial_failure_criterion=True,
    no_physical_capacity_claimed=True)
(root/'ActualColdMaterialMeasures.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('742 ACTUAL LAST MATERIAL MEASURES',json.dumps(state[-1],sort_keys=True))
print('742 RECONSTRUCTED ENVELOPE PATH',[(s['index'],s['reconstructed_envelope_coordinate_MPa']) for s in state])
