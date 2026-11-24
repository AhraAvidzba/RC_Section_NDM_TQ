"""ДЛЯ ТЕСТОВ. Проверяет пространственное равновесие и однородность принятой посттрещинной последовательности."""
from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
from runpy import run_path
import numpy as np


def TestAnalyse(root,phase):
    """Повторно собирает поля трёх станций из операторов; напряжения материала остаются нативными данными."""
    here=Path(__file__).resolve().parent
    mod=run_path(str(here/'Validate-TestHardPrepared.py'));proof,fields=mod['TestValidate'](root,phase)
    rows=mod['TestRows'](root/'Native'/(phase+'.csv'))
    dim=next(v for v in rows if v[0]=='DIM');dofs,physical,points,concrete=map(int,dim[1:5]);half=float(dim[5])
    b=np.zeros((points,6,dofs));gradient=b.copy()
    for kind,*row in mod['TestRows'](root/'Native'/(phase+'.Operators.csv')):
        if kind=='B':
            p,d=int(row[0])-1,int(row[1])-1
            b[p,:,d]=np.asarray(row[2:8],float);gradient[p,:,d]=np.asarray(row[8:],float)
    sigma=np.zeros((len(fields),3,points,6))
    for kind,*row in rows:
        if kind=='S':sigma[int(row[0])-1,int(row[1])+1,int(row[2])-1,:]=np.asarray(row[3:9],float)
    measured=[]
    for i,field in enumerate(fields):
        c=np.einsum('pcd,d->pc',b,field);z=np.einsum('pcd,d->pc',gradient,field)
        strain=np.array([c+side*half*z for side in (-1,0,1)])[:,:concrete];s=sigma[i,:,:concrete]
        measured.append(dict(index=i+1,lambda_value=proof['states'][i]['lambda'],control_epsilon=float(field[0]),
          max_epsilon_zz_difference=float(np.max(abs(strain[:,:,2]-field[0]))),
          max_shear_strain=float(np.max(abs(strain[:,:,3:]))),
          max_transverse_stress_MPa=float(np.max(abs(s[:,:,[0,1,3,4,5]]))),
          minimum_sigma_zz_MPa=float(np.min(s[:,:,2])),maximum_sigma_zz_MPa=float(np.max(s[:,:,2]))))
    return dict(phase=phase,states=len(fields),dofs=dofs,physical_dofs=physical,points=points,concrete_points=concrete,
      independent_equilibrium=proof,point_fields=measured,
      independent_material_integration=False,physical_capacity_certified=False,
      record_sha256=sha256((root/'Native'/(phase+'.csv')).read_bytes()).hexdigest().upper())


def TestCheck(root):
    """Сохраняет оба результата без объявления полного теста или несущей способности пройденными."""
    phases={name:TestAnalyse(root,name) for name in ('PostCrackPrescribed.Accepted','PostCrackContinuation.Accepted')}
    assert phases['PostCrackPrescribed.Accepted']['states']>=2 and phases['PostCrackContinuation.Accepted']['states']>=2
    proof=dict(phases=phases,all_saved_weak_equations_and_six_resultants_verified=True,
      capacity_certified=False,entire_goal_completed=False)
    p=root/'PostCrackFieldsProof.json';p.write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('FIELDS VERIFIED',[(k,v['states']) for k,v in phases.items()],'PROOF',sha256(p.read_bytes()).hexdigest().upper())
    for phase,data in phases.items():
        print(phase,'LAST',data['point_fields'][-1])
    return proof


if __name__=='__main__':
    p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True)
    TestCheck(p.parse_args().root.resolve())
