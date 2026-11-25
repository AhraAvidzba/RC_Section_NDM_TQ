"""ДЛЯ ТЕСТОВ. Сопоставляет неизменённый и уточнённый общий корректор на одной доказанной истории20 без ослабления приёмки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from argparse import ArgumentParser
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--control',action='store_true');args=p.parse_args()
parent=base/'Cold867FixedMetricKernel_20261126';source=base/'Cold869FixedMetricNextScan_20261126'
root=base/('Cold871UnchangedCorrectorControl20_20261126' if args.control else 'Cold870RefinedArcCorrector20_20261126')
proof=json.loads((source/'ColdAcceptedStressContinuityProof.json').read_bytes());assert proof['postcommit_stress_drift_within1e_minus11_MPa'] and proof['native']['own_process_closed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0,source/'Input/Reverse1.bin')
prep=json.loads((root/'Preparation.json').read_bytes());values={Path(__file__).name:Path(__file__).read_bytes(),
 'SourceVerified20ContinuityProof.json':(source/'ColdAcceptedStressContinuityProof.json').read_bytes()}
if not args.control:
    archive=root/'RefinedCorrectorCompilationAncestor';archive.mkdir()
    for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
    name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
    old='''        built = Evaluate(current, True, useGradientTangent, Not (mConsistentControl Or mRefinedControlTangent), _
            mConsistentControl, mRefinedControlTangent)'''
    new='''        built = Evaluate(current, True, useGradientTangent, Not (mConsistentControl Or mRefinedControlTangent), _
            mConsistentControl, (mRefinedControlTangent Or mConsistentControl)) ' Управляемая поправка использует то же уточнение производной, что исходящий прогноз; силовая невязка неизменна.'''
    assert code.count(old)==1;code=code.replace(old,new,1)
    values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
    compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');old="-ModuleNames 'modTestCoupledElasticState'";assert compiler.count(old)==1
    values['Compile-TestExternalPoint.ps1']=compiler.replace(old,"-ModuleNames 'CCoupledNewtonSolver'",1).replace('\n','\r\n').encode('utf-8-sig')
    values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
    values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
scope=dict(case='UnchangedControlFromVerified20' if args.control else 'RefinedPrimaryCorrectorFromVerified20',
 actual_import_this_pass=json.loads((parent/'docs/regression/CodeUpdate.json').read_bytes())['Modules'] if args.control else ['CCoupledNewtonSolver'],
 actual_imports_this_replay=[] if args.control else ['CCoupledNewtonSolver'],actual_component_count=256,
 parent=str(parent),source_verified20=str(source),source20_exact=True,only_primary_corrector_branch_difference_scale_changed=not args.control,
 material_stress_law_and_histories_unchanged=True,force_tolerances_unchanged=True,maximum_calls96=True,primary_phase_iterations30=True,
 same_existing_public_test_driver=True,no_future_field_used=True,physical_capacity_certified=False,no_release_promotion=True)
prep['directed_scope']=scope
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'CONTROL' if args.control else 'ONE PRIMARY CORRECTOR CALL CHANGED;REFINED DELTA/64;NO MATERIAL/TOLERANCE/BUDGET CHANGE')
