"""ДЛЯ ТЕСТОВ. Проверяет материальные префиксы971, все уравнения и независимую полную/половинную метрику без их принятия."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=base/'Cold971MaterialPrefixCoordinate_20261126')
parser.add_argument('--unloading',action='store_true');parser.add_argument('--outgoing-pair',action='store_true');args=parser.parse_args()
assert not (args.unloading and args.outgoing_pair)
root=args.root;parent=base/'Cold968PrefixNewtonCapture_20261126'
code=(here/'Check-TestRecordedYieldCoordinate.py').read_text('utf-8-sig')
changes={
    "root=base/'Cold936RecordedYieldCoordinate_20261126';parent=base/'Cold935Point11Matrices_20261126'":
        'root=Path('+repr(str(root))+');parent=Path('+repr(str(parent))+')',
    "audit['imported_this_pass']==['modTestCoupledElasticState']":"audit['imported_this_pass']==['CCoupledNewtonSolver']",
    'source_actual934_state25_and_initial_residual_bitexact_by_native_guard=True':'source_actual968_state25_and_initial_residual_bitexact_by_native_guard=True',
}
for old,new in changes.items():assert code.count(old)==1;code=code.replace(old,new,1)
env={'__file__':str(here/'Check-TestRecordedYieldCoordinate.py'),'__name__':'material_prefix_existing_entry_test'}
exec(compile(code,str(here/'Check-TestRecordedYieldCoordinate.py'),'exec'),env)
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
for folder in ('src','tests'):
    for path in (parent/folder).rglob('*'):
        if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
origin=json.loads((root/'MaterialPrefixOrigin.json').read_bytes());source=env['source'];R=env['R'];op=env['op'];coords=env['coords']
assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
alpha=json.loads((root/'SourceAutomatic962ArcProof.json').read_bytes())['info'][9]

def TestLength(delta):
    """Восстанавливает исходную тензорную меру всех129 точек трёх станций без четырёх служебных множителей."""
    strain=np.einsum('spcd,d->spc',op,delta)
    return float(np.sqrt(np.einsum('spc,spc,p,c->',strain,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]))))

r=R((root/'Native/RecordedYieldCoordinate.bin').read_bytes());assert r.VectorTest().tolist()==[4.]
fields={};infos={}
for index in range(1,5):
    info=r.VectorTest();infos[index]=info
    if info[5]:
        field,values,forces=[r.VectorTest() for _ in range(3)]
        for _ in range(129):r.VectorTest();r.VectorTest()
        fields[index]=field
assert r.position==len(r.raw)
pairs=[];reference=float(source['lambdas'][-1]);accepted=source['fields'][-1]
for full_index,half_index in (() if args.unloading else ((1,3),(2,4))):
    full_info=infos[full_index];half_info=infos[half_index]
    item=dict(full_case=full_index,half_case=half_index,complete_Newton_line_search=full_index==1,
        full_plus_half_steps=int(full_info[10]+half_info[10]),within30_total_steps=bool(full_info[10]+half_info[10]<=30),
        automatic_station_selection_cost_excluded_here=True,accepted_by_native_full_half_gate=False)
    if env['cases'][full_index-1]['independently_equilibrated'] and env['cases'][half_index-1]['independently_equilibrated']:
        full,half=fields[full_index],fields[half_index];length=TestLength(full-accepted)
        error=TestLength(full-(2*half-accepted))/length;lambda_error=float(abs(2*half_info[6]-reference-full_info[6]))
        radius=float(np.hypot(length,alpha*(full_info[6]-reference)));allowance=(.005*radius)**2-(length*error)**2
        noise=float(8*source['options'][4]*(1+abs(reference)))
        bound=max(noise,float(np.sqrt(allowance)/alpha) if allowance>0 else noise)
        item.update(field_error=error,lambda_error=lambda_error,lambda_bound=bound,
            joint_error=float(np.hypot(length*error,alpha*lambda_error)/radius),
            material_target_half_of_original_increment=bool(abs(half_info[1]-.5*(origin['accepted_coordinate_MPa']+full_info[1]))<1e-16),
            passes_independent_existing_shape_and_lambda_error_limits=bool(error<=.005 and lambda_error<=bound),
            field_length=length,full_lambda=float(full_info[6]),half_lambda=float(half_info[6]))
    pairs.append(item)
proof=env['proof'];proof.update(all257_actual_components_and_saved_book_bitexact968=True,no_new_numerical_code_or_compilation=True,
    actual_imports_this_pass=[],point18_selected_from_program962_event=not (args.unloading or args.outgoing_pair),
    station_selected_from_actual968_unloading_crossing=bool(args.unloading or args.outgoing_pair),selection_not_rerun_inside_this_entry=True,
    no_future_equilibrium_field_in_consumed_fixture=True,starting_field_is_actual_failed968_iteration=True,
    full_fraction_to_selected_surface=origin.get('fraction_to_zero'),material_full_half_pairs=pairs,
    boundary_and_unloading_cases_are_not_full_half_pair=bool(args.unloading),
    native_lambda_uncertainty_not_recorded_in_existing_entry=True,full_half_not_jointly_accepted_or_counted_as_production_recovery=True,
    physical_capacity_certified=False,cold_gain_claimed=False)
(root/'MaterialPrefixCoordinateProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('MATERIAL PREFIX CASES VERIFIED;PAIRS',pairs)
