"""ДЛЯ ТЕСТОВ. Проверяет неизменность968 и направление линеаризации по фактическим пробам отклонённого Newton."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from argparse import ArgumentParser
import json,sys,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=base/'Cold968PrefixNewtonCapture_20261126')
parser.add_argument('--changed-floor',action='store_true');args=parser.parse_args()
root=args.root;parent=base/'Cold967LongerSecantPrefix_20261126'
sys.argv=[str(here/'Check-TestControlledSecantPrefix.py'),'--root',str(root),'--root-budget','10']
if args.changed_floor:sys.argv+=['--spatial-floor']
run_path(str(here/'Check-TestControlledSecantPrefix.py'),run_name='prefix_capture_baseline_test')
if not args.changed_floor:assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
for folder in ('src','tests'):
    for path in (parent/folder).rglob('*'):
        if path.is_file() and not (args.changed_floor and path.name=='CCoupledNewtonSolver.cls'):
            assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
unchanged_inputs=['Input/Reverse1.bin','Input/ControlledPrefix.bin']
if not args.changed_floor:unchanged_inputs+=['Native/ControlledPrefix.Result.bin']
for name in unchanged_inputs:
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
read_source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
source=read_source(root/'Input/Reverse1.bin')
scan=base/'Cold930FineLinePlusSurface_20261126';old_source=read_source(scan/'Input/Reverse1.bin')
for key in ('calibration','offset','direction','fields','lambdas'):np.testing.assert_array_equal(source[key],old_source[key])
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
state_bytes=(scan/'Native/ArcScan.States.bin').read_bytes();history_file=root/'SourceAcceptedSurfaceStates.bin'
if history_file.exists():assert history_file.read_bytes()==state_bytes
else:history_file.write_bytes(state_bytes)
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];states={};r=R(state_bytes)
for _ in range(129):
    point,side=r.VectorTest();states[int(point),int(side)]=r.VectorTest()
assert r.position==len(r.raw)
helper=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig')
begin=helper.index('def TestYield(');end=helper.index('def TestAssemble(',begin)
exec(compile(helper[begin:end],str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
records=[]
for path in sorted((root/'Native').glob('NewtonRejected.*.bin')):
    r=R(path.read_bytes());r.position=0
    meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
    matrix=np.array([r.VectorTest() for _ in range(61)]);ending=r.VectorTest();count=int(r.VectorTest()[0])
    probes=[[r.VectorTest() for _ in range(3)] for _ in range(count)];assert r.position==len(r.raw)
    np.testing.assert_array_equal(accepted,source['fields'][-1]);assert len(meta)==9 and matrix.shape==(61,61)
    np.testing.assert_allclose(rhs[:60],-residual,rtol=0,atol=0)
    linear=matrix@delta;linear_defect=float(max(abs(linear-rhs)))
    row_scale=np.ones(61);row_scale[56:60]=meta[4]
    origin=-rhs;items=[];initial_surfaces={key:TestYield(*key,current) for key in states}
    for p,field,actual_residual in probes:
        fraction=float(p[0]);assert 0<fraction<=1 and len(p)==6
        np.testing.assert_allclose(field,current+fraction*delta[:60],rtol=2e-14,atol=2e-17)
        assert abs(p[1]-(meta[2]+fraction*delta[-1]))<1e-13
        actual=np.r_[actual_residual,p[4]];predicted=origin+fraction*linear
        actual_derivative=(actual-origin)/fraction;defect=(actual_derivative-linear)*row_scale
        worst=int(np.argmax(abs(defect)))
        physical_error=float(max(abs(defect[:56])));physical_size=float(max(abs(linear[:56])))
        surfaces={key:TestYield(*key,field) for key in states}
        crossings=[dict(point=key[0],side=key[1],from_value_MPa=initial_surfaces[key],to_value_MPa=value)
            for key,value in surfaces.items() if initial_surfaces[key]*value<0]
        items.append(dict(fraction=fraction,recorded_equilibrium_norm_MPa=float(p[2]),recorded_tolerance_MPa=float(p[3]),
            accepted=bool(p[5]),actual_physical_residual_inf_MPa=float(max(abs(actual[:56]))),
            predicted_physical_residual_inf_MPa=float(max(abs(predicted[:56]))),
            directional_derivative_error_MPa=physical_error,linear_physical_direction_norm_MPa=physical_size,
            relative_directional_derivative_error=physical_error/max(physical_size,1e-300),worst_equation_1based=worst+1,
            elastic_predictor_loading_surface_sign_changes=crossings))
    column=np.maximum(np.max(abs(matrix),axis=0),1e-300);balanced=matrix/column
    rows=np.maximum(np.max(abs(balanced),axis=1),1e-300);singular=np.linalg.svd(balanced/rows[:,None],compute_uv=False)
    boundaries=[]
    if probes:
        p,field,_=probes[-1]
        for key,value in initial_surfaces.items():
            if value*TestYield(*key,field)>=0:continue
            lo=0.;hi=float(p[0]);from_positive=value>0
            for _ in range(64):
                mid=.5*(lo+hi)
                if mid==lo or mid==hi:break
                evaluated=TestYield(*key,current+mid*delta[:60])
                if (evaluated>0)==from_positive:lo=mid
                else:hi=mid
            boundaries.append(dict(point=key[0],side=key[1],bracket=[lo,hi],
                from_value_MPa=value,first_recorded_fraction=float(p[0]),
                boundary_is_loading_switch_not_strength_failure=True))
    nearest=min(initial_surfaces,key=lambda key:abs(initial_surfaces[key]))
    records.append(dict(file=path.name,sha256=sha256(path.read_bytes()).hexdigest().upper(),metadata=meta.tolist(),ending=ending.tolist(),
        original_accepted25_field_exact=True,linear_solve_defect_MPa=linear_defect,
        balanced_singular_ratio=float(singular[-1]/singular[0]),
        nearest_elastic_predictor_loading_surface=dict(point=nearest[0],side=nearest[1],value_MPa=initial_surfaces[nearest]),
        loading_switch_below_smallest_recorded_fraction=boundaries,probes=items))
proof=json.loads((root/'ControlledSecantPrefixProof.json').read_bytes())
proof.update(all257_sources_and_saved_book_bitexact967=not args.changed_floor,returned_binary_bitexact967=not args.changed_floor,
    all256_nonnewton_components_bitexact967=True,spatial_floor_explicitly_changed=bool(args.changed_floor),
    independent_surface_states_sha256=sha256(state_bytes).hexdigest().upper(),independent_surface_states_origin=str(scan/'Native/ArcScan.States.bin'),
    independent_surface_source_history25_bitexact_current=True,
    capture_adds_no_material_or_numerical_call=True,rejected_newton_records=records,
    directional_difference_not_full_independent_jacobian=True,physical_capacity_certified=False)
(root/'PrefixNewtonCaptureProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('SPATIAL FLOOR VARIANT;OTHER256 PRESERVED' if args.changed_floor else 'IMMUTABLE967 REPLAY BITEXACT',
    'REJECTED NEWTON RECORDS',len(records))
for record in records:
    last=record['probes'][-1] if record['probes'] else {}
    print(record['file'],'ITERATION',record['metadata'][1],'LAMBDA',record['metadata'][2],
        'LINEAR DEFECT',record['linear_solve_defect_MPa'],'LAST PROBE',last)
