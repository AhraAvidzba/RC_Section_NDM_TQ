"""ДЛЯ ТЕСТОВ. Независимо воспроизводит выбор численной меры по тензорным операторам и проверяет неизменность рабочего решателя."""
from pathlib import Path
from runpy import run_path
import json,re,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold855BalancedInitialControl_20261125';parent=base/'Cold851LoadingSecantReparameterization_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls','src/Materials/CCoupledConcreteReturn.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
def TestMethods(path):
    """Сохраняет точные тела методов; сравнение не меняет числа, допуски и условия."""
    rows=re.findall(r'(?ms)^((?:Private|Friend|Public) (?:Function|Sub|Property Get) (\w+)\b.*?^End (?:Sub|Function|Property))',path.read_text('utf-8-sig'))
    assert len(rows)==len({n for _,n in rows});return {n:s for s,n in rows}
path='src/Solver/CCoupledDeformationContinuation.cls';a,b=TestMethods(parent/path),TestMethods(root/path)
assert set(b)-set(a)=={'SelectInitialDeformationControl','TestInitialControlSelection'}
changed={n for n in a if a[n]!=b[n]};assert changed=={'Trace'}
old='''    If mSolver.UsesSpatialHistory Then selected = mSolver.SelectTransverseDeformationControl(previousField, field, _
        mDirection, mProjection, current.Coordinate, nominal, mControlName)
    mUsesTransverseControl = selected
    If Not selected Then selected = mSolver.SelectDeformationControl(previousField, field, mProjection, current.Coordinate, nominal, mControlName)'''
assert b['Trace'].replace('    selected = SelectInitialDeformationControl(previousField, field, current.Coordinate, nominal)',old,1)==a['Trace']
code=(root/path).read_text('utf-8-sig');assert 'Private Const PROBE_LIMIT As Long = 96' in code
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
reader=Reader((root/'Input/InitialControl.Fields.bin').read_bytes());reader.position=0
count=int(reader.VectorTest()[0]);cases=[(reader.VectorTest(),reader.VectorTest()) for _ in range(count)];assert reader.position==len(reader.raw)
labels=json.loads((root/'Input/InitialControl.Cases.json').read_bytes());assert len(cases)==len(labels)==7
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
mean=np.average(B[:25,4:6],axis=0,weights=coords[:25,2])
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
direction=source['direction'][4:];direction=direction/np.linalg.norm(direction)
qprojection=direction@mean
data=Reader((root/'Native/InitialControl.bin').read_bytes());assert data.raw[:16]==b'NDMINITCTRL00001'
results=[]
for index,(previous,current) in enumerate(cases,1):
    meta,projection=data.VectorTest(),data.VectorTest();assert meta[0]==index and meta[5]==meta[6]
    delta=current-previous;changes=np.append(delta[:4],mean@delta);chosen=int(np.argmax(abs(changes)))
    general_spacing=float(abs(changes[chosen]));qchange=float(qprojection@delta);transverse_spacing=abs(qchange)
    expected_transverse=transverse_spacing>0 and transverse_spacing>=general_spacing
    expected_ready=max(general_spacing,transverse_spacing)>0
    assert bool(meta[1])==expected_ready and bool(meta[2])==expected_transverse
    if expected_ready:
        if expected_transverse:expected=np.sign(qchange)*qprojection;spacing=transverse_spacing
        else:
            expected=np.zeros(60)
            if chosen<4:expected[chosen]=np.sign(changes[chosen])
            else:expected=np.sign(changes[chosen])*mean[chosen-4]
            spacing=general_spacing
        assert max(abs(expected-projection))<1e-13
        assert abs(meta[3]-expected@current)<1e-17 and abs(meta[4]-spacing)<1e-17
        assert np.array_equal(projection[56:],np.zeros(4))
    else:assert len(projection)==1
    results.append(dict(case=index,label=labels[index-1]['label'],selected=expected_ready,transverse=expected_transverse,
        largest_component=chosen+1,general_spacing=general_spacing,transverse_spacing=transverse_spacing))
assert data.position==len(data.raw)
assert not results[0]['transverse'] and results[0]['largest_component']==4
assert results[2]['transverse'] and results[-1]['transverse'] and not results[1]['selected']
proof=dict(native=native,only_Trace_initial_selection_call_changed=True,all_other_working_methods_and_materials_exact851=True,
    actual_original96_and30_budgets=True,accepted_fields_and_revision_unchanged=True,cases=results,
    no_new_nonlinear_solution=True,cold_workflow_not_yet_run=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'BalancedInitialControlProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('SEVEN INITIAL CONTROL CASES VERIFIED;ACTUAL HARD SHEAR SELECTS TWIST;PURE/MIXED DOMINANT SHEAR RETAINS Q;HISTORY UNCHANGED')
for item in results:print(item)
