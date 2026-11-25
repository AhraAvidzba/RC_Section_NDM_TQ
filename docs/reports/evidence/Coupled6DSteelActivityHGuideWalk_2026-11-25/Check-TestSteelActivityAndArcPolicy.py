"""ДЛЯ ТЕСТОВ. Независимо проверяет пластическую активность стали и неизменность автоматического H-перехода."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold845SteelActivityArcPolicy_20261125';parent=base/'Cold844AutomaticHGuideArc_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
prep=json.loads((root/'Preparation.json').read_bytes())
assert all(sha256((root/n).read_bytes()).hexdigest().upper()==v for n,v in prep['inputs'].items())
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and len(audit['other_actual_components_exact_parent_except_identifier_case'])==252

def TestMethods(path):
    """Сохраняет точные тела методов без изменения условий и арифметики."""
    rows=re.findall(r'(?ms)^((?:Private|Friend|Public) (?:Function|Sub|Property Get) (\w+)\b.*?^End (?:Sub|Function|Property))',path.read_text('utf-8-sig'))
    assert len(rows)==len({name for _,name in rows});return {name:body for body,name in rows}

name='src/Solver/CCoupledNewtonSolver.cls';old,new=TestMethods(parent/name),TestMethods(root/name)
assert set(new)-set(old)=={'HasSteelPlasticIncrement'} and all(old[k]==new[k] for k in old)
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls','src/Materials/CCoupledConcreteReturn.cls',
             'src/Solver/CCoupledMaterialConstraint.cls','src/Solver/CCoupledEnvelopeEquations.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
name='src/Solver/CCoupledArcContinuation.cls';old,new=TestMethods(parent/name),TestMethods(root/name)
assert set(new)-set(old)=={'NextStepRatioLimit'}
assert {k for k in old if old[k]!=new[k]}=={'Initialize','TryStep','IsEnvelopeUnloading'}
guard='    If mSolver.HasSteelPlasticIncrement(field) Then Exit Function \' Замороженная H бетона не исключает пластическое продвижение стали.\n'
assert new['IsEnvelopeUnloading'].replace(guard,'',1)==old['IsEnvelopeUnloading']
name='src/Solver/CCoupledDeformationContinuation.cls';gated=base/'Cold831GatedLoadingArc_20261125'
code=(root/name).read_text('utf-8-sig');prior=(gated/name).read_text('utf-8-sig')
assert code.replace('GeomMin(arc.NextStepRatioLimit, 0.8 * Sqr(0.005 / GeomMax(arc.FieldError, 0.00000001)))',
                    'GeomMin(2#, 0.8 * Sqr(0.005 / GeomMax(arc.FieldError, 0.00000001)))',1)==prior

Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
reader=Reader((root/'Input/SteelActivity.Fields.bin').read_bytes());reader.position=0
count=int(reader.VectorTest()[0]);fields=np.array([reader.VectorTest() for _ in range(count)]);assert reader.position==len(reader.raw)
labels=json.loads((root/'Input/SteelActivity.Cases.json').read_bytes());assert count==len(labels)==13
rows=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestRows']
g,L,w,v,coords,axis,b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
operators=np.array([b+s*.001*gradient for s in (-1,0,1)])
strain=np.einsum('spcd,fd->fspc',operators,fields)
axial=np.einsum('fspc,pc->fsp',strain,axis)[:,:,25:]
longitudinal=axis[25:,2]==1
fy_t=np.where(longitudinal,400.,240.);fy_c=np.where(longitudinal,350.,240.)
expected=np.where(200000*axial>fy_t,1,np.where(200000*axial < -fy_c,-1,0))
driver=(root/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
for text in ('Value2 = 350# / units.InputStressToInternal(1#)','Value2 = 400# / units.InputStressToInternal(1#)',
             'grades.Cells(2, 3).Value2 = 200000#: grades.Cells(2, 4).Value2 = 240#: grades.Cells(2, 5).Value2 = 240#'):
    assert text in driver
# Принятая история этого опыта упругая по стали: критерий можно восстановить без скрытой пластической истории.
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
history_axial=np.einsum('fspc,pc->fsp',np.einsum('spcd,fd->fspc',operators,source['fields']),axis)[:,:,25:]
assert np.all(200000*history_axial<=fy_t) and np.all(200000*history_axial>=-fy_c)
seen=np.zeros((count,3,18),int);cases={}
for row in rows(root/'Native/SteelActivity.csv'):
    if row[0]=='CASE':
        case=int(row[1])-1;assert row[3]==row[4];actual=row[2]=='True';assert row[2] in ('True','False')
        assert actual==bool(np.any(expected[case]));cases[case]=actual
    else:
        assert row[0]=='BAR';case,side,point,branch=map(int,row[1:5]);case-=1;side+=1;point-=26
        assert 0<=point<18 and branch==expected[case,side,point]
        assert abs(float(row[5])-axial[case,side,point])<=2e-15
        seen[case,side,point]+=1
assert np.all(seen==1) and len(cases)==13
for label,side in (('LongPlusOnly',2),('LongMinusOnly',0)):
    index=next(i for i,x in enumerate(labels) if x['label']==label)
    assert np.any(expected[index,side]) and not np.any(expected[index,1]) and cases[index]
none=rows(root/'Native/SteelActivity.None.csv');assert len(none)==1 and none[0][0]=='NONE'
assert int(none[0][1])==int(none[0][2]) and none[0][3:]==['False','False']
trained=rows(root/'Native/SteelActivity.Trained.csv');assert len(trained)==4 and trained[0][0]=='TRAIN'
force,actual,area,bararea,stress,plastic,backstress,revision=map(float,trained[0][1:])
assert abs(area-180*120)<1e-8 and abs(bararea-np.pi*36)<1e-11
assert abs(stress-401.5)<1e-10 and abs(plastic-(.003-401.5/200000))<1e-15 and abs(backstress-1.5)<1e-10
reference_force=(area-2*bararea)*1000*.003+2*bararea*401.5
assert abs(force-reference_force)<1e-7 and abs(actual-force)<1e-5
for row,level in zip(trained[1:],(.0025,.0035,-.003)):
    assert row[0]=='READ' and float(row[1])==level and float(row[3])==float(row[4])==revision
    relative=200000*(level-plastic)-backstress
    assert bool(int(float(row[2])))==bool(relative>400 or relative < -350)

def TestReadGuide(path):
    """Читает только полные автоматические записи со всеми битами Double."""
    r=Reader(path.read_bytes());assert r.raw[:16]==b'NDMAUTOH00000001';records={}
    while r.position<len(r.raw):
        meta=r.VectorTest();packet=[r.VectorTest() for _ in range(15)]
        field,actual,projection,effort,present=[r.VectorTest() for _ in range(5)]
        guide=[r.VectorTest() for _ in range(3)] if present[0] else None
        records[int(meta[0])]=dict(meta=meta,packet=packet,field=field,actual=actual,projection=projection,effort=effort,present=present,guide=guide)
    assert r.position==len(r.raw);return records

records=TestReadGuide(root/'Native/AutomaticHGuide.bin');assert set(records)=={3}
current=records[3];previous=TestReadGuide(parent/'Native/AutomaticHGuide.bin')[3]
for key in ('meta','field','actual','projection','effort','present'):
    np.testing.assert_array_equal(current[key],previous[key])
for key in ('packet','guide'):
    for a,b in zip(current[key],previous[key]):np.testing.assert_array_equal(a,b)
info=current['packet'][0];guide=current['guide'];assert info[18]==6 and info[28]==2 and current['meta'][2]==1
assert info[16]<=30 and info[17]<=30 and info[19]==info[16]+info[17] and current['effort'][0]==0
assert rows(root/'Native/ArcNextRatio.csv')==[['LIMIT','3','1']]
mechanics,M=run_path(str(here/'TestUnacceptedMechanics.py'))['TestUnacceptedRoots'](root,source,{
    31:dict(field=current['field'],lambda_value=float(info[23]),accuracy=.0001,actual=current['actual']),
    32:dict(field=current['packet'][11],lambda_value=float(info[24]),accuracy=.0001),
    33:dict(field=guide[2],lambda_value=float(guide[0][6]),accuracy=.0001)})
proof=dict(native=native,actual_components=256,imported_components=4,other_components_verified=252,
    kinematic_cases=13,individual_station_bar_branches_verified=int(seen.sum()),both_bar_families=True,
    noncentral_only_plasticity_detected=True,accepted_history_stresses_fields_and_revision_unchanged=True,
    trained_plastic_NDS=dict(force_N=actual,stress_MPa=stress,plastic_strain=plastic,backstress_MPa=backstress,
                            elastic_unloading_inactive=True,further_and_reverse_plasticity_active=True),
    no_steel_inactive=True,original_Newton_methods_exact=True,material_laws_exact=True,
    automatic_Hguide_certificate_bit_identical_parent=True,phi_boundary_next_ratio_limit=1,
    actual_returned_event_lambda=float(info[23]),actual_probe_calls=int(info[18]),mechanics=mechanics,
    physical_capacity_certified=False,actual_cold_integration_not_yet_run=True,whole_goal_completed=False)
(root/'SteelActivityArcPolicyProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('STEEL ACTIVITY VERIFIED:13 FIELDS,702 BAR/STATION BRANCHES,TRUE PLASTIC HISTORY,NO STEEL;AUTOMATIC H EVENT BIT IDENTICAL')
print(json.dumps(proof,ensure_ascii=False))
