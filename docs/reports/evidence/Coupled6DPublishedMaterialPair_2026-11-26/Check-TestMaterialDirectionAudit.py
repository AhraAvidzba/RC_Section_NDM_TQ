"""ДЛЯ ТЕСТОВ. Проверяет H/сталь/пространственную ориентацию опубликованных материальных корней и пригодность направления."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json,sys,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=base/'Cold977MaterialDirectionAudit_20261126')
root=parser.parse_args().root
sys.argv=[str(here/'Check-TestPublishedMaterialPair.py'),'--root',str(root)]
run_path(str(here/'Check-TestPublishedMaterialPair.py'),run_name='material_direction_publication_test')
proof=json.loads((root/'PublishedMaterialPairProof.json').read_bytes())
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];source=read(root/'Input/Reverse1.bin')
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/MaterialDirection.bin').read_bytes())
assert r.raw[:16]==b'NDMMATDIREC00001';header=r.VectorTest();reference,revision,baseline,resolution=header
assert reference==source['lambdas'][-1] and baseline>=0 and resolution>0
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
tolerance=float(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+reference*source['direction']))[:56])))
assert abs(resolution-tolerance*source['calibration'][3]/source['calibration'][0])<1e-12*resolution
records=[];tag=r.VectorTest()
while tag[0]>0:
    assert len(tag)==15 and tag[13]==tag[14]==revision and tag[4]==baseline and tag[6]==resolution
    index=int(tag[0]);publication=proof['publication_records'][index-1]
    assert tag[1]==proof['cases'][index-1]['lambda_value']
    assert bool(tag[8])==(tag[9]>0)
    assert abs(tag[9]-publication['signed_spatial_advance_along_previous_secant'])<2e-12*max(abs(tag[9]),1e-20)+1e-21
    assert abs(tag[10]-publication['previous_spatial_increment'])<2e-12*tag[10]
    decreased=bool(tag[1]<reference-8*source['options'][4]*(1+abs(reference)));assert bool(tag[11])==decreased
    rejected=bool(tag[3] and baseline>0 and decreased and not tag[7] and tag[5]<=baseline)
    assert bool(tag[12])==rejected
    elastic=bool(publication.get('maximum_concrete_elastic_predictor_loading_value_MPa',1)<0)
    ordinary_unloading=bool(tag[3] and tag[5]==baseline and not tag[7] and elastic and decreased)
    material_advance=bool(tag[3] and tag[5]-baseline>resolution)
    item=dict(case=index,native_solved=bool(tag[2]),independently_equilibrated=proof['cases'][index-1]['independently_equilibrated'],
        lambda_value=float(tag[1]),H_available=bool(tag[3]),accepted_H=float(baseline),trial_H=float(tag[5]),H_increment=float(tag[5]-baseline),
        existing_H_resolution=float(resolution),resolved_H_growth=material_advance,steel_plastic_increment=bool(tag[7]),
        forward_along_previous_spatial_secant=bool(tag[8]),signed_spatial_advance=float(tag[9]),
        lambda_decrease_beyond_existing_noise=decreased,rejected_by_existing_unloading_rule=rejected,
        all_concrete_predictors_strictly_elastic=elastic,ordinary_elastic_unloading=ordinary_unloading,
        no_acceptance_or_capacity_claim=True)
    records.append(item);tag=r.VectorTest()
assert tag.tolist()==[0.,float(revision),float(revision)] and r.position==len(r.raw)
if root.name=='Cold978LoadingMaterialDirection_20261126':
    parent=base/'Cold977MaterialDirectionAudit_20261126'
    assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
    for folder in ('src','tests'):
        for p in (parent/folder).rglob('*'):
            if p.is_file():assert p.read_bytes()==(root/p.relative_to(parent)).read_bytes()
    old=R((parent/'Input/RecordedYieldSeed.bin').read_bytes());old.position=0
    old_meta,*old_vectors=[old.VectorTest() for _ in range(5)]
    new=R((root/'Input/RecordedYieldSeed.bin').read_bytes());new.position=0
    new_meta,*new_vectors=[new.VectorTest() for _ in range(5)]
    for a,b in zip(old_vectors,new_vectors):np.testing.assert_array_equal(a,b)
    for i in (0,1,3,4,5,6,7,8):assert old_meta[i]==new_meta[i]
    origin=json.loads((root/'MaterialPrefixOrigin.json').read_bytes());value=origin['accepted_coordinate_MPa']
    assert new_meta[2]==value-(old_meta[2]-value) and new_meta[9]==.5*(value+new_meta[2])
proof.update(direction_records=records,ordinary_elastic_unloading_cases=sum(x['ordinary_elastic_unloading'] for x in records),
    equilibrated_cases_with_resolved_H_growth=sum(x['independently_equilibrated'] and x['resolved_H_growth'] for x in records),
    existing_unloading_guard_not_relaxed=True,no_material_history_acceptance=True,
    physical_capacity_certified=False,cold_gain_claimed=False,whole_goal_completed=False)
(root/'MaterialDirectionProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('MATERIAL DIRECTION AUDIT VERIFIED',records)
