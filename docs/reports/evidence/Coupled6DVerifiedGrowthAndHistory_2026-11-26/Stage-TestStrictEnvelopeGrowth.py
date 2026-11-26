"""ДЛЯ ТЕСТОВ. Исключает обход половинной проверки увеличенного H-шага вблизи численного разрешения координаты."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold887VerifiedEnvelopeGrowth_20261126';root=base/'Cold888StrictEnvelopeGrowth_20261126'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'StrictGrowthCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes())
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
old='            verifyGrowth = (point.Coordinate - after > MATERIAL_STEP_GROWTH * (after - before) + 4# * mEnvelopeResolution)'
new='            verifyGrowth = (point.Coordinate - after > MATERIAL_STEP_GROWTH * (after - before))'
assert code.count(old)==1;code=code.replace(old,new,1)
data=code.replace('\n','\r\n').encode('utf-8-sig');(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
after=1e-4;resolution=1e-12;before=after-2*resolution;target=after+4*resolution
old_gate=target-after>1.25*(after-before)+4*resolution;new_gate=target-after>1.25*(after-before)
assert not old_gate and new_gate
negative=dict(after=after,before=before,target=target,resolution=resolution,old_gate_bypasses_pair=not old_gate,new_gate_requires_pair=new_gate,
              original_kernel_still_requires_target_above_twice_resolution=True,no_floating_or_material_tolerance_added=True)
for name,data in {'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
                  'NegativeResolutionPairGate.json':(json.dumps(negative,sort_keys=True,indent=2)+'\n').encode(),Path(__file__).name:Path(__file__).read_bytes()}.items():
 (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='StrictVerifiedEnvelopeGrowth',pair_trigger_no_resolution_bypass=True,
    every_reinforced_H_increment_over_previous_times1point25_requires_pair=True,negative_candidate887_preserved_not_tested_or_promoted=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'STRICT PAIR TRIGGER WITHOUT H-RESOLUTION EXCEPTION;887COUNTEREXAMPLE PRESERVED')
