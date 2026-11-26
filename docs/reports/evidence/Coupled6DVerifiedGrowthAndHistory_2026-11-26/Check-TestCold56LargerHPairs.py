"""ДЛЯ ТЕСТОВ. Проверяет увеличенные H-пары холодного состояния56 в исходных строгих допусках без принятия неподтверждённых корней."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold886Cold56LargerHPairs_20261126';parent=base/'Cold881SpatialSteelLimits_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];source=read(root/'Input/Reverse1.bin');assert len(source['fields'])==56
frames,_=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Input/SourceColdRawHistory.bin')
for key in ('fields','fractions','lambdas'):np.testing.assert_array_equal(source[key],frames[-1][key])
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)]);E=source['calibration'][0]
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
shared=(here/'Check-TestSharedCurveHControl.py').read_text('utf-8-sig');begin=shared.index('def TestLength(');end=shared.index("r=R((root/'Native/HCornerPairs.bin').read_bytes())",begin)
exec(compile(shared[begin:end],str(here/'Check-TestSharedCurveHControl.py'),'exec'),globals())
accepted=source['fields'][-1];incoming=accepted-source['fields'][-2]
file=here/'Check-TestVerifiedHCornerPairs.py';code=file.read_text('utf-8-sig')
begin=code.index("r=R((root/'Native/VerifiedHPairs.bin').read_bytes())");end=code.index('proof=dict(',begin);code=code[begin:end]
for old,new in [
 ('Native/VerifiedHPairs.bin','Native/Cold56LargerHPairs.bin'),('header[0]==20','header[0]==56'),
 ("TestH('EnvelopeEnergy.Accepted',20)","TestH('EnvelopeEnergy.Accepted',56)"),("TestH('EnvelopeEnergy.Accepted',19)","TestH('EnvelopeEnergy.Accepted',55)"),
 ('range(1,4)','range(1,7)'),('meta[1]==.25**(index-1)','meta[1]==2**(index-1)'),
 ("phase='VerifiedH.Case'","phase='Cold56H.Case'"),('len(fields)==21','len(fields)==57'),
 ('fields[:20]','fields[:56]'),('TestH(phase,21)','TestH(phase,57)')]:
    assert old in code,old;code=code.replace(old,new)
exec(compile(code,str(file),'exec'),globals())
proof=dict(native=native,source_states=56,source_is_exact_actual_cold880=True,actual_component_count=256,
    all_working_sources_exact881=True,cases=cases,verified_pairs=count,independently_rebuilt_roots=2*count,
    global_working_continuation_unchanged=True,strict_force_and_material_targets_and30iterations_unchanged=True,
    original_cold_history_equilibrium=mechanics,no_history_or_capacity_acceptance=True,physical_capacity_certified=False)
(root/'Cold56LargerHPairsProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('COLD56 STRICT LARGER H PAIRS',count,'CASES',[(x['ratio'],x['solved'],x['primary_iterations']) for x in cases],';NO PHYSICAL LIMIT')
