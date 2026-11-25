"""ДЛЯ ТЕСТОВ. Проверяет защиту точного контекста и сохранность всех физических проверок изменённого контроллера."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--gated',action='store_true');a=p.parse_args()
root=base/('Cold831GatedLoadingArc_20261125' if a.gated else 'Cold827RetainedArcControl_20261125');parent=base/'Cold824AutomaticLoadingCold_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
for name in ('src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledNewtonSolver.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls','Input/Reverse1.bin'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
name='src/Solver/CCoupledDeformationContinuation.cls'
before=(parent/name).read_text('utf-8-sig');after=(root/name).read_text('utf-8-sig')
def TestMethods(code):
    """Разбирает фактические тела методов VBA, не нормализуя условия или арифметику."""
    rows=re.findall(r'(?ms)^((?:Private|Friend|Public) (?:Function|Sub|Property Get) (\w+)\b.*?^End (?:Sub|Function|Property))',code)
    assert len(rows)==len({n for _,n in rows})
    return {name:body for body,name in rows}
old,new=TestMethods(before),TestMethods(after)
added={'CanRetainSignedArcControl','TestRetainedArcContext'}
if a.gated:added|={'CanStartRecentLoadingArc','TestProactiveLoadingContext'}
assert set(new)==set(old)|added
changed={name for name in old if old[name]!=new[name]}
assert changed=={'Trace','TryArcContinuationProbe','Details'}
anchor='        If Not probed Then\n'
assert old['Trace'].count(anchor)==new['Trace'].count(anchor)==1
assert old['Trace'].split(anchor)[1]==new['Trace'].split(anchor)[1]
assert 'mProbeCount = mProbeCount + arc.Calls' in new['TryArcContinuationProbe']
assert 'mProbeCount + 3 > PROBE_LIMIT' in new['TryArcContinuationProbe']
assert 'mSignedArcAttempts >= 4' in new['TryArcContinuationProbe']
assert 'mSignedArcExpectedRevision = mSolver.AcceptedRevision + 1' in new['TryArcContinuationProbe']
validate=run_path(str(here/'Validate-TestHardPrepared.py'))
rows=validate['TestRows'](root/'Native/RetainedArc.Context.csv');assert len(rows)==(8 if a.gated else 7)
for mode,row in enumerate(rows):
    assert row[0]=='CONTEXT';values=list(map(int,row[1:]))
    assert values[0]==mode and values[1]==int(mode==0) and values[2]==45 and values[3]==45-int(mode==1)
if a.gated:
    rows=validate['TestRows'](root/'Native/ProactiveLoading.Context.csv');assert len(rows)==7
    for mode,row in enumerate(rows):
        assert row[0]=='PROACTIVE' and list(map(int,row[1:]))==[mode,int(mode==0),45]
    assert 'continuingOutgoing Or CLng(status(28)) = 4' in new['TryArcContinuationProbe']
    assert 'Not mSignedArcOutgoing' in new['CanRetainSignedArcControl']
mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted');assert len(fields)==43
for name in ('Native/EnvelopeEnergy.Accepted.csv','Native/EnvelopeEnergy.Accepted.Operators.csv'):
    assert (root/name).read_bytes()==(base/'Cold823AutomaticLoadingWalk_20261125'/name).read_bytes(),name
proof=dict(native=native,source43_equilibria_independently_rebuilt=mechanics,
    all_original_material_and_numerical_kernel_methods_exact824=True,
    unchanged_controller_methods=sorted(set(old)-changed),changed_controller_methods=sorted(changed),
    all_physical_and_peak_steel_checks_after_probe_exact824=True,
    exact_context_positive_and_six_negative_controls=True,source_field_and_revision_unchanged_by_classifier=True,
    gated_after_verified_outgoing_step=a.gated,proactive_unique_boundary_contexts7=a.gated,
    total_context_cases=(15 if a.gated else 7),
    global96_and_pair30_contracts_preserved=True,actual_retained_continuation_not_yet_run=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'RetainedArcContextProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'CONTEXT CASES',15 if a.gated else 7,'UNCHANGED CONTROLLER METHODS',len(old)-len(changed),';ALL PHYSICAL ACCEPTANCE EXACT824')
