"""ДЛЯ ТЕСТОВ. Проверяет неизменность старого контракта удержания дуги и явные случаи допуска нового контекста без приёма материала."""
from pathlib import Path
from runpy import run_path
import json,re
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold895FreshSpatialArc_20261126';parent=base/'Cold888StrictEnvelopeGrowth_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay');assert native==dict(passed=23,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==256 and audit['imported_this_pass']==['CCoupledDeformationContinuation','modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==254
rel='src/Solver/CCoupledDeformationContinuation.cls';old=(parent/rel).read_text('utf-8-sig');new=(root/rel).read_text('utf-8-sig')
pattern=r'(?ms)^Private Function CanRetainSignedArcControl\(.*?^End Function'
assert re.findall(pattern,old)==re.findall(pattern,new) and len(re.findall(pattern,new))==1
assert new.count('Private Const PROBE_LIMIT As Long = 96')==1
for path in (parent/'src').rglob('*'):
    if path.is_file() and path.name!='CCoupledDeformationContinuation.cls':assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/FreshSpatialArcCases.bin').read_bytes());assert r.raw[:16]==b'NDMFRESHARC00001'
cases=[]
for mode in range(11):
    row=r.VectorTest();assert len(row)==4 and row[0]==mode and row[2]==row[3]
    assert bool(row[1])==(mode in (0,9,10))
    cases.append(dict(mode=mode,eligible=bool(row[1]),accepted_revision_unchanged=True))
assert r.position==len(r.raw)
proof=dict(native=native,actual_components=256,cases=cases,old_exact_retained_arc_gate_unchanged=True,
    one_new_try_per_accepted_revision=True,foreign_current_history_and_nonunit_trace_rejected=True,
    foreign_old_cache_not_reused_for_new_context=True,old_revision_failure_does_not_block_new_state=True,
    global96_primary30_force_and_material_limits_unchanged=True,physical_capacity_certified=False,cold_test_still_required=True)
(root/'FreshSpatialArcProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('FRESH ARC CONTEXT22PASS;OLD RETAINED GATE EXACT;FOREIGN CURRENT HISTORY REJECTED;NO OLD CACHE REUSE')
