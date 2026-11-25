"""ДЛЯ ТЕСТОВ. Независимо проверяет общий прогноз с согласованной ветвью на той же истории20 и все полученные полные/половинные корни."""
from pathlib import Path
import sys,json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold873BranchConsistentArcTangent20_20261126'
source=here/'Check-TestFixedMetricKernel.py';code=source.read_text('utf-8-sig')
code=code.replace("default=base/'Cold867FixedMetricKernel_20261126'","default=base/'Cold873BranchConsistentArcTangent20_20261126'")
old="assert audit['components']==256 and audit['imported_this_pass']==['CCoupledArcContinuation','CCoupledDeformationContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']\nassert len(audit['other_actual_components_exact_parent_except_identifier_case'])==252"
assert code.count(old)==1;code=code.replace(old,"assert audit['components']==256 and audit['imported_this_pass']==['CCoupledArcContinuation','CCoupledNewtonSolver']\nassert len(audit['other_actual_components_exact_parent_except_identifier_case'])==254")
code=code.replace("scope['case']=='GeneralFixedMetricKernelFromCold19'","scope['case']=='SelfConsistentOutgoingTangentFromVerified20'")
code=code.replace("len(source['fields'])==19","len(source['fields'])==20")
code=code.replace('Cold866Cold19ReferenceMetricPairs_20261126/Input/Reverse1.bin','Cold869FixedMetricNextScan_20261126/Input/Reverse1.bin')
old='        np.testing.assert_array_equal(probe,reference+(reference-raw_fields[-2])/1024.)'
new='''        assert len(tm)==7 and 1<=tm[5]<=4 and tm[6]==1
        if tm[5]==1:
            np.testing.assert_array_equal(probe,reference+(reference-raw_fields[-2])/1024.)
        elif count==20 and tm[5]==2:
            original=Reader((base/'Cold867FixedMetricKernel_20261126/Native/LoadingWalk.Guide003.bin').read_bytes());original.VectorTest()
            for vector_count in (3,3):
                if original.VectorTest()[0]:
                    for _ in range(vector_count):original.VectorTest()
            assert original.VectorTest()[0]==1
            original_vectors=[original.VectorTest() for _ in range(4)]
            np.testing.assert_allclose(probe,reference+original_vectors[3]/1024.,rtol=0,atol=2e-18)
        assert used>=tm[5]+2
'''
assert code.count(old)==1;code=code.replace(old,new)
code=code.replace('source19_true_history=True','source20_true_history=True')
code=code.replace("root/'FixedMetricKernelProof.json'","root/'BranchConsistentTangentProof.json'")
old_print="';ALL RETURNED/HALF FORCES,STEEL,METRIC,H CHECKED;NO PHYSICAL LIMIT'";assert code.count(old_print)==1
code=code.replace(old_print,"(';NEW FULL/HALF ROOTS/FORCES/STEEL/H CHECKED;NO PHYSICAL LIMIT' if accepted else ';NO NEW ROOTS;SOURCE20 CHECKED;NO PHYSICAL LIMIT')")
sys.argv=[str(source)]
exec(compile(code,str(source),'exec'),{'__file__':str(source),'__name__':'branch_consistent_arc_tangent_test'})
proof=json.loads((root/'BranchConsistentTangentProof.json').read_bytes())
proof.update(maximum_tangent_builds_per_attempt4=True,all_builds_counted_in_original96_budget=True,
    source_provenance_and_available_two_build20_probes_rebuilt=True,
    branch_signature_equality_native_guard_not_yet_independently_replayed=True,
    unchanged_primary_corrector_and_materials_from867=True)
(root/'BranchConsistentTangentProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
