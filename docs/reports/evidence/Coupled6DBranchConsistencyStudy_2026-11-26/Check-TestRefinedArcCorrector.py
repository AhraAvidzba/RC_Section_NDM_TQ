"""ДЛЯ ТЕСТОВ. Применяет исходный независимый контроль полной/половинной ветви к одинаковой истории20 и единственному изменению корректора."""
from pathlib import Path
from argparse import ArgumentParser
from hashlib import sha256
import sys,json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--control',action='store_true');a=p.parse_args()
root=base/('Cold871UnchangedCorrectorControl20_20261126' if a.control else 'Cold870RefinedArcCorrector20_20261126');parent=base/'Cold867FixedMetricKernel_20261126'
old='''        built = Evaluate(current, True, useGradientTangent, Not (mConsistentControl Or mRefinedControlTangent), _
            mConsistentControl, mRefinedControlTangent)'''
new='''        built = Evaluate(current, True, useGradientTangent, Not (mConsistentControl Or mRefinedControlTangent), _
            mConsistentControl, (mRefinedControlTangent Or mConsistentControl)) ' Управляемая поправка использует то же уточнение производной, что исходящий прогноз; силовая невязка неизменна.'''
if a.control:
    assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
for f in (parent/'src').rglob('*'):
    if not f.is_file():continue
    name=f.relative_to(parent);before=f.read_bytes();after=(root/name).read_bytes()
    if not a.control and name.as_posix()=='src/Solver/CCoupledNewtonSolver.cls':
        code=after.decode('utf-8-sig').replace('\r\n','\n');assert code.count(new)==1
        assert code.replace(new,old,1)==before.decode('utf-8-sig').replace('\r\n','\n')
    else:assert before==after,name
source=here/'Check-TestFixedMetricKernel.py';code=source.read_text('utf-8-sig')
code=code.replace("default=base/'Cold867FixedMetricKernel_20261126'","default=base/"+repr(root.name))
if not a.control:
    old_audit="assert audit['components']==256 and audit['imported_this_pass']==['CCoupledArcContinuation','CCoupledDeformationContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']\nassert len(audit['other_actual_components_exact_parent_except_identifier_case'])==252"
    assert code.count(old_audit)==1
    code=code.replace(old_audit,"assert audit['components']==256 and audit['imported_this_pass']==['CCoupledNewtonSolver']\nassert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255")
code=code.replace("scope['case']=='GeneralFixedMetricKernelFromCold19'","scope['case']=="+repr('UnchangedControlFromVerified20' if a.control else 'RefinedPrimaryCorrectorFromVerified20'))
assert code.count("len(source['fields'])==19")==1;code=code.replace("len(source['fields'])==19","len(source['fields'])==20")
code=code.replace('Cold866Cold19ReferenceMetricPairs_20261126/Input/Reverse1.bin','Cold869FixedMetricNextScan_20261126/Input/Reverse1.bin')
code=code.replace('source19_true_history=True','source20_true_history=True')
code=code.replace("root/'FixedMetricKernelProof.json'","root/'RefinedCorrectorProof.json'")
old_print="';ALL RETURNED/HALF FORCES,STEEL,METRIC,H CHECKED;NO PHYSICAL LIMIT'"
assert code.count(old_print)==1
code=code.replace(old_print,"(';NEW RETURNED/HALF ROOTS CHECKED;NO PHYSICAL LIMIT' if accepted else ';NO NEW OR HALF ROOTS;SOURCE20 FORCES/STEEL CHECKED;NO PHYSICAL LIMIT')")
sys.argv=[str(source)]
namespace={'__file__':str(source),'__name__':'verified20_arc_corrector_test'}
exec(compile(code,str(source),'exec'),namespace)
proof=json.loads((root/'RefinedCorrectorProof.json').read_bytes())
proof.update(unchanged_control=a.control,only_primary_corrector_branch_scale_changed=not a.control,
             source20_identical_between_variants=True,comparison_root=root.name)
(root/'RefinedCorrectorProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
