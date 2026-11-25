"""ДЛЯ ТЕСТОВ. Уточняет фактический порядок трёх импортов749; состав, исходники, книга и содержательные проверки не меняются."""
from pathlib import Path
from hashlib import sha256
import json

root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold749DepartureCandidateAudit_20261125')
prep=json.loads((root/'Preparation.json').read_bytes())
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
requested=prep['directed_scope']['actual_import_this_pass'];actual=compiled['Modules']
assert requested==['CCoupledNewtonSolver','CCoupledDeformationContinuation','modTestCoupledElasticState']
assert actual==['CCoupledDeformationContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']
assert len(actual)==len(requested)==len(set(actual)) and set(actual)==set(requested)
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
assert json.loads((root/'CompilePointClosure.json').read_bytes())['State']=='Absent'
proof=dict(initial_source_audit_failed_on_order_only=True,requested_order=requested,actual_source_traversal_order=actual,
    exact_three_names_preserved=True,no_code_or_book_edit=True,full255_component_audit_still_required=True)
(root/'ImportOrderAuditCorrection.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
prep['directed_scope']['requested_import_order']=requested
prep['directed_scope']['actual_import_this_pass']=actual
prep['directed_scope']['import_order_follows_actual_source_traversal']=True
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('749 DECLARED ACTUAL ORDER CORRECTED; EXACT3 MODULES; FULL CONTENT AUDIT NOT WAIVED')
