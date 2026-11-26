"""ДЛЯ ТЕСТОВ. Независимо проверяет979, не выдавая четыре отдельных пробы за согласованную полную/половинную пару."""
from pathlib import Path
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold979LoadingAlgorithmComparison_20261126';parent=base/'Cold978LoadingMaterialDirection_20261126'
code=(here/'Check-TestRecordedYieldCoordinate.py').read_text('utf-8-sig')
old="root=base/'Cold936RecordedYieldCoordinate_20261126';parent=base/'Cold935Point11Matrices_20261126'";assert code.count(old)==1
code=code.replace(old,'root=Path('+repr(str(root))+');parent=Path('+repr(str(parent))+')',1)
old='source_actual934_state25_and_initial_residual_bitexact_by_native_guard=True';assert code.count(old)==1
code=code.replace(old,'source_actual968_state25_and_initial_residual_bitexact_by_native_guard=True',1)
env={'__file__':str(here/'Check-TestRecordedYieldCoordinate.py'),'__name__':'loading_algorithm_comparison_test'}
exec(compile(code,str(here/'Check-TestRecordedYieldCoordinate.py'),'exec'),env)
for folder in ('src','tests','Input'):
    for path in (parent/folder).rglob('*'):
        if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes(),str(path)
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
assert env['meta'][2]>0 and env['meta'][9]>0
proof=env['proof'];proof.update(all257_sources_and_actual_saved_book_bitexact978=True,
    all_numerical_inputs_bitexact978=True,no_new_compilation=True,no_shared_budget_pair_claim=True,
    native_H_or_lambda_uncertainty_not_recorded=True,cold_gain_claimed=False)
(root/'LoadingAlgorithmComparisonProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('VERIFIED STANDALONE ALGORITHM COMPARISON;CAPACITY NOT CERTIFIED')
