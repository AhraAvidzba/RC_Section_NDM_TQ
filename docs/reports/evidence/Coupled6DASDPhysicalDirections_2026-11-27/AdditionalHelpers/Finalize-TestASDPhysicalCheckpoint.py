"""ДЛЯ ТЕСТОВ. Закрывает завершённые численные опыты My с явным отрицательным результатом и сохраняет продолжение цели."""
from pathlib import Path
from hashlib import sha256
import json
import shutil
import struct

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
evidence=repo/'docs/reports/evidence/Coupled6DASDPhysicalDirections_2026-11-27'
root=base/'Cold1155ASDLinearTrustDifferenceMy_20261127'
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged']
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
closure=json.loads((root/'Case03/Native/RealColdReplay.txt.ProcessClosure.json').read_bytes())
assert closure['ChildHasExited'] and closure['ChildExitCode']==0 and closure['Passed']
assert json.loads((root/'CompilePointClosure.json').read_text('utf-8-sig'))['State']=='Absent'
summary=json.loads((root/'Case03/Native/PhysicalDirectionSummary.json').read_bytes())
assert summary['status']=='NumFail' and summary['all60_equations_and_six_resultants_verified'] and summary['accepted_states']==72
assert json.loads((root/'Case03/Native/PhysicalReplaySubprocess.json').read_bytes())['exit_code']==0
values=struct.unpack_from('<10d',(root/'Case03/Native/MainCapacity.bin').read_bytes(),20)
assert not values[1] and values[0]==379.8046875
files=json.loads((evidence/'AdditionalFiles.json').read_bytes())
sources=[root/relative for relative in ['src/Solver/CCoupledNewtonSolver.cls','Preparation.json','EarlySourceProof.json',
           'CompilePoint.log','CompilePointClosure.json','docs/regression/CodeUpdate.json']]
sources += [p for p in (root/'Case03/Native').iterdir() if p.is_file() and p.suffix!='.csv']
for source in sources:
    target=evidence/'Trial1155'/source.relative_to(root);target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source,target)
    files[target.relative_to(evidence).as_posix()]=sha256(source.read_bytes()).hexdigest().upper()
summary.update(native_call_seconds=values[0],physical_ultimate_confirmed=False,not_adopted_in_production=True,
               maximum_admissible_lambda=6.575169710628729,no_increase_in_maximum_lambda=True,
               exact_private_book_retained_on_D=str(root/'docs/regression/Diagnostic.xlsm'),
               exact_private_book_sha256=compiled['WorkbookSHA256'],book_not_in_compact_trial_evidence=True,
               own_compiler_and_test_Excel_closed=True)
(evidence/'Trial1155/Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for source in (base/'Cold1153ASDPhysicalDirections_20261127/Case03/Native').glob('IndependentMy*json'):
    if source.name in ['IndependentMyProbes.json','IndependentMyProbes-last-hybr.json','IndependentMyProbes-last-trf.json']:
        continue
    target=evidence/'AdditionalMy'/source.name;target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source,target)
    files[target.relative_to(evidence).as_posix()]=sha256(source.read_bytes()).hexdigest().upper()
for name in ['Finalize-TestASDPhysicalCheckpoint.py','Replay-TestASDPhysicalDirections.py']:
    source=here/name;target=evidence/'AdditionalHelpers'/name;shutil.copyfile(source,target)
    files[target.relative_to(evidence).as_posix()]=sha256(source.read_bytes()).hexdigest().upper()
(evidence/'AdditionalFiles.json').write_text(json.dumps(files,sort_keys=True,indent=2)+'\n',encoding='utf-8')
external=json.loads((evidence/'AdditionalMy/IndependentMySteelEndpointRefined.json').read_bytes())
assert external['results'][0]['equilibrium_verified'] and not external['results'][1]['equilibrium_verified']
assert external['results'][0]['lambda_value']==6.703514362435932
book=repo/'workbook/output/RC_Section_NDM.xlsm'
assert sha256(book.read_bytes()).hexdigest().upper()=='1E6443C2CFDD0058FE5EA559D1D45A8E94CC5199A3ADE69149A06CD4265D8559'
state_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(state_path.read_bytes())
state.update(asd_linear_trust_difference_negative_20261127=summary,asd_My_external_endpoints_20261127=external,
             current_load_stage='N compression, negative Mx and pure T confirmed; My remains open; both linear-QR/step candidates negative and not adopted; endpoint epsilonSteel=0.010 externally verified at My=6.703514 kNm',
             active_jobs=[],active_cold_exec_session=None,all_started_stages_terminal=True,
             goal_status='active',entire_goal_completed=False,last_goal_turn_classification='progress',
             next_actions=['Keep further formatting deferred',
                           'Do not rerun cold candidates 1154/1155 unchanged; maximum lambda did not improve',
                           'Inspect native single-update ASD consistent material tangent and accepted-history re-equilibration before another long cold My search',
                           'Verify any changed tangent independently on material tensors before section use',
                           'Keep physical/tolerance acceptance unchanged; unverified epsilonSteel=0.025 endpoint is not a limit',
                           'After resolving separate N/M/T, continue NMT then Q then NMTQ; ASD before CDP'])
state_path.write_text(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8')
report_path=repo/'docs/reports/Coupled6DASDPhysicalDirections_2026-11-27.md'
report=report_path.read_text('utf-8')
report+='''

Продолжение численной диагностики: кандидат 1154 с общим линейным QR-резервом оставил максимум My неизменным и завершился NumFail за 395,64 с; его конечная история побайтно совпала с проверенным исходным результатом. Кандидат 1155 с шагом центральной разности 0,001 в численных координатах также сохранил максимум 6,575170 кН·м и NumFail, за 379,80 с. Все его 72 принятых состояния независимо проверены. Увеличение числа состояний не объявлено улучшением физического результата. Оба варианта не перенесены в рабочее ядро. Точные частные книги сохранены на D с закреплёнными SHA; компактные исходники и отчёты находятся в Trial1154/Trial1155 внутри доказательств.

При последовательном внешнем продолжении от отдельно проверенного правого корня принято ещё одно равновесие при My=6,626179 кН·м; устойчивое прохождение всей ветви не получено. От той же исходной принятой истории с точным исключением линейной связи средней деформации продольной стали найдена допустимая точка при εs=0,010 и My=6,703514 кН·м: физическая невязка 7,36·10⁻⁷ МПа при допустимой 2,09·10⁻⁶ МПа, использование бетонного ограничения 0,3176. Это отдельный внешний корень, а не холодный результат VBA или доказанная непрерывная ветвь. При εs=0,025 проба около 7,06978 кН·м сохранила невязку 1,07·10⁻⁴ МПа после нескольких численных уточнений и отвергнута; собственный успешный статус SciPy не заменяет исходного силового допуска.

Все запущенные опыты этого этапа завершены, их собственные Excel закрыты. Цель активна. Следующая доработка должна опираться на согласованную производную фактически используемого однократного ASD-обновления и проверку входящего равновесия принятой истории; повторение неизменных холодных вариантов 1154/1155 не запланировано. Оформление, NMT и остальные незавершённые этапы не объявлены выполненными.
'''
report_path.write_text(report,encoding='utf-8')
print('BOTH MY CANDIDATES NEGATIVE AND NOT ADOPTED;ALL STARTED JOBS TERMINAL;PUBLIC BOOK UNCHANGED;GOAL ACTIVE')
