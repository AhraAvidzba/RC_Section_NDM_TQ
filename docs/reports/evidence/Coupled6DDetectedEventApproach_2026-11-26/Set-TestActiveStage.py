"""ДЛЯ ТЕСТОВ. Обновляет журнал текущей сессии без изменения статуса цели или расчётных исходников."""
from pathlib import Path
from datetime import datetime
import argparse,json

p=argparse.ArgumentParser(description=__doc__);p.add_argument('--session',type=int,required=True);p.add_argument('--root',required=True);p.add_argument('--terminal',type=int,nargs='*',default=[])
p.add_argument('--stage',default='directed native material diagnosis');p.add_argument('--next-action')
p.add_argument('--secondary-session',type=int);p.add_argument('--secondary-stage',default='directed native regressions')
p.add_argument('--secondary-root')
p.add_argument('--third-session',type=int);p.add_argument('--third-root');p.add_argument('--third-stage',default='directed native regressions')
a=p.parse_args();repo=Path.cwd();path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json'
if a.session <= 0 or (a.secondary_session is not None and a.secondary_session <= 0) or (a.third_session is not None and a.third_session <= 0):
    p.error('Нужен положительный идентификатор фактически возвращённой сессии.')
state=json.loads(path.read_bytes());state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
    pause_requested_explicitly=False,pause_request_revoked=True,active_exec_session=a.session,
    active_jobs=[dict(session=a.session,root=a.root,stage=a.stage)],all_started_stages_terminal=False)
if a.next_action:state['next_action']=a.next_action
if a.secondary_session is not None:
    state['active_jobs'].append(dict(session=a.secondary_session,root=a.secondary_root or a.root,stage=a.secondary_stage))
if a.third_session is not None:
    state['active_jobs'].append(dict(session=a.third_session,root=a.third_root or a.root,stage=a.third_stage))
state['terminal_sessions_this_stage']=sorted(set(state['terminal_sessions_this_stage']+a.terminal))
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):target.write_bytes(raw)
print('ACTIVE',a.session,a.root)
