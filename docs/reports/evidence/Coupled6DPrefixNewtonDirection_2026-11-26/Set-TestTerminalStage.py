"""ДЛЯ ТЕСТОВ. Отмечает только фактически завершённые сессии; цель остаётся активной между этапами."""
from pathlib import Path
from datetime import datetime
import argparse,json

p=argparse.ArgumentParser(description=__doc__)
p.add_argument('--sessions',type=int,nargs='+',required=True);p.add_argument('--next-action',required=True)
a=p.parse_args();repo=Path.cwd();path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json'
state=json.loads(path.read_bytes())
state['terminal_sessions_this_stage']=sorted(set(state['terminal_sessions_this_stage']+a.sessions))
state['active_jobs']=[j for j in state.get('active_jobs',[]) if j['session'] not in a.sessions]
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
    active_exec_session=state['active_jobs'][0]['session'] if state['active_jobs'] else None,
    all_started_stages_terminal=not state['active_jobs'],next_action=a.next_action)
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):target.write_bytes(raw)
print('TERMINAL',a.sessions,'REMAINING JOBS',len(state['active_jobs']),'GOAL ACTIVE')
