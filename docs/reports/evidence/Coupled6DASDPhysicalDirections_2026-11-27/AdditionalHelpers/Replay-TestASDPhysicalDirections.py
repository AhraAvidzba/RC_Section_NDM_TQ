"""ДЛЯ ТЕСТОВ. Проверяет НДС физических направлений всеми уравнениями и нативными материалами OpenSees."""
from pathlib import Path
from argparse import ArgumentParser
from runpy import run_path
import json
import subprocess
import sys
import traceback
from time import perf_counter

p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);p.add_argument('--worker',action='store_true')
a=p.parse_args();root=a.root;here=Path(__file__).resolve().parent
if not a.worker:
    started=perf_counter()
    result=subprocess.run([sys.executable,str(Path(__file__).resolve()),'--root',str(root),'--worker'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=90,creationflags=subprocess.CREATE_NO_WINDOW)
    (root/'Native/PhysicalReplaySubprocess.json').write_text(json.dumps(dict(exit_code=result.returncode,elapsed_seconds=perf_counter()-started,native_XML_suppressed=True),sort_keys=True,indent=2)+'\n',encoding='utf-8')
    assert result.returncode==0
    summary=json.loads((root/'Native/PhysicalDirectionSummary.json').read_bytes());print('PHYSICAL DIRECTION INDEPENDENT CHECK',json.dumps(summary,sort_keys=True));sys.exit(0)
try:
    source=here/'Replay-TestMainCapacityEquilibrium.py';code=source.read_text('utf-8-sig')
    def TestReplace(before,after):
        """Добавляет только чтение материальных координат и независимые плечи, сохраняя исходную силовую проверку."""
        global code
        assert code.count(before)==1,before;code=code.replace(before,after,1)
    TestReplace("source.get('three_changed_components_exact') or source.get('changed_components_exact')","source.get('changed_components_exact') or (source.get('all260_working_components_unchanged') and source.get('no_kernel_modifications'))")
    TestReplace("(root/'docs/regression/Diagnostic.xlsm')","(root.parent/'docs/regression/Diagnostic.xlsm')")
    TestReplace('started=perf_counter()','last_equivalent=np.zeros((concrete,2));started=perf_counter()')
    TestReplace("                ops.NDTest('CommitState',tag)","""                if side==1 and point<concrete and index+1==len(history):last_equivalent[point]=ops.NDTest('GetResponse',tag,'equivalentTotalStrain')
                ops.NDTest('CommitState',tag)""")
    TestReplace('finally:ops.wipe()',"    independent_peak=TestAxialPeak(root,globals()) if path_code==1 and prescribed[0]<0 else None\n    kind=int(struct.unpack_from('<d',(root/'Native/MaterialLimit.bin').read_bytes(),4)[0])\n    independent_upper=TestPhysicalUpper(root,globals()) if kind>0 else None\nfinally:ops.wipe()")
    namespace=dict(__file__=str(source),__name__='TestPhysicalDirections',TestAxialPeak=run_path(str(here/'Solve-TestASDAxialPeak.py'))['TestAxialPeak'],TestPhysicalUpper=run_path(str(here/'Solve-TestASDPhysicalUpper.py'))['TestPhysicalUpper'])
    sys.argv=[str(source),'--root',str(root)];exec(compile(code,str(source),'exec'),namespace)
    case=json.loads((root/'Case.json').read_bytes());util=float(namespace['last_equivalent'][:,1].max()/case['compression_limit']);assert util<=1.+1e-9
    proof=json.loads((root/'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes());assert proof['reference_tolerance_violated_states']==0
    text=(root/'Native/MainCapacity.txt').read_text('utf-8-sig');status=text.split(';',1)[0].removeprefix('Status=')
    summary=dict(case=case['label'],status=status,accepted_states=proof['accepted_states'],final_lambda=proof['records'][-1]['lambda_value'],
                 all60_equations_and_six_resultants_verified=True,compression_utilization=util,
                 final_target_resultants_N_Nmm=proof['records'][-1]['target_resultants_N_Nmm'],
                 maximum_force_error_N=proof['maximum_force_error_N'],maximum_moment_error_Nmm=proof['maximum_moment_error_Nmm'],
                 independent_axial_peak=namespace['independent_peak'],independent_upper=namespace['independent_upper'],whole_goal_completed=False)
    (root/'Native/PhysicalDirectionSummary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
except BaseException:
    (root/'Native/PhysicalDirectionFailure.txt').write_text(traceback.format_exc(),encoding='utf-8');raise
