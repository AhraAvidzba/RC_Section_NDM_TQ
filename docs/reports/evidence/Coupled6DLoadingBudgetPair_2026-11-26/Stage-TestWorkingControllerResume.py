"""ДЛЯ ТЕСТОВ. Подключает существующий рабочий Trace к холодно восстановленной истории 147 состояний."""
from argparse import ArgumentParser
from hashlib import sha256
from pathlib import Path
from runpy import run_path
import json,re


def TestStage(root,parent):
    """Меняет один тестовый метод; рабочие классы и численные настройки штатного контроллера не изменяет."""
    here=Path(__file__).resolve().parent
    run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
    prep=json.loads((root/'Preparation.json').read_bytes())
    for before,after in [('docs/regression/CodeUpdate.json','ParentCodeUpdate.json'),('CompileClosure.json','ParentCompileClosure.json')]:
        (root/after).write_bytes((root/before).read_bytes());prep['inputs'][after]=sha256((root/after).read_bytes()).hexdigest().upper()
    ancestor=root/('ResumeBuildAncestor_'+root.name);ancestor.mkdir()
    for name in ('CompilePoint.log','CompilePointClosure.json'):
        if (root/name).exists():(root/name).replace(ancestor/name)
    name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
    body=(here/'TestWorkingControllerResume.bas.txt').read_text('utf-8-sig')
    code,count=re.subn(r'(?ms)^Private Sub TestIntegratedYieldEvent\(\).*?^End Sub',lambda _:body,code)
    assert count==1
    (root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'));prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    for name in (Path(__file__).name,'TestWorkingControllerResume.bas.txt','Compile-TestExternalPoint.ps1'):
        (root/name).write_bytes((here/name).read_text('utf-8-sig').encode('utf-8-sig') if name.endswith('.ps1') else (here/name).read_bytes())
        prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update(case='WorkingControllerColdResume147',actual_working_Trace_API=True,
        prescribed_future_field=False,prescribed_material_point_or_level=False,source_classes_unchanged=True,
        numerical_reference_is_SciPy=False,original_force_criteria=True,actual_full_stock_claimed=False,
        physical_capacity_certified=False,actual_import_this_pass=['modTestCoupledElasticState'])
    (root/'Preparation.json').write_bytes((json.dumps(prep,indent=2,sort_keys=True)+'\n').encode())
    print('WORKING TRACE COLD RESUME READY; NO PRESCRIBED EVENT OR FUTURE FIELD')


if __name__=='__main__':
    p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);p.add_argument('--parent',type=Path,required=True)
    a=p.parse_args();TestStage(a.root.resolve(),a.parent.resolve())
