"""ДЛЯ ТЕСТОВ. Разрешает меньшие доли штатного поиска на точной системе54/55, сохраняя исходную приёмку и общий бюджет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold925RecordedTailSystem_20261126';root=base/'Cold929RecordedTailFineLine_20261126'
continuity=json.loads((base/'Cold928RecordedTailBoundaryScan_20261126/RecordedTailBoundaryScanProof.json').read_bytes())
pairs=continuity['boundary_pairs'];assert pairs[-1]['maximum_point5_center_stress_difference_MPa']<2e-11
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'FineLineCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
anchor='        If mSteelBranchBuilds > 0 Then minimumFraction = STEEL_MIN_FRACTION'
assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''
        If mControlActive And mUseSpatialHistory Then minimumFraction = 0.0000152587890625 ' Допускает чтение по входящую сторону близкой материальной поверхности; силовая приёмка неизменна.
'''.rstrip(),1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledNewtonSolver'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
anchor='        $lines.Add([string]$excel.Run("\'RC_Section_NDM.xlsm\'!modTestCoupledElasticState.TestReadRecordedTailSystem"))'
assert runner.count(anchor)==1
runner=runner.replace(anchor,'        $null=$excel.Run("\'RC_Section_NDM.xlsm\'!modTestColdNewtonCapture.TestSetCaptureRoot",$Root)\n'+anchor,1)
values[name]=runner.replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExactTailFineArmijoLine',actual_import_this_pass=['CCoupledNewtonSolver'],
 actual_component_count=257,minimum_fraction=1/65536,original_force_control_and30_iterations=True,
 numerical_change_only_finer_line_candidates=True,extra_material_history_acceptance=False,
 actual_rejected_directions_recorded=True,all_materials_and_all_other_numerics_exact925=True,
 no_working_or_release_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT54/55; FINER ARMIJO LINE ONLY; SAME30/ACCEPTANCE')
