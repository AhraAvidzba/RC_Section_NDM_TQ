"""ДЛЯ ТЕСТОВ. Читает фактический бетонный отклик вдоль последней отклонённой поправки934 для графика пользователя."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold936RecordedYieldCoordinate_20261126';matrices=base/'Cold935Point11Matrices_20261126'
root=base/'Cold937MaterialDiagramView_20261126'
prep935=json.loads((matrices/'Preparation.json').read_bytes());center=sum(prep935['directed_scope']['point11_minus_surface_crossing_fraction_bracket'])/2
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((matrices/'Input/ActualRejectedNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
np.testing.assert_array_equal(accepted,source['fields'][-1]);factor=source['lambdas'][-1]
fractions=sorted(set(np.linspace(0.,1/65536,17).tolist()+np.linspace(center-1e-6,center+1e-6,61).tolist()+[center]))
predictions=[accepted]+[current+f*delta[:60] for f in fractions]
cases=[dict(case=1,name='accepted_zero',lambda_increment=0)]+[
    dict(case=i+2,name='actual934_Newton_fraction_'+str(f),newton_fraction=f,
        lambda_increment=float(meta[2]+f*delta[60]-factor)) for i,f in enumerate(fractions)]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());name='tools/build_workbook/Run-Audit02ModeTests.ps1'
runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected material diagram setup macro.'}
        $null=$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadUnresolvedArcScan")
        $lines.Add('READONLY_PRIVATE_SCAN_COMPLETED: no new equilibrium search')
        Save-Progress
''',1)
values={name:runner.replace('\n','\r\n').encode('utf-8-sig'),
 'Input/ArcScanFields.bin':vector([len(cases)])+b''.join(vector(f) for f in predictions),
 'Input/ArcScanCases.json':(json.dumps(cases,sort_keys=True,indent=2)+'\n').encode(),
 'Input/ActualRejectedNewton.bin':(matrices/'Input/ActualRejectedNewton.bin').read_bytes(),
 'SourcePoint11MatricesProof.json':(matrices/'SurfaceMatricesProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='UserActualConcreteAndSteelDiagramAt934Failure',actual_import_this_pass=['modTestCoupledElasticState'],
 actual_imports_this_replay=[],actual_component_count=257,immutable_compiled_book_reused=True,compiled_source_root=parent.name,
 original25_material_history_unchanged=True,no_new_NDS_or_material_acceptance=True,
 point=11,station_z_over_L=-.001,source_current934_record3=True,
 point11_minus_crossing_fraction=center,endpoint_count=len(cases),physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ACTUAL MATERIAL GRAPH',len(cases),'ENDPOINTS;NO NEW NDS')
