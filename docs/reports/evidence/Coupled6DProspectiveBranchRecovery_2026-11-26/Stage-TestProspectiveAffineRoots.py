"""ДЛЯ ТЕСТОВ. Проверяет прогнозы Newton от самой точки новой ветви, включая её фактическую невязку, а не только её матрицу."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
matrices=base/'Cold914ProspectiveBranchMatrix_20261126';parent=base/'Cold900Cold32OutgoingJacobian_20261126'
root=base/'Cold916ProspectiveAffineRoots_20261126'
assert json.loads((matrices/'ProspectiveBranchMatrixProof.json').read_bytes())['base_residual_maximum_difference_from_actual_capture_MPa']==0
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((matrices/'Input/FrozenNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,old_delta=[r.VectorTest() for _ in range(6)]
original=np.array([r.VectorTest() for _ in range(61)]);assert r.position==len(r.raw)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
np.testing.assert_array_equal(accepted,source['fields'][-1]);factor=source['lambdas'][-1]
predictions=[accepted,current];cases=[dict(case=1,name='accepted_zero',lambda_increment=0),
    dict(case=2,name='actual_failed_current',lambda_increment=float(meta[2]-factor))]
r=R((matrices/'Native/ProspectiveBranchMatrix.bin').read_bytes())
for level in (0,1/256,1/128,1/64,1/32,.25,1):
    info,seed,scales=[r.VectorTest() for _ in range(3)];assert info[0]==level and info[2] and info[3]
    values=r.VectorTest();A=np.array([r.VectorTest() for _ in range(61)]);r.VectorTest()
    for _ in range(8):
        h=r.VectorTest();r.VectorTest()
        if h[1]:r.VectorTest()
    if level<1/128:continue
    native_residual=values.copy();native_residual[56:60]/=meta[4]
    native_residual[-1]=-rhs[-1]+original[-1,:60]@(seed-current)+original[-1,-1]*(info[1]-meta[2])
    column=np.max(abs(A),axis=0);row=np.max(abs(A/column),axis=1)
    correction=np.linalg.solve(A/column/row[:,None],-native_residual/row)/column
    rootfield=seed+correction[:60];rootfactor=info[1]+correction[-1]
    linear_error=float(max(abs(A@correction+native_residual)))
    assert linear_error<1e-15
    predictions.append(rootfield);cases.append(dict(case=len(cases)+1,name='affine_Newton_from_postyield_'+str(level),
        prospective_level=level,lambda_increment=float(rootfactor-factor),actual_lambda=float(rootfactor),
        virtual_linear_residual_MPa=linear_error,seed_residual_MPa=float(max(abs(native_residual[:56])))))
assert r.position==len(r.raw) and len(cases)==7
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());name='tools/build_workbook/Run-Audit02ModeTests.ps1'
runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected affine prospective setup macro.'}
        $null=$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadUnresolvedArcScan")
        $lines.Add('READONLY_PRIVATE_SCAN_COMPLETED: no new equilibrium search')
        Save-Progress
''',1)
values={name:runner.replace('\n','\r\n').encode('utf-8-sig'),
 'Input/ArcScanFields.bin':vector([len(cases)])+b''.join(vector(f) for f in predictions),
 'Input/ArcScanCases.json':(json.dumps(cases,sort_keys=True,indent=2)+'\n').encode(),
 'Input/FrozenNewton.bin':(matrices/'Input/FrozenNewton.bin').read_bytes(),
 'SourceProspectiveBranchMatrixProof.json':(matrices/'ProspectiveBranchMatrixProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ActualNewtonProspectiveAffineRootReadOnly',actual_import_this_pass=json.loads((parent/'docs/regression/CodeUpdate.json').read_bytes())['Modules'],
 actual_imports_this_replay=[],actual_component_count=256,immutable_compiled_book_reused=True,compiled_source_root=parent.name,
 exact_original_scalar_control_row_and_load_column=True,affine_residual_from_postyield_seed=True,
 no_material_history_acceptance=True,no_native_NDS_search=True,physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'FIVE AFFINE BRANCH ROOT PREDICTIONS;FULL MATERIAL/SIX FORCE CHECK REQUIRED')
