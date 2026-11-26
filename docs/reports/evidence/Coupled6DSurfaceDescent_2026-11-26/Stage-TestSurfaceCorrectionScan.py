"""ДЛЯ ТЕСТОВ. Проверяет фактический материал вдоль независимых поправок обеих касательных, включая малые доли."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
matrices=base/'Cold931SurfaceMatrices_20261126';parent=matrices;root=base/'Cold932SurfaceCorrectionScan_20261126'
scope=run_path(str(here/'Check-TestSurfaceMatrices.py'))
source=scope['source'];current=scope['current'];accepted=scope['accepted'];factor=source['lambdas'][-1]
predictions=[accepted,current];cases=[dict(case=1,name='accepted_zero',lambda_increment=0),
 dict(case=2,name='actual_929_failed_current',lambda_increment=float(scope['meta'][2]-factor))]
selected=[x for x in scope['predictions'] if (x['matrix_index']==0 and x['mode']=='matched_seed') or x['matrix_index'] in (4,5)]
assert len(selected)==5
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](matrices/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+t*.001*gradient for t in (-1,0,1)])
for prediction in selected:
    for exponent in range(17):
        f=2.**(-exponent);field=prediction['field']+f*prediction['correction'][:60]
        lam=prediction['lambda_value']+f*prediction['correction'][-1]
        bars=200000*np.einsum('spc,pc->sp',np.einsum('spcd,d->spc',op,field)[:,25:],axis[25:])
        long=axis[25:,2]==1;fy=np.where(long[None,:],np.where(bars>=0,400,350),240)
        assert np.all(abs(bars)<fy), 'Нужна сборка нелинейной стали, а не упругое допущение.'
        predictions.append(field);cases.append(dict(case=len(cases)+1,name='surface_correction_'+str(prediction['matrix_index'])+'_'+prediction['mode']+'_'+str(f),
            matrix_index=prediction['matrix_index'],matrix_level=prediction['level'],mode=prediction['mode'],correction_fraction=f,
            lambda_increment=float(lam-factor),linear_backward_error=prediction['linear_backward_error']))
assert len(cases)==87
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());name='tools/build_workbook/Run-Audit02ModeTests.ps1'
runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected surface correction setup macro.'}
        $null=$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadUnresolvedArcScan")
        $lines.Add('READONLY_PRIVATE_SCAN_COMPLETED: no new equilibrium search')
        Save-Progress
''',1)
values={name:runner.replace('\n','\r\n').encode('utf-8-sig'),
 'Input/ArcScanFields.bin':vector([len(predictions)])+b''.join(vector(x) for x in predictions),
 'Input/ArcScanCases.json':(json.dumps(cases,sort_keys=True,indent=2)+'\n').encode(),
 'SourceSurfaceMatricesProof.json':(matrices/'SurfaceMatricesProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='IncomingOutgoingCorrectionMaterialReadOnly',actual_import_this_pass=['modTestCoupledElasticState'],
 actual_imports_this_replay=[],actual_component_count=257,immutable_compiled_book_reused=True,compiled_source_root=parent.name,
 accepted_history25_unchanged=True,all_candidate_steel_strains_confirmed_elastic=True,
 original_scalar_row_and_load_column_used=True,no_new_native_NDS_or_history_acceptance=True,
 physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'85 REAL MATERIAL ENDPOINTS ALONG FIVE NEWTON DIRECTIONS;NO NDS ACCEPTANCE')
