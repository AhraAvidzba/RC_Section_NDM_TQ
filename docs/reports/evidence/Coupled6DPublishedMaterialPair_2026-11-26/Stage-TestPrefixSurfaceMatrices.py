"""ДЛЯ ТЕСТОВ. Читает существующим входом касательные точки21/+1 на фактической отклонённой системе968 без нового кода VBA."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold969PrefixSpatialFloor_20261126';origin=base/'Cold968PrefixNewtonCapture_20261126'
root=base/'Cold970PrefixSurfaceMatrices_20261126'
proof=json.loads((origin/'PrefixNewtonCaptureProof.json').read_bytes());record=proof['rejected_newton_records'][1]
assert record['metadata'][7:]==[1.,1.]
raw=(origin/'Native'/record['file']).read_bytes();assert sha256(raw).hexdigest().upper()==record['sha256']
boundary=record['loading_switch_below_smallest_recorded_fraction'][0];assert boundary['point']==21 and boundary['side']==1
center=sum(boundary['bracket'])/2
levels=[0.,center-1e-5,center+1e-5,center+1e-6,center+1e-7,center+1e-8,1/65536,1/1024,1/128]
positions=[0]
for _ in range(67):
    count,=struct.unpack_from('<i',raw,positions[-1]);positions.append(positions[-1]+4+8*count)
fixture=raw[:positions[-1]]+raw[positions[6]:positions[-1]]
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector'];updates={'Input/FrozenNewton.bin':fixture,
    'Input/SurfaceMatrixLevels.bin':vector(levels),'Input/ActualRejectedNewton.bin':raw,
    'SourcePrefixNewtonCaptureProof.json':(origin/'PrefixNewtonCaptureProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
name='tools/build_workbook/Run-Audit02ModeTests.ps1';code=(root/name).read_text('utf-8-sig')
old='            $null = $excel.Run("\'RC_Section_NDM.xlsm\'!modTestColdNewtonCapture.TestSetCaptureRoot", $MacroArgument1)\n'
assert code.count(old)==1;code=code.replace(old,'',1)
old='modTestCoupledElasticState.TestRunControlledSecantPrefix';assert code.count(old)==1
updates[name]=code.replace(old,'modTestCoupledElasticState.TestReadSurfaceMatrices',1).replace('\n','\r\n').encode('utf-8-sig')
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ActualPrefixPoint21PlusIncomingOutgoingMatrices',actual_import_this_pass=[],actual_component_count=257,
    all257_actual_components_and_saved_book_bitexact969=True,no_new_numerical_code_or_compilation=True,
    actual_rejected_record968=2,source_refined_matrix=True,original25_history_and_loadpath_preserved=True,
    first_matrix_and_residual_must_match_actual_record_before_adjacent_reading=True,readonly_no_NDS_search_or_acceptance=True,
    point21_plus_loading_surface_fraction_bracket=boundary['bracket'],matrix_levels=levels,
    no_future_equilibrium_input=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT969 BOOK;EXISTING SURFACE MATRIX ENTRY;ACTUAL968 REFINED SYSTEM;NO NEW VBA OR NDS SEARCH')
