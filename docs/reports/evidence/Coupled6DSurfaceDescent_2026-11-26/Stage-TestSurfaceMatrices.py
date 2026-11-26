"""ДЛЯ ТЕСТОВ. Готовит касательные обеих сторон нового отказа929 от неизменной истории25 и записанной скалярной строки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import struct,json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold925RecordedTailSystem_20261126';source=base/'Cold929RecordedTailFineLine_20261126'
root=base/'Cold931SurfaceMatrices_20261126'
proof=json.loads((Path.cwd()/'docs/reports/evidence/Coupled6DRecordedNewtonTail_2026-11-26/FineLineSurfaceOriginProof.json').read_bytes())
center=sum(proof['plus_station_crossing_fraction_bracket'])/2
levels=[0.,center-1e-6,center+1e-6,center+1e-7,center+1e-8,1/65536,1/1024]
def TestBounds(raw):
    """Находит векторы фактической записи без потери точности Double."""
    positions=[0]
    for _ in range(67):
        count,=struct.unpack_from('<i',raw,positions[-1]);positions.append(positions[-1]+4+8*count)
    return positions
a=(source/'Native/NewtonRejected.00001.bin').read_bytes();b=(source/'Native/NewtonRejected.00002.bin').read_bytes()
pa,pb=TestBounds(a),TestBounds(b)
for i in (1,2,3,4):assert a[pa[i]:pa[i+1]]==b[pb[i]:pb[i+1]]
fixture=a[:pa[-1]]+b[pb[6]:pb[-1]]
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'SurfaceMatricesCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*9+'\n\nPublic Function TestReadRecordedTailSystem() As String';assert code.count(anchor)==1
entry=(here/'TestReadSurfaceMatrices.bas.txt').read_text('utf-8-sig').rstrip()
values[name]=code.replace(anchor,prefix*10+'\n\n'+entry+'\n\n\nPublic Function TestReadRecordedTailSystem() As String',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestReadSurfaceMatrices'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
assert runner.count('modTestCoupledElasticState.TestReadRecordedTailSystem')==1
values[name]=runner.replace('modTestCoupledElasticState.TestReadRecordedTailSystem','modTestCoupledElasticState.TestReadSurfaceMatrices',1).replace('\n','\r\n').encode('utf-8-sig')
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
values.update({'Input/FrozenNewton.bin':fixture,'Input/SurfaceMatrixLevels.bin':vector(levels),
 'Input/ActualRefinedNewton.bin':b,'Input/ActualRejectedNewton.bin':a,
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Exact929IncomingOutgoingBranchMatrices',actual_import_this_pass=['modTestCoupledElasticState'],
 actual_component_count=257,all_numerical_classes_exact925=True,accepted_history25_unchanged=True,
 exact_recorded929_pair=[1,2],matrix_levels=levels,original_scalar_row_and_load_column_preserved=True,
 no_new_NDS_search_or_material_acceptance=True,physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'SEVEN INCOMING/OUTGOING MATRICES;EXACT929 CURRENT/HISTORY;NO NEW NDS')
