"""ДЛЯ ТЕСТОВ. Готовит точный повтор последнего полного снимка718, не повторяя поиск предела."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json,struct
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold720PostCrackRawReplay_20261124';parent=base/'Cold718ControlledLoadBalance_20261124'
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](parent/'Native/Stock.RawHistory.bin')
frame=frames[-1];assert len(frame['fields'])>=63 and raw['incomplete_tail_bytes']==0
checkpoint=run_path(str(here/'Read-TestContinuationCheckpoint.py'))['TestRead'](parent/'Native/Stock.Checkpoints.csv')
assert checkpoint['sequence']==frame['number'] and checkpoint['count']==len(frame['fields'])
np.testing.assert_allclose(np.asarray([p['field'] for p in checkpoint['points']]),frame['fields'],rtol=1e-14,atol=1e-300)
template=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Native/PostCrackPrescribed.Accepted2.bin')
for key in ('fields','fractions','lambdas'):assert np.array_equal(frame[key][:2],template[key]),key
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'RawReplayCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
data=bytearray(b'NDMPOSTREPLAY001'+struct.pack('<i',60))
def TestVector(values):
    """Копирует исходные IEEE-754 Double; размеры записаны как VBA Long."""
    values=np.asarray(values,dtype='<f8');return struct.pack('<i',len(values))+values.tobytes()
for key in ('calibration','options','offset','direction'):data+=TestVector(template[key])
data+=struct.pack('<i',len(frame['fields']))
for fraction,factor,field in zip(frame['fractions'],frame['lambdas'],frame['fields']):data+=struct.pack('<2d',fraction,factor)+TestVector(field)
values['Input/PostCrackReplay.bin']=bytes(data)
values['Input/PostCrackReplayOptions.bin']=struct.pack('<i3d',3,0,0,0)
values['Input/PostCrackSourceRawHistory.bin']=(parent/'Native/Stock.RawHistory.bin').read_bytes()
path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
anchor='Optional ByVal tangentModeTest As Long = 0) As String';assert code.count(anchor)==1
code=code.replace(anchor,'Optional ByVal tangentModeTest As Long = 0, Optional ByVal replayOnlyTest As Boolean = False) As String',1)
anchor='    Dim handle As Integer, hasFile As Boolean\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''    If replayOnlyTest Then
        RunTestCoupledPostCrackTension = TestReplayPostCrackRaw(solver, operators, concrete, response, _
            Array(EC, 0.2, 30#, FT, 16#, 0.002, GT, 80#, LENGTH), Array(4#, 64#, 30#, ABS_TOL, REL_TOL, 0.001, 3#), offset, direction, projection, evidenceRoot)
        Exit Function
    End If
''',1)
code+='\n'+(here/'TestPostCrackRawReplay.bas.txt').read_text('utf-8-sig')
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=(here/'Compile-TestExternalPoint.ps1').read_bytes()
runner=(parent/'Run-TestRealPostCrack.ps1').read_text('utf-8-sig').replace('RunTestCoupledPostCrackCapacity','RunTestPostCrackRawReplay').replace('PostCrackTension.txt','PostCrackReplay.txt')
values['Run-TestPostCrackRawReplay.ps1']=runner.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
for name in (Path(__file__).name,'TestPostCrackRawReplay.bas.txt'):values[name]=(here/name).read_bytes()
for name,content in values.items():(root/name).write_bytes(content);prep['inputs'][name]=sha256(content).hexdigest().upper()
prep['directed_scope'].update(case='ExactPostCrack718RawReplay',actual_import_this_pass=['modTestCoupledElasticState'],
    source_raw_history=raw,raw_history_source=str(parent),source_states=len(frame['fields']),
    last_source_strain=float(frame['fields'][-1,0]),no_new_cold_or_capacity_search=True,
    all_current_material_and_geometry_parameters_exact=True,original_binary_fields_and_loads_preserved=True,
    replay_mode=0,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('720 EXACT RAW REPLAY',len(frame['fields']),'STATES; NO SEARCH; ORIGINAL MATERIALS, GEOMETRY AND BINARY FIELDS')
