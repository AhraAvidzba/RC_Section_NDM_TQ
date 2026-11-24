"""ДЛЯ ТЕСТОВ. Изолирует один прямой шаг от точного принятого префикса, без нового поиска предела."""
from argparse import ArgumentParser
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json, struct
import numpy as np


def TestStage(root, parent, count, target, maximum):
    """Меняет только двоичные данные теста; собранный VBA переиспользуется без изменения материала."""
    here=Path(__file__).resolve().parent
    frames, raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](parent/'Input/PostCrackSourceRawHistory.bin')
    frame=frames[-1]
    assert 2 <= count <= len(frame['fields']) and 1 <= maximum <= 30
    assert target > frame['fields'][count-1,0]
    run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
    source=(parent/'Input/PostCrackReplay.bin').read_bytes()
    Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']; reader=Reader(source)
    size,=reader.ReadTest('<i'); assert size==60
    for _ in range(4): reader.VectorTest()
    prefix=source[:reader.position]; original_count,=reader.ReadTest('<i')
    assert original_count==len(frame['fields'])
    data=bytearray(prefix+struct.pack('<i',count))
    for i in range(count):
        data+=struct.pack('<ddi',frame['fractions'][i],frame['lambdas'][i],size)
        data+=np.asarray(frame['fields'][i],dtype='<f8').tobytes()
    values={'Input/PostCrackReplay.bin':bytes(data),
        'Input/PostCrackReplayOptions.bin':struct.pack('<i3d',3,1,target,maximum)}
    runner=(parent/'Run-TestPostCrackRawReplay.ps1').read_text('utf-8-sig').replace('-TimeoutSeconds 600','-TimeoutSeconds 180')
    values['Run-TestPostCrackRawReplay.ps1']=runner.replace('\n','\r\n').encode('utf-8-sig')
    values[Path(__file__).name]=Path(__file__).read_bytes()
    prep=json.loads((root/'Preparation.json').read_bytes())
    for name,content in values.items():
        (root/name).write_bytes(content);prep['inputs'][name]=sha256(content).hexdigest().upper()
    prep['directed_scope'].update(case='OnePostCrackDirectProbe',source_states=count,
        replay_mode=1,target_strain=target,maximum_iterations=maximum,
        source_prefix_bits_exact=True,no_new_cold_or_capacity_search=True,
        original_binary_source_history=raw,physical_capacity_certified=False)
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('DIRECT PROBE INPUT READY',count,target,maximum,'NO NEW CAPACITY SEARCH')


if __name__=='__main__':
    p=ArgumentParser(description=__doc__)
    p.add_argument('--root',type=Path,required=True);p.add_argument('--parent',type=Path,required=True)
    p.add_argument('--count',type=int,required=True);p.add_argument('--target',type=float,required=True)
    p.add_argument('--maximum',type=int,default=30)
    a=p.parse_args();TestStage(a.root.resolve(),a.parent.resolve(),a.count,a.target,a.maximum)
