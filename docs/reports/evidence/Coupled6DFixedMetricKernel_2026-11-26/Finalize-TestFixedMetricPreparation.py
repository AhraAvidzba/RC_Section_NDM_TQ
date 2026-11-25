"""ДЛЯ ТЕСТОВ. Синхронизирует русские пояснения новых коротких входов и контрольные суммы до единственной компиляции."""
from pathlib import Path
from hashlib import sha256
import json
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold867FixedMetricKernel_20261126')
assert not (root/'CompilePoint.log').exists()
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
for signature,comment in [('Public Sub TestRunFixedMetricArcWalk()',"' -------------------- ДЛЯ ТЕСТОВ --------------------\n    ' Продолжает настоящую историю общим решателем: до32 попыток, шесть долей на ревизию, общий бюджет96."),
 ('Public Sub TestWriteFixedMetricGuide(ByVal attempt As Long, ByVal arc As CCoupledArcContinuation)',"' -------------------- ДЛЯ ТЕСТОВ --------------------\n    ' Записывает диагностические ориентиры версии3 до принятия истории.")]:
    assert code.count(signature)==1
    code=code.replace(signature,signature+'\n    '+comment,1)
data=code.replace('\n','\r\n').encode('utf-8-sig');(root/name).write_bytes(data)
prep=json.loads((root/'Preparation.json').read_bytes());prep['inputs'][name]=sha256(data).hexdigest().upper()
(root/Path(__file__).name).write_bytes(Path(__file__).read_bytes())
prep['inputs'][Path(__file__).name]=sha256(Path(__file__).read_bytes()).hexdigest().upper()
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARATION PINNED;NATIVE VERIFICATION PENDING')
