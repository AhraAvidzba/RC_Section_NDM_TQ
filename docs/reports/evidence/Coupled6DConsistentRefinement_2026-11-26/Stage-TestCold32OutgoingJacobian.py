"""ДЛЯ ТЕСТОВ. Добавляет чтение фактического рабочего касательного прогноза и уровня1/1024 на той же холодной истории32."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Stage-TestCold32DirectionalJacobian.py'
code=file.read_text('utf-8-sig').replace('Cold899Cold32DirectionalJacobian_20261126','Cold900Cold32OutgoingJacobian_20261126')
code=code.replace('Array(0#, 0.015625, 1#, 8#)','Array(0#, 0.0009765625, 0.015625, 1#, 8#)')
code=code.replace('levels=[0,1/64,1,8]','levels=[0,1/1024,1/64,1,8]')
anchor='    TestWriteReverseBinaryVector handle, projection\n'
assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''    Dim outgoing() As Double, outgoingProbe() As Double, outgoingReady As Boolean
    outgoingReady = mReverseSolverTest.TrySpatialLoadTangent(mReverseBaseTest, mReverseDirectionTest, previous, accepted, factor, outgoing, outgoingProbe)
    TestWriteReverseBinaryVector handle, Array(Abs(CLng(outgoingReady)))
    If Not outgoingReady Then Err.Raise 5, , "Рабочий касательный прогноз не прочитан; отказ не является разрушением."
    TestWriteReverseBinaryVector handle, outgoing
    TestWriteReverseBinaryVector handle, outgoingProbe
''',1)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'cold32_outgoing_jacobian_staging'})
