"""ДЛЯ ТЕСТОВ. Задаёт численные варианты энергетической координаты данными, без повторной компиляции одного VBA-проекта."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
from runpy import run_path
import shutil
import struct


def TestOptions(root,step,newton,energy_fraction,strength_mode):
    """Меняет только четыре численных параметра теста; точность сил и закон материала остаются исходными."""
    assert 1e-8<=step<=1e-3 and newton in (0,1) and 1e-6<=energy_fraction<=.01 and strength_mode in (0,1)
    values=(step,float(newton),energy_fraction,float(strength_mode))
    name='Input/EnvelopeOptions.bin'; (root/name).write_bytes(struct.pack('<i4d',4,*values))
    prep=json.loads((root/'Preparation.json').read_bytes())
    prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update(case='EnvelopeNumericalVariant',finite_difference_relative_step=step,
        complete_Newton_line_search=bool(newton),initial_relative_energy_step=energy_fraction,
        numerical_control_reference_strength=('ft' if strength_mode==0 else 'fc'),
        actual_candidate_and_current_lambda_bounds_required=False,actual_lambda_uncertainty_still_recorded=False,
        actual_candidate_control_projection_logged_and_preserved=False,
        nonlinear_control_target_is_energy=True,original_force_tolerance=True)
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')


def TestStage(repo,root,source,checkpoint,binary,step,newton,energy_fraction,strength_mode):
    """Подготавливает один исходник, читающий настройки из проверяемого двоичного входа."""
    here=Path(__file__).resolve().parent
    run_path(str(here/'Stage-TestEnvelopeEnergy.py'))['TestStage'](repo,root,source,checkpoint,binary,True,False)
    prep=json.loads((root/'Preparation.json').read_bytes())
    name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
    anchor='    Print #handle, "MODE;EnvelopeConsumedEnergyCoordinate;OriginalForceTolerance;CDPOnly;NotTotalThermodynamicDissipation;NotCapacity"'
    assert code.count(anchor)==1
    code=code.replace(anchor,anchor+'''
    Dim testNumerics As Variant, inputHandle As Integer: inputHandle = FreeFile
    Open mPolishRootTest & "\\Input\\EnvelopeOptions.bin" For Binary Access Read As #inputHandle
    testNumerics = TestReadPolishVector(inputHandle)
    If Seek(inputHandle) <> LOF(inputHandle)+1 Then Err.Raise 5, , "Лишние байты численных настроек энергетического теста."
    Close #inputHandle
    If LBound(testNumerics) <> 1 Or UBound(testNumerics) <> 4 Then Err.Raise 5, , "Нужно четыре численных параметра энергетического теста."
    If testNumerics(1)<0.00000001 Or testNumerics(1)>0.001 Then Err.Raise 5, , "Шаг производной энергетического теста вне диапазона."
    If (testNumerics(2)<>0# And testNumerics(2)<>1#) Or (testNumerics(4)<>0# And testNumerics(4)<>1#) Then Err.Raise 5, , "Недопустимый режим численного теста."
    If testNumerics(3)<0.000001 Or testNumerics(3)>0.01 Then Err.Raise 5, , "Приращение энергии энергетического теста вне диапазона."
    Print #handle, "NUMERICS" & TestHardPreparedVector(testNumerics)
''',1)
    anchor='    rowScale = calibration(0)/calibration(3)'
    assert code.count(anchor)==1
    code=code.replace(anchor,anchor+'\n    If testNumerics(4)=1# Then rowScale = calibration(0)/calibration(2)',1)
    assert code.count('baseline*0.001')==2
    code=code.replace('baseline*0.001','baseline*testNumerics(3)')
    anchor='solution,relativeDifferenceStep:=0.000001, _'
    assert code.count(anchor)==1
    code=code.replace(anchor,'solution,relativeDifferenceStep:=testNumerics(1), _',1)
    anchor='automaticScaling:=True,freshJacobian:=True,completeNewtonLineSearch:=True)'
    assert code.count(anchor)==1
    code=code.replace(anchor,'automaticScaling:=True,freshJacobian:=True,completeNewtonLineSearch:=(testNumerics(2)<>0#))',1)
    assert max(map(len,code.splitlines()))<1024
    (root/name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'));prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    name=Path(__file__).name;(root/name).write_bytes(Path(__file__).read_bytes());prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    TestOptions(root,step,newton,energy_fraction,strength_mode)
    print('ENERGY NUMERICS INPUT READY; SAME PHYSICS AND FORCE TOLERANCE')


def TestClone(root,parent,step,newton,energy_fraction,strength_mode,binary=None):
    """Переиспользует только неизменяемый собранный код; новая нативная сессия восстанавливает всю историю с нуля."""
    assert not root.exists() and root.parent==parent.parent
    compilation=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    closure=json.loads((parent/'CompileClosure.json').read_bytes())
    assert compilation['Compiled'] and not compilation['ProductionChanged'] and closure['State']=='Absent'
    assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compilation['WorkbookSHA256']
    shutil.copytree(parent,root,ignore=shutil.ignore_patterns('Temp','Native','Run.ndjson'))
    (root/'Temp').mkdir();(root/'Native').mkdir()
    prep=json.loads((root/'Preparation.json').read_bytes())
    if binary is not None:
        import numpy as np
        read=run_path(str(Path(__file__).resolve().parent/'Check-TestActualReverseCapture.py'))['TestReadBinary']
        before,after=read(parent/'Input/Reverse1.bin'),read(binary)
        for key in ('calibration','options','offset','direction'):
            assert np.array_equal(before[key],after[key]),key
        count=min(len(before['fields']),len(after['fields']))
        assert np.array_equal(before['fields'][:count],after['fields'][:count])
        assert np.array_equal(before['lambdas'][:count],after['lambdas'][:count])
        name='Input/Reverse1.bin';(root/name).write_bytes(binary.read_bytes())
        prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
        prep['directed_scope'].update(current_input_states=len(after['fields']),current_lambda=float(after['lambdas'][-1]),
            binary_input_origin=str(binary),shared_native_input_prefix_verified=True)
    name=Path(__file__).name;(root/name).write_bytes(Path(__file__).read_bytes())
    prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    TestOptions(root,step,newton,energy_fraction,strength_mode)
    (root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent),
        compiled_workbook_sha256=compilation['WorkbookSHA256'],parent_compilation_and_closure_preserved=True,
        fresh_compilation_claimed=False,mutable_material_state_reused=False,native_history_replay_required=True),indent=2)+'\n',encoding='utf-8')
    print('IMMUTABLE COMPILED SOURCE REUSED; FRESH MATERIAL HISTORY REQUIRED')


if __name__=='__main__':
    p=ArgumentParser(description=__doc__);p.add_argument('--repo',type=Path,default=Path.cwd());p.add_argument('--root',type=Path,required=True)
    for name in ('source','checkpoint','binary','parent'):p.add_argument('--'+name,type=Path)
    p.add_argument('--step',type=float,default=1e-4);p.add_argument('--newton',type=int,default=1)
    p.add_argument('--energy-fraction',type=float,default=.001);p.add_argument('--strength-mode',type=int,default=0)
    a=p.parse_args()
    if a.parent:TestClone(a.root.resolve(),a.parent.resolve(),a.step,a.newton,a.energy_fraction,a.strength_mode,a.binary.resolve() if a.binary else None)
    else:TestStage(a.repo.resolve(),a.root.resolve(),a.source.resolve(),a.checkpoint.resolve(),a.binary.resolve(),a.step,a.newton,a.energy_fraction,a.strength_mode)
