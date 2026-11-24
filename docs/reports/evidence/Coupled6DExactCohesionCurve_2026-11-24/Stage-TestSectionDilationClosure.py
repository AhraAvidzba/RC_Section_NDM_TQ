"""ДЛЯ ТЕСТОВ. Включает проверенное замыкание дилатансии для вырождающегося растянутого возврата и проверяет полную кривую сечения."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold707SectionDilationClosure_20261124';parent=base/'Cold706PostCrackDilationDense_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'SectionClosureCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
(root/'Matrix/Native').mkdir(parents=True,exist_ok=True)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='src/Materials/CCoupledConcreteResponse.cls';code=(root/path).read_text('utf-8-sig')
payload=(here/'TestPostCrackDilationRoot.cls.txt').read_text('utf-8-sig');assert code.count(payload)==1
generic=payload.replace("' -------------------- ДЛЯ ТЕСТОВ --------------------\n",'')
generic=generic.replace('TestTryPostCrackLocalGuide','TryDilationClosure').replace('TestReadDilationResidual','ReadDilationClosureResidual').replace('TestBisectDilationResidual','BisectDilationClosure')
generic=generic.replace('Private Function TryDilationClosure(ByRef strain() As Double, ByRef stress() As Double, ByRef substeps As Long) As Boolean',
    'Private Function TryDilationClosure(ByRef strain() As Double, ByRef committed As TConcreteHistory, ByRef result As TConcreteHistory, ByRef stress() As Double) As Boolean')
generic=generic.replace('Dim guide As TConcreteHistory, prior As TConcreteHistory, result As TConcreteHistory','Dim guide As TConcreteHistory, prior As TConcreteHistory')
generic=generic.replace('mHistory(1, mCommitted)','committed')
generic=generic.replace('ReadDilationClosureResidual(strain, guide,','ReadDilationClosureResidual(strain, committed, guide,')
generic=generic.replace('Private Function ReadDilationClosureResidual(ByRef strain() As Double, ByRef guide As TConcreteHistory,',
    'Private Function ReadDilationClosureResidual(ByRef strain() As Double, ByRef committed As TConcreteHistory, ByRef guide As TConcreteHistory,')
generic=generic.replace('BisectDilationClosure(strain, guide,','BisectDilationClosure(strain, committed, guide,')
generic=generic.replace('Private Function BisectDilationClosure(ByRef strain() As Double, ByRef guide As TConcreteHistory,',
    'Private Function BisectDilationClosure(ByRef strain() As Double, ByRef committed As TConcreteHistory, ByRef guide As TConcreteHistory,')
generic=generic.replace('ByVal left As Double, ByVal right As Double, ByVal leftResidual As Double, ByRef stress() As Double)',
    'ByVal left As Double, ByVal right As Double, ByVal leftResidual As Double, ByRef result As TConcreteHistory, ByRef stress() As Double)')
generic=generic.replace('leftResidual, stress)','leftResidual, result, stress)').replace('centerResidual, stress)','centerResidual, result, stress)')
generic=generic.replace('Dim result As TConcreteHistory, middle As Double','Dim middle As Double')
generic=generic.replace(': substeps = 1','')
generic=re.sub(r'(?m)^\s*If TryDilationClosure Then substeps = 1\n','',generic)
generic=generic.replace('= UpdatePoint(strain, committed, result, stress, virgin, _','= EvaluatePointState(strain, committed, result, stress, virgin, _')
anchor=' As Boolean\n';at=generic.index(anchor)+len(anchor)
generic=generic[:at]+'    On Error GoTo ArithmeticFailure\n'+generic[at:]
marker='    Next attempt\nEnd Function';assert generic.count(marker)==1
generic=generic.replace(marker,'''    Next attempt
    Exit Function
ArithmeticFailure:
    Dim number As Long, source As String, description As String
    number = Err.Number: source = Err.Source: description = Err.Description
    If number <> 6 And number <> 11 Then Err.Raise number, source, description
    TryDilationClosure = False
End Function''',1)
assert 'substeps' not in generic and 'mHistory(1, mCommitted)' not in generic
code=code.replace(payload,generic+'\n'+(here/'CandidateDilationClosureGate.cls.txt').read_text('utf-8-sig'),1)
method,=re.findall(r'(?ms)^Private Function UpdatePoint\(.*?^End Function',code[:code.index('Private Function TryDilationClosure(')])
updated=method.replace('Function UpdatePoint(','Function EvaluatePointState(',1).replace('UpdatePoint =','EvaluatePointState =')
code=code.replace(method,updated,1)
code=code.replace('holdDilationTest','holdDilation').replace('flowResidualTest','flowResidual').replace('flowReadyTest','flowReady')
anchor='''    If UpdatePoint(strain, committed, updated, stress, virginElastic) Then
        StoreStatePosition updated, strain, stress''';assert code.count(anchor)==1
code=code.replace(anchor,'''    If UpdatePoint(strain, committed, updated, stress, virginElastic) Then
        If NeedsDilationClosure(committed, updated, stress, virginElastic) Then
            Dim closed As TConcreteHistory, closedStress(1 To 6) As Double
            If TryDilationClosure(strain, committed, closed, closedStress) Then
                updated = closed: virginElastic = False
                For i = 1 To N_COMPONENTS: stress(i) = closedStress(i): Next i
            End If
        End If
        StoreStatePosition updated, strain, stress''',1)
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledConcreteResponse'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
for name in (Path(__file__).name,'CandidateDilationClosureGate.cls.txt'):values[name]=(here/name).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='SectionCurveWithLocalDilationClosure',actual_import_this_pass=['CCoupledConcreteResponse'],
    all_concrete_points_and_stations=True,no_reference_to_point1_in_working_helpers=True,
    original_full_step_criteria=True,original_valid_return_retained_if_optional_recovery_fails=True,
    physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('SECTION LOCAL CLOSURE STAGED; GENERAL POINT HISTORY; ORIGINAL FULL STEP AND FAILURE FALLBACK')
