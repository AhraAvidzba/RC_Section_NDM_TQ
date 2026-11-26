"""ДЛЯ ТЕСТОВ. Изолирует папку записи Newton от остальных тестовых флагов; численные операторы решателя не меняет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold918AffineBranchRecovery_20261126';root=base/'Cold922ColdNewtonCapture_20261126'
assert json.loads((base/'Cold921AffineRecoveryColdStress_20261126/VerifiedColdStressReplayProof.json').read_bytes())['source_states']==25
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ColdCaptureCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestNewtonRejectedCapture';new='modTestColdNewtonCapture.TestCapture'
assert code.count(old)==3;code=code.replace(old,new)
assert code.replace(new,old)==(parent/name).read_text('utf-8-sig')
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
oldcode=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
entry,=re.findall(r'(?ms)^Public Sub TestNewtonRejectedCapture\(.*?^End Sub',oldcode)
entry=entry.replace('TestNewtonRejectedCapture(', 'TestCapture(',1).replace('mPolishRootTest','mCaptureFolderTest').replace('TestWriteReverseBinaryVector','TestWriteBinaryVector')
writer,=re.findall(r'(?ms)^Private Sub TestWriteReverseBinaryVector\(.*?^End Sub',oldcode)
writer=writer.replace('TestWriteReverseBinaryVector(','TestWriteBinaryVector(',1)
module='''Attribute VB_Name = "modTestColdNewtonCapture"
Option Explicit

' ==========================================================================
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запись уже вычисленных систем и отклонённых направлений Newton.
' Папка независима от остальных исследовательских режимов; материал повторно не вызывается.
' ==========================================================================
Private mCaptureFolderTest As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Однократно задаёт существующую папку записи; не меняет исходные поля, материалы и режимы решателя.
Public Sub TestSetCaptureRoot(ByVal folder As String)
    If Len(folder) = 0 Or Len(mCaptureFolderTest) > 0 Then Err.Raise 5, , "Папка записи Newton должна задаваться один раз."
    If Len(Dir$(folder & "\\Native", vbDirectory)) = 0 Then Err.Raise 5, , "Не найдена папка нативной записи Newton."
    mCaptureFolderTest = folder
End Sub

'''+entry+'''

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает длину и исходные Double без округления в текст; протокол совпадает с ранее проверенной записью вектора.
'''+writer+'\n'
values['tests/modTestColdNewtonCapture.bas']=module.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames @('CCoupledNewtonSolver','modTestColdNewtonCapture')",1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
assert 'modTestColdNewtonCapture' not in prep['modules'];prep['modules'].append('modTestColdNewtonCapture')
prep['directed_scope']=dict(case='IndependentColdNewtonCaptureModule',actual_import_this_pass=['CCoupledNewtonSolver','modTestColdNewtonCapture'],
 actual_component_count=257,numerical_Newton_body_exact918_except_test_callback_module=True,
 test_capture_and_binary_writer_same_as_existing_verified_methods=True,
 original_test_module_and_flags_unchanged=True,no_additional_material_evaluations=True,
 no_release_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'TWO IMPORTS;ONE NEW TEST MODULE;EXACT NUMERICAL BODY918;ISOLATED CAPTURE FOLDER')
