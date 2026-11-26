"""ДЛЯ ТЕСТОВ. Готовит холодный контроль сечения на таблицах объёмного OpenSees; не переносит готовые поля или материальные истории."""
from pathlib import Path
from hashlib import sha256
import ast,json,math,shutil

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold988FixedForceExternal_20261126';root=base/'Cold1000ASDHardShearCold_20261126'
assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged']
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
assert json.loads((parent/'CompilePointClosure.json').read_bytes())['State']=='Absent'
root.mkdir()
for directory in ('src','tests','tools'):
    shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','Input','docs/regression','workbook/output'):(root/directory).mkdir(parents=True,exist_ok=True)
files={
    'docs/regression/Diagnostic.xlsm':(parent/'docs/regression/Diagnostic.xlsm').read_bytes(),
    'workbook/output/RC_Section_NDM.xlsm':(repo/'workbook/output/RC_Section_NDM.xlsm').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'CompiledSourceReuse.json':(json.dumps(dict(parent=str(parent)),indent=2)+'\n').encode(),
    'Guard-IsolatedOffice.ps1':(parent/'Guard-IsolatedOffice.ps1').read_bytes(),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    Path(__file__).name:Path(__file__).read_bytes(),
    'TestASDHardShearCold.bas.txt':(here/'TestASDHardShearCold.bas.txt').read_bytes(),
}
source=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
marker='Public Function TestStartFixedForceComparison('
assert source.count(marker)==1
start=source.index(marker)
source=source[:start]+(here/'TestASDHardShearCold.bas.txt').read_text('utf-8-sig')+'\n\n'+source[start:]
assert max(map(len,source.splitlines()))<1024
files['tests/modTestCoupledElasticState.bas']=source.replace('\n','\r\n').encode('utf-8-sig')
# Применяется буквально тот же генератор таблиц, который исполнялся в объёмном контроле 999.
reference=(base/'Cold999OpenSeesBufferedCore_20261126/validate_opensees_hardshear.py').read_bytes()
tree=ast.parse(reference);function=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='envelope')
namespace=dict(math=math,E=30000.);exec(compile(ast.Module(body=[function],type_ignores=[]),'<pinned-envelope>','exec'),namespace)
envelope=namespace['envelope'];dc=.423*(1.-30./(.002*30000.));plastic=.002-30./((1.-dc)*30000.)
curves={'T':envelope(3.,3.,0.,3.*125./.12,.12/125.),
        'C':envelope(16.,30.,plastic,-math.log1p(-dc)/plastic,80./125.)}
lines=['ASDFlow1;30000;0.2']
for name,rows in curves.items():
    floor=min(row[1] for row in rows[1:])*1e-6
    lines.append(f'BRANCH;{name};{len(rows)};{floor:.17g}')
    for strain,nominal,damage in rows:lines.append(';'.join(format(v,'.17g') for v in (strain,nominal,nominal/(1.-damage))))
files['Input/matched.csv']=('\r\n'.join(lines)+'\r\n').encode()
files['Input/MaterialContract.json']=(json.dumps(dict(reference999_sha256=sha256(reference).hexdigest().upper(),
    concrete_E_MPa=30000.,poisson=.2,uniaxial_tables=curves,
    same_ASD_equivalent_stress_surfaces=True,
    integration_difference='NDM closes the ASD update at fixed strain; OpenSees uses its native single material update',
    steel_monotonic_backbone=dict(E_MPa=200000.,long_fy_tension_MPa=400.,long_fy_compression_MPa=350.,
        hoop_fy_MPa=240.,all_postyield_slopes_MPa=1500.),
    steel_history_difference='NDM uses kinematic hardening; OpenSees longitudinal Hysteretic and hoop Steel01 are not identical cyclic laws',
    geometry_difference='NDM subtracts displaced concrete at steel quadrature; OpenSees embedded trusses overlap solid concrete',
    no_future_equilibrium_fields_or_committed_material_states=True,physical_capacity_certified=False),indent=2,sort_keys=True)+'\n').encode()
runner=(parent/'Run-TestRealColdReplay.ps1').read_text('utf-8-sig')
runner=runner.replace('WorkingGuideTargetedRegressions','ASDHardShearCold').replace('RunTestReversePolish','TestRunASDHardShearCold')
files['Run-TestASDHardShearCold.ps1']=runner.encode('utf-8-sig')
for name,data in files.items():(root/name).write_bytes(data)
prep=json.loads((parent/'Preparation.json').read_bytes())
prep['inputs']={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper()
                for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
prep['directed_scope']=dict(case='ASDHardShearCold',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all256_numerical_components_exact988=True,
    no_future_equilibrium_input=True,virgin_material_history=True,spatial_law_family_aligned_but_integration_not_identical=True,
    original_force_tolerance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1000 COLD ASD;NO READY FIELDS;SAME TABLES999;INTEGRATION/STEEL/EMBEDDING DIFFERENCES EXPLICIT')
