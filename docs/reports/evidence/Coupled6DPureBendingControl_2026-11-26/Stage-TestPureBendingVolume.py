"""ДЛЯ ТЕСТОВ. Готовит объёмный чистый изгиб OpenSees относительно штатного плоского предела, сохраняя пространственный ASD и армирование."""
from pathlib import Path
from hashlib import sha256
import json

repo=Path.cwd();base=Path('D:/CodexScratch/RC_Section_NDM_TQ');here=Path(__file__).resolve().parent
parent=base/'Cold1005PureBendingReference_20261126';root=base/'Cold1008PureBendingVolume_20261126';assert not root.exists()
proof=json.loads((parent/'PureBendingReferenceProof.json').read_bytes());assert proof['planar_reference_physical_limit_verified']
root.mkdir()
for name,data in [('Guard-IsolatedOffice.ps1',(parent/'Guard-IsolatedOffice.ps1').read_bytes()),
                  ('validate_opensees_hardshear.py',(repo/'tools/validate_opensees_hardshear.py').read_bytes()),
                  ('PureBendingReferenceProof.json',(parent/'PureBendingReferenceProof.json').read_bytes()),
                  (Path(__file__).name,Path(__file__).read_bytes())]:(root/name).write_bytes(data)
prep=dict(section_target=[0.,proof['Mx_ultimate_Nmm'],0.,0.,0.,0.],variant='matched',length_mm=800.,cut_mm=400.,dz_mm=25.,
    nonlinear_length_mm=125.,end_condition='balanced-tractions',capture_cut=True,
    requested_factors=[.25,.5,.75,.9,.95,.98,1.,1.02,1.05,1.1,1.2],
    source_sha256=sha256((root/'validate_opensees_hardshear.py').read_bytes()).hexdigest().upper(),
    only_pure_Mx=True,no_gradient_of_M_from_Q=True,auxiliary_elastic_regions_not_homogeneous_beam=True,
    material_diagrams_differ_from_planar_reference=['ASDConcrete3D has tensile cracking and softening','steel postyield tangent1500MPa versus planar plateau'],
    above_stop_requested_factors_not_executed_are_not_failure_proofs=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1008 3D PURE MX;REFERENCE',proof['Mx_ultimate_Nmm'],'NMM;SPATIAL MATERIAL DIFFERENCES EXPLICIT')
