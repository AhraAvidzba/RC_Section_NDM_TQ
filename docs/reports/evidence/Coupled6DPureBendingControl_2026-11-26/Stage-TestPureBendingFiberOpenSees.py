"""ДЛЯ ТЕСТОВ. Закрепляет внешний контроль нескольких моментов около независимо проверенного плоского предела."""
from pathlib import Path
from hashlib import sha256
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1005PureBendingReference_20261126';root=base/'Cold1007PureBendingFiberOpenSees_20261126';assert not root.exists()
proof=json.loads((parent/'PureBendingReferenceProof.json').read_bytes());assert proof['planar_reference_physical_limit_verified']
root.mkdir();(root/'Input').mkdir()
ratios=[.25,.5,.75,.9,.95,.98,1.,1.02,1.05,1.1,1.2]
files={'Input/PlanarReference.bin':(parent/'Native/PlanarReference.bin').read_bytes(),
       'Input/PureBendingReferenceProof.json':(parent/'PureBendingReferenceProof.json').read_bytes(),
       'Input/Cases.json':(json.dumps(dict(ratios=ratios,reference_Mx_Nmm=proof['Mx_ultimate_Nmm'],cold_independent_each_case=True,
            matching_planar_stress_strain_curves=True,physical_limits_checked_separately=True),indent=2,sort_keys=True)+'\n').encode(),
       'Run-TestPureBendingFiberOpenSees.py':(here/'Run-TestPureBendingFiberOpenSees.py').read_bytes(),
       'Guard-IsolatedOffice.ps1':(parent/'Guard-IsolatedOffice.ps1').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in files.items():(root/name).write_bytes(data)
(root/'Preparation.json').write_text(json.dumps(dict(inputs={name:sha256(data).hexdigest().upper() for name,data in files.items()},
    selected_ratios=ratios,method='OpenSees ElasticMultiLinear fibers in zeroLengthSection',
    actual_reference_book_sha256=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))['WorkbookSHA256'],
    no_future_external_equilibrium_input=True),indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('PREPARED1007 FIBER OPENSEES;11 INDEPENDENT MOMENTS;EXACT PLANAR CURVES/GEOMETRY')
