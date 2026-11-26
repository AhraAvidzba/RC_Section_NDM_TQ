"""ДЛЯ ТЕСТОВ. Проверяет точные касательные нового отказа934 с исходной историей25 и сохранённой скалярной строкой."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;file=here/'Check-TestSurfaceMatrices.py'
code=file.read_text('utf-8-sig').replace('Cold931SurfaceMatrices_20261126','Cold935Point11Matrices_20261126')
scope={'__file__':str(file),'__name__':'point11_surface_matrices_test'}
exec(compile(code,str(file),'exec'),scope)
root=scope['root'];compiled=root.parent/'Cold931SurfaceMatrices_20261126'
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(compiled/'docs/regression/Diagnostic.xlsm').read_bytes()
for p in (compiled/'tests').rglob('*'):
    if p.is_file():assert p.read_bytes()==(root/p.relative_to(compiled)).read_bytes()
proof=scope['proof'];proof.update(actual_source_record934=3,source_matrix_refined=True,
    immutable_compiled931_reused=True,no_fresh_compilation=True,point=11,station_z_over_L=-.001)
(root/'SurfaceMatricesProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
