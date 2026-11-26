"""ДЛЯ ТЕСТОВ. Решает независимую кривую сжатия настоящим ASDConcrete3D и тем же TRF/SVD, без использования будущих полей VBA."""
from pathlib import Path
from hashlib import sha256

here = Path(__file__).resolve().parent
source = here / 'Solve-TestPureAxialOpenSees.py'
code = source.read_text('utf-8-sig')
changes = [
    ("output = root / 'Native/IndependentAxialTRF.json'", "output = root / 'Native/IndependentCompressionPeakTRF.json'"),
    ('direction[0] > 0 and np.all(direction[1:] == 0)', 'direction[0] < 0 and np.all(direction[1:] == 0)'),
    ('assert r.position == len(history_raw)', '''assert r.position == len(history_raw)
control = json.loads((root / 'Input/CompressionPeakControl.json').read_bytes())
assert sha256(history_raw).hexdigest().upper() == control['source_native_history_sha256']
assert args.step == control['axial_coordinate_step'] and args.end == control['end_compressive_strain_magnitude']
assert args.seconds == control['wall_limit_seconds']
history = history[:control['native_prefix_states']]'''),
    ('target = float(current[0])', 'target = -float(current[0])'),
    ('return np.r_[target, value[:-1] * scale[1:]]', 'return np.r_[-target, value[:-1] * scale[1:]]'),
    ('dict(axial_strain=target, lambda_value=factor', 'dict(axial_strain=-target, lambda_value=factor'),
    ('except TimeoutError as error:', 'except Exception as error:'),
    ('termination = str(error)', "termination = type(error).__name__ + ': ' + str(error)"),
    ("source_helper_sha256=sha256(Path(__file__).read_bytes()).hexdigest().upper(),",
     "source_helper_sha256=sha256(Path(__file__).read_bytes()).hexdigest().upper(),\n                 common_driver_sha256='" + sha256(source.read_bytes()).hexdigest().upper() + "',")
]
for before, after in changes:
    assert code.count(before) == 1, before
    code = code.replace(before, after, 1)
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestIndependentCompressionPeak'))
