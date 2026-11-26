"""ДЛЯ ТЕСТОВ. Применяет независимые алгоритмы SciPy к неизменённому материальному отклику VBA."""

from argparse import ArgumentParser
import ctypes
from ctypes import wintypes
from hashlib import sha256
import json
import os
from pathlib import Path
import sys
import time
import traceback
import numpy as np
import scipy
from scipy.optimize import root, least_squares


def TestStartTicks():
    """Читает точное время запуска собственного процесса в единицах .NET, без округления Unix-времени."""
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.GetCurrentProcess.restype = wintypes.HANDLE
    kernel.GetProcessTimes.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.FILETIME),
        ctypes.POINTER(wintypes.FILETIME), ctypes.POINTER(wintypes.FILETIME), ctypes.POINTER(wintypes.FILETIME)]
    creation, ending, cpu_kernel, cpu_user = (wintypes.FILETIME() for _ in range(4))
    if not kernel.GetProcessTimes(kernel.GetCurrentProcess(), ctypes.byref(creation), ctypes.byref(ending),
        ctypes.byref(cpu_kernel), ctypes.byref(cpu_user)):
        raise ctypes.WinError(ctypes.get_last_error())
    return 504911232000000000 + (creation.dwHighDateTime << 32) + creation.dwLowDateTime


def TestMain(folder):
    """Не создаёт Excel: единственный собственный VBA-хост обрабатывает запросы через частную папку."""
    started = time.monotonic()
    owner = {'pid': os.getpid(), 'start_ticks': str(TestStartTicks()), 'script': str(Path(__file__).resolve()),
        'script_sha256': sha256(Path(__file__).read_bytes()).hexdigest().upper(), 'python': sys.executable,
        'scipy': scipy.__version__, 'bridge': str(folder), 'created_excel': False}
    (folder / 'PythonOwner.json').write_text(json.dumps(owner, indent=2) + '\n', encoding='utf-8')
    sequence = 0
    def TestRequest(operation, values=None):
        """Передаёт числа с 17 значащими цифрами; ответы содержат исходные IEEE Double, без CSV-округления."""
        nonlocal sequence
        if time.monotonic() - started > 180:
            raise TimeoutError('Independent optimizer time budget expired')
        if values is not None and not np.all(np.isfinite(values)):
            raise ValueError('Nonfinite external field is not a material evaluation')
        sequence += 1
        request = operation + ';' + str(sequence)
        if values is not None:
            request += ';' + ';'.join(format(float(value), '.17g') for value in values)
        with (folder / 'Requests.ndjson').open('a', encoding='utf-8') as log:
            log.write(json.dumps({'sequence': sequence, 'operation': operation,
                'values': None if values is None else np.asarray(values).tolist()}) + '\n')
        temporary = folder / 'request.tmp'
        temporary.write_text(request + '\n', encoding='ascii')
        temporary.replace(folder / 'request.csv')
        response, error = folder / f'response-{sequence}.bin', folder / f'response-{sequence}.err'
        pending = time.monotonic()
        while not response.exists():
            if error.exists():
                raise RuntimeError('VBA evaluation: ' + error.read_text('cp1251').strip())
            if time.monotonic() - pending > 30 or time.monotonic() - started > 180:
                raise TimeoutError('VBA response time budget expired')
            time.sleep(.002)
        return np.frombuffer(response.read_bytes(), dtype='<f8').copy()
    result = {'physical_capacity_certified': False, 'independent_material_reintegration': False,
        'independent_numerical_algorithms': True, 'methods': [], 'native_validation_required': True}
    best = None
    best_norm = float('inf')
    try:
        seed = TestRequest('S')
        scale = TestRequest('C')
        assert seed.shape == scale.shape == (61,) and np.all(scale > 0)
        best = seed.copy()
        cached_x, cached_r = None, None
        def TestResidual(y):
            """Все законы и история остаются в VBA; точный кэш действует только для совпадающего полного вектора."""
            nonlocal best, best_norm, cached_x, cached_r
            x = seed + scale * y
            if cached_x is not None and np.array_equal(x, cached_x):
                return cached_r.copy()
            value = TestRequest('R', x)
            assert value.shape == (61,) and np.all(np.isfinite(value))
            cached_x, cached_r = x.copy(), value.copy()
            norm = float(np.max(np.abs(value)))
            if norm < best_norm:
                best, best_norm = x.copy(), norm
            with (folder / 'Progress.ndjson').open('a', encoding='utf-8') as log:
                log.write(json.dumps({'sequence': sequence, 'seconds': time.monotonic() - started,
                    'residual_MPa': norm, 'best_residual_MPa': best_norm, 'lambda': float(x[-1])}) + '\n')
            return value
        def TestJacobian(y):
            """Читает обычную касательную VBA в порядке строк и учитывает только масштабы неизвестных."""
            x = seed + scale * y
            matrix = TestRequest('J', x)
            assert matrix.shape == (61 * 61,) and np.all(np.isfinite(matrix))
            return matrix.reshape(61, 61) * scale[None, :]
        TestResidual(np.zeros(61))
        for method in ('hybr', 'trf'):
            if best_norm <= 1e-8 or time.monotonic() - started > 140:
                break
            initial = (best - seed) / scale
            try:
                if method == 'hybr':
                    solved = root(TestResidual, initial, jac=TestJacobian, method='hybr',
                        options={'xtol': 1e-11, 'maxfev': 350, 'factor': .2, 'diag': np.ones(61)})
                else:
                    solved = least_squares(TestResidual, initial, jac=TestJacobian, method='trf',
                        ftol=1e-12, xtol=1e-12, gtol=1e-12, max_nfev=250, x_scale='jac')
                result['methods'].append({'method': method, 'success': bool(solved.success),
                    'message': str(solved.message), 'nfev': int(solved.nfev),
                    'residual_MPa': float(np.max(np.abs(solved.fun))), 'best_residual_MPa': best_norm})
            except Exception as error:
                result['methods'].append({'method': method, 'success': False, 'error': repr(error),
                    'best_residual_MPa': best_norm})
        result['best_residual_MPa'] = best_norm
        result['best_lambda'] = float(best[-1])
        best.astype('<f8').tofile(folder / 'BestCandidate.bin')
        result['native_validation'] = TestRequest('V', best).tolist()
        result['native_equilibrium_found'] = bool(result['native_validation'][0])
    except Exception as error:
        result['error'] = repr(error)
        result['traceback'] = traceback.format_exc()
        result['native_equilibrium_found'] = False
    finally:
        result['seconds'] = time.monotonic() - started
        result['requests'] = sequence
        (folder / 'Result.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        final = folder / 'done.tmp'
        final.write_text('NATIVE_EQUILIBRIUM_FOUND\n' if result.get('native_equilibrium_found') else 'NO_CONFIRMED_EQUILIBRIUM\n', encoding='ascii')
        final.replace(folder / 'done.txt')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--bridge', type=Path, required=True)
    TestMain(parser.parse_args().bridge.resolve())
