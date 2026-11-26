"""ДЛЯ ТЕСТОВ. Независимо проверяет ранг фактической неразрешённой сопряжённой матрицы через SVD NumPy."""
from argparse import ArgumentParser
import csv
import json
from pathlib import Path
import numpy as np


def TestRank(root):
    """Различает несовместность и неединственность внутренних координат; не объявляет физический предел или разрешение шага."""
    folder = root / 'Native'
    stage = list(csv.reader((folder / 'HAdjointFailureStage.csv').read_text('cp1251').splitlines(), delimiter=';'))
    result = {'first_failure_stage': stage[0], 'physical_capacity_certified': False}
    file = folder / 'HAdjointFailure.csv'
    if not file.exists():
        result['linear_matrix_captured'] = False
        return result
    rows = list(csv.reader(file.read_text('cp1251').splitlines(), delimiter=';'))
    count = int(next(row[1] for row in rows if row[0] == 'DIM'))
    matrix = np.zeros((count, count)); rhs = np.zeros(count)
    for kind, *row in rows:
        if kind == 'A': matrix[int(row[1])-1, int(row[0])-1] = float(row[2])
        elif kind == 'B': rhs[int(row[0])-1] = float(row[1])
    assert np.array_equal(rhs, np.eye(count)[-1]) and np.all(np.isfinite(matrix))
    columns = np.max(abs(matrix), axis=0); columns[columns == 0] = 1.
    scaled = matrix / columns
    scales = np.max(abs(scaled), axis=1); scales[scales == 0] = 1.
    scaled = scaled / scales[:, None]; target = rhs / scales
    u, singular, vt = np.linalg.svd(scaled, full_matrices=True)
    estimates = []
    for relative in (1e-10, 1e-12, 1e-14):
        rank = int(np.sum(singular > relative * singular[0]))
        projection = u[:, :rank].T @ target
        weights = (vt[:rank].T @ (projection / singular[:rank])) / columns
        residual = matrix @ weights - rhs
        estimates.append(dict(relative_rank_threshold=relative, rank=rank,
            compatibility_residual=float(np.max(abs(target-u[:, :rank] @ projection))),
            original_residual=float(np.max(abs(residual))), adjoint_l1=float(np.sum(abs(weights))),
            max_weight=float(np.max(abs(weights)))))
    result.update(linear_matrix_captured=True, size=count, singular_values=singular.tolist(), estimates=estimates,
        eigen_or_SVD_used_in_working_code=False)
    return result


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__); parser.add_argument('--root', type=Path, required=True)
    args = parser.parse_args(); result = TestRank(args.root.resolve())
    (args.root / 'HAdjointRankProof.json').write_bytes((json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + '\n').encode())
    print('H ADJOINT RANK CONTROL', {key:value for key,value in result.items() if key != 'singular_values'})
