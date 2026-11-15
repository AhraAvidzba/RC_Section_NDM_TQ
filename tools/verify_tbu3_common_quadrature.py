"""Сверяет материальные точки и принятые состояния общего разбиения TBU3."""

import argparse
from fractions import Fraction
from hashlib import sha256
from itertools import product
import json
from math import isclose, isfinite, sqrt
from pathlib import Path
import re

from validate_tbu3_torque import hoop_figure_curves, interpolate


CASES = {"1x12": (1, 12, 60), "2x6": (2, 6, 87), "3x4": (3, 4, 120)}
AREA = 508 * 410 - 356 * 258
STAGE_COUNT = 4


def expected_concrete():
    """Пересекает каждую ячейку с отверстием и восстанавливает центр остатка."""
    dx, dy = Fraction(508, 12), Fraction(410, 12)
    points = []
    for i, j in product(range(12), repeat=2):
        x, y = -254 + dx * i, -205 + dy * j
        lo_x, hi_x = max(x, -178), min(x + dx, 178)
        lo_y, hi_y = max(y, -129), min(y + dy, 129)
        cut = max(0, hi_x - lo_x) * max(0, hi_y - lo_y)
        area = dx * dy - cut
        if area:
            cx = (dx * dy * (x + dx / 2) - cut * (lo_x + hi_x) / 2) / area
            cy = (dx * dy * (y + dy / 2) - cut * (lo_y + hi_y) / 2) / area
            points.append(("Concrete", 0, float(cx), float(cy), float(area), 0, 0, 0, 0))
    return points


def expected_steel():
    """Строит продольные точки и узлы Гаусса на фиксированной оси хомута."""
    points = [("TBU3_8", 0, 206 * i / 3, -154, 510, 0, 0, 1, 0) for i in range(-3, 4)]
    points.extend(("TBU3_4", 0, 216 * i, 167, 129, 0, 0, 1, 0) for i in range(-1, 2))
    points.extend(("TBU3_3", 0, x, 78 * i, 71, 0, 0, 1, 0) for x, i in product((-218, 218), range(-1, 2)))

    root = sqrt(6 / 5)
    nodes = (sqrt(3 / 7 - 2 / 7 * root), sqrt(3 / 7 + 2 / 7 * root))
    weights = ((18 + sqrt(30)) / 36, (18 - sqrt(30)) / 36)
    for horizontal, side in product((True, False), (-1, 1)):
        fixed, end, cut = (173, 222, 254 / 3) if horizontal else (222, 173, 205 / 3)
        edges = (-end, -cut, 0, cut, end)
        direction = (1, 0, 0, 0) if horizontal else (0, 1, 0, 0)
        for lo, hi in zip(edges, edges[1:]):
            for (node, weight), sign in product(zip(nodes, weights, strict=True), (-1, 1)):
                active = (lo + hi + sign * node * (hi - lo)) / 2
                x, y = (active, side * fixed) if horizontal else (side * fixed, active)
                area = (129 / 76) * (hi - lo) / 2 * weight
                points.append(("TBU3_HOOP", 1, x, y, area, *direction))
    return points


def assert_closure(path, report):
    """Удостоверяет, что завершились обёртка и её собственный Excel."""
    process = re.search(r"^TEST_EXCEL_PROCESS: id=(\d+); startTicks=(\d+)\s*$", report, re.M)
    assert process, f"В протоколе отсутствует собственный PID {path}"
    pid, started = map(int, process.groups())
    closure_path = path.with_name(path.name + ".ProcessClosure.json")
    record = json.loads(closure_path.read_text(encoding="utf-8-sig"))
    assert record["Passed"] and record["ChildHasExited"] and record["ChildExitCode"] == 0, path
    closure = record["Closure"]
    identity, final = closure["Identity"], closure["FinalState"]
    assert closure["Passed"] and identity["Valid"] and final["State"] == "Absent", path
    assert pid == identity["PID"] == final["PID"], path
    assert started == int(identity["Ticks"]) == int(final["StartTicksUTC"]), path


def native_report(path, source_hash=None):
    """Разбирает только завершённый протокол с неизменным источником."""
    report = path.read_text(encoding="utf-8-sig")
    assert "WATCHDOG_COMPLETED: exit=0" in report and "SOURCE_UNCHANGED: True" in report, path
    assert "RUNTIME ERROR:" not in report, path
    if source_hash:
        assert f"SOURCE_SHA256: {source_hash.upper()}" in report, path
    assert_closure(path, report)
    assert "TBU3_GAUGE_X_MM;-222;222" in report and "TBU3_GAUGE_BASE_MM;-76;76" in report, path
    assert "TBU3_HOOP_SPACING_MM;76" in report, path
    return report.splitlines()


def read_case(path, label, source_hash, expected):
    """Проверяет каждую точку и шесть сил во всех принятых состояниях."""
    lines = native_report(path, source_hash)
    spans, cuts, degree_count = CASES[label]
    assert f"TBU3_GAUGE_RESOLUTION;{spans};{cuts};96;176;{degree_count}" in lines, label
    reached_final_stage = "LAST_REQUESTED_STAGE_ACCEPTED: True" in lines
    if not reached_final_stage:
        assert "LAST_REQUESTED_STAGE_ACCEPTED: False" in lines, label
    stop_diagnostics = [line for line in lines if line.startswith(("DIAGNOSTICS:", "TWIST_STOP:"))]
    assert reached_final_stage or stop_diagnostics, label
    points, accepted, integrals, gauges = [], [], [], []
    checked_replays = 0
    for line in lines:
        columns = line.split(";")
        if columns[0] == "TBU3_QUADRATURE":
            assert int(columns[1]) == len(points) + 1, label
            points.append((columns[2], int(columns[3]), *map(float, columns[4:])))
        elif columns[0] == "TBU3_INTEGRALS":
            integrals.append(tuple(map(float, columns[1:])))
        elif columns[0] == "TBU3_GAUGE_AVERAGE":
            gauges.append(tuple(map(float, columns[1:])))
        elif len(columns) == 11 and re.fullmatch(r"(?:Twist)?\d+", columns[0]) and columns[-1] == "Accepted":
            accepted.append((columns[0], *map(float, columns[1:-1])))
        elif len(columns) == 11 and columns[0].isdigit() and columns[-1] == "Unresolved":
            # Силовая остановка повторно выводит последний принятый набор усилий.
            if accepted and columns[1]:
                assert tuple(map(float, columns[1:9])) == accepted[-1][1:9], label
                checked_replays += 1

    def point_key(point):
        return (*point[:2], *(round(value, 8) for value in point[2:4]))

    point_error = 0.0
    assert len(points) == len(expected) == 176, label
    for actual, target in zip(sorted(points, key=point_key), expected, strict=True):
        assert actual[:2] == target[:2], (label, actual, target)
        error = max(abs(a - b) for a, b in zip(actual[2:], target[2:], strict=True))
        assert error < 1e-8, (label, actual, target)
        point_error = max(point_error, error)

    count = len(accepted)
    assert count == len(integrals) == len(gauges) >= 3, label
    assert f"TOTAL_TBU3_FORCE_EQUILIBRIUM: passed={count + checked_replays}; failed=0" in lines, label
    assert [row[0] for row in accepted[:3]] == ["0", "1", "2"], label
    if reached_final_stage:
        assert count == STAGE_COUNT and accepted[-1][0] == "3", label
    assert all(right[1] > left[1] for left, right in zip(accepted, accepted[1:])), label
    residuals = []
    divisors = (AREA, AREA * 508, AREA * 508, AREA * 508, AREA, AREA)
    for force, computed, gauge in zip(accepted, integrals, gauges, strict=True):
        step, stage = force[:2]
        assert all(map(isfinite, (*force[1:], *computed, *gauge))), (label, stage)
        assert not step.isdigit() or int(step) == stage, (label, stage)
        assert computed[0] == stage and gauge[:3] == force[1:4], label
        target = (0, -29.6e6 - 14.87e6 * stage, 0, 10.35e6 * stage, 0, 0)
        residual = max(abs(a - b) / divisor for a, b, divisor in zip(computed[1:], target, divisors, strict=True))
        assert isclose(residual, force[8], rel_tol=0, abs_tol=2e-13) and residual < 5e-6, (label, stage)
        assert isclose(-computed[2] / 1e6, gauge[1], abs_tol=1e-9, rel_tol=0), label
        assert isclose(computed[4] / 1e6, gauge[2], abs_tol=1e-9, rel_tol=0), label
        residuals.append(residual)

    return points, dict(report=str(path), degrees=degree_count, maximum_point_error=point_error,
                        reached_requested_stage=reached_final_stage, stop_diagnostics=stop_diagnostics,
                        resultants=integrals, residual_mpa=residuals, gauge_average=gauges,
                        stage_seconds=[row[9] for row in accepted])


def compare(source, evidence, stock, literature):
    """Сопоставляет базисы на общем T и отделяет эффект разбиения при 2x6."""
    expected = sorted(expected_concrete() + expected_steel(),
                      key=lambda point: (*point[:2], *(round(value, 8) for value in point[2:4])))
    source_hash = sha256(source.read_bytes()).hexdigest()
    points, results = {}, {}
    for label in CASES:
        points[label], results[label] = read_case(evidence / f"Native{label}" / "Native.txt", label, source_hash, expected)
    order_error = 0.0
    for label in ("2x6", "3x4"):
        for a, b in zip(points["1x12"], points[label], strict=True):
            assert a[:2] == b[:2], label
            order_error = max(order_error, *(abs(v - w) for v, w in zip(a[2:], b[2:], strict=True)))
    assert order_error < 1e-8

    stock_lines = native_report(stock)
    assert "TBU3_GAUGE_RESOLUTION;2;6;96;144;87" in stock_lines
    stock_gauges = [tuple(map(float, line.split(";")[1:])) for line in stock_lines if line.startswith("TBU3_GAUGE_AVERAGE;")]
    assert len(stock_gauges) >= STAGE_COUNT and [row[0] for row in stock_gauges[:STAGE_COUNT]] == list(range(STAGE_COUNT))
    assert "LAST_REQUESTED_STAGE_ACCEPTED: True" in stock_lines

    experiment = hoop_figure_curves(literature)["experiment"]
    series = {label: result["gauge_average"] for label, result in results.items()}
    series["2x6_stock_hoop"] = stock_gauges
    comparisons = []
    for stage in (2, 3):
        participants = {label: rows for label, rows in series.items() if any(row[0] == stage for row in rows)}
        if len(participants) < 2:
            continue
        torque = min(row[2] for rows in participants.values() for row in rows if row[0] == stage)
        readings = {label: [1e6 * interpolate([(row[2], row[side]) for row in rows], torque) for side in (3, 4)]
                    for label, rows in participants.items()}
        values = {label: sum(reading) / 2 for label, reading in readings.items()}
        delta_hoop = values["2x6"] - values["2x6_stock_hoop"] if {"2x6", "2x6_stock_hoop"} <= values.keys() else None
        basis = [value for label, value in values.items() if label in CASES]
        comparisons.append(dict(stage=stage, common_torque_knm=torque, gauge_microstrain=readings,
                                experiment_microstrain=1e6 * interpolate(experiment, torque),
                                hoop_quadrature_effect_2x6_microstrain=delta_hoop,
                                common_quadrature_basis_range_microstrain=max(basis) - min(basis) if len(basis) > 1 else None))
    return dict(evidence_status="PASS", all_cases_reached_requested_stage=all(r["reached_requested_stage"] for r in results.values()),
                source_sha256=source_hash, concrete_points=96, longitudinal_points=16, hoop_points=64,
                ordered_points_maximum_difference=order_error, cases=results, comparisons=comparisons)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="Проверочная книга с общим разбиением")
    parser.add_argument("evidence", type=Path, help="Папка протоколов Native1x12, Native2x6, Native3x4")
    parser.add_argument("stock", type=Path, help="Нативный протокол 2x6 со штатным разбиением хомута")
    parser.add_argument("literature", type=Path, help="PDF Kagermanov–Ceresa с графиком измерения")
    options = parser.parse_args()
    print(json.dumps(compare(options.source, options.evidence, options.stock, options.literature), indent=2, ensure_ascii=False))
