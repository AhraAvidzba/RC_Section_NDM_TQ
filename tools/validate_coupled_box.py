"""Compare a homogeneous TBU3-shaped box against independent elastic FE.

The sectionproperties triangular mesh uses the documented test dimensions,
but omits the reinforcement and uses Poisson's ratio zero. Forces are N,
moments are N mm, stresses are MPa and the torsion constant is mm^4.
"""

import argparse
import json
from importlib.metadata import version
from math import sqrt
from pathlib import Path

from sectionproperties.analysis import Section
from sectionproperties.pre.library import rectangular_section


POINTS = [(0, 167), (0, -167), (216, 0), (-216, 0), (210, 170), (-210, -170)]
FORCE = 100_000
MX = 15_000_000
MY = -11_000_000
TORQUE = 10_000_000
QX = 28_000
QY = -19_000


def finite_element_box(maximum_triangle_area):
    """Solve on an independently triangulated, geometrically exact rectangle pair."""
    outer = rectangular_section(d=410, b=508).shift_section(x_offset=-254, y_offset=-205)
    opening = rectangular_section(d=258, b=356).shift_section(x_offset=-178, y_offset=-129)
    geometry = (outer - opening).create_mesh(mesh_sizes=maximum_triangle_area)
    section = Section(geometry)
    section.calculate_geometric_properties()
    section.calculate_warping_properties()
    # NDM integrates My = integral(x*sigma). Sectionproperties uses its negative.
    stresses = section.get_stress_at_points(POINTS, n=FORCE, mxx=MX, myy=-MY, mzz=TORQUE, vx=QX, vy=QY)
    if any(stress is None for stress in stresses):
        raise ValueError("A control point is outside concrete")
    return {
        "max_triangle_area_mm2": maximum_triangle_area,
        "elements": len(section.elements),
        "area_mm2": float(section.get_area()),
        "j_mm4": float(section.get_j()),
        "stress_MPa": [[float(value) for value in stress] for stress in stresses],
    }


def relative_norm(actual, expected):
    """Compute relative Euclidean error using the control values as the scale."""
    return sqrt(sum((a - e) ** 2 for a, e in zip(actual, expected)) / sum(e**2 for e in expected))


def compare_excel(native_report, reference):
    """Compare the stresses at matching coordinates and keep each basis level."""
    rigidities = {}
    fields = {}
    for line in native_report.read_text(encoding="utf-8-sig").splitlines():
        row = line.split(";")
        if row[0] == "ELASTIC_RIGIDITY":
            rigidities[int(row[1])] = float(row[2])
        elif row[0] == "ELASTIC_STRESS":
            coordinates = tuple(map(float, row[2:4]))
            fields.setdefault(int(row[1]), {})[coordinates] = list(map(float, row[4:]))
    if set(rigidities) != {1, 2, 3} or set(fields) != set(rigidities):
        raise ValueError("The native report must contain all three basis levels")

    comparisons = []
    expected = reference["stress_MPa"]
    for basis in sorted(rigidities):
        field = [fields[basis][point] for point in POINTS]
        comparisons.append({
            "spans_per_direction": basis,
            "j_mm4": rigidities[basis],
            "relative_j_error": abs(rigidities[basis] / reference["j_mm4"] - 1),
            "relative_stress_error": {
                name: relative_norm([value[index] for value in field], [value[index] for value in expected])
                for index, name in enumerate(("sigma_z", "tau_xz", "tau_yz"))
            },
            "stress_MPa": field,
        })
    return comparisons


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--native-report", type=Path)
    args = parser.parse_args()
    meshes = [finite_element_box(area) for area in (400, 100, 25)]
    result = {"sectionproperties": version("sectionproperties"), "coordinates_mm": POINTS, "meshes": meshes}
    if args.native_report:
        result["excel"] = compare_excel(args.native_report, meshes[-1])
    print(json.dumps(result, indent=2, allow_nan=False))


if __name__ == "__main__":
    main()
