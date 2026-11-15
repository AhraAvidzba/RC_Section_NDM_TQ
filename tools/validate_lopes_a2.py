"""Независимая упругая проверка номинального сечения A2 на чистое кручение."""

import json
import argparse
from importlib.metadata import version

from sectionproperties.analysis import Section
from sectionproperties.pre import Geometry
from shapely.geometry import Polygon


OUTER = [(-300, -300), (300, -300), (300, 300), (-300, 300)]
POISSON = 0.20
MEASURED_RIGIDITY_KNM2 = 97_950


def finite_element_torsion(maximum_triangle_area, wall):
    """Считает J однородного сечения. Измеренную жёсткость использует после решения."""
    half = 300 - wall
    slope = half - 50
    opening = [(-slope, -half), (slope, -half), (half, -slope), (half, slope),
               (slope, half), (-slope, half), (-half, slope), (-half, -slope)]
    area_exact = 600**2 - (2 * half)**2 + 4 * 50**2 / 2
    geometry = Geometry(Polygon(OUTER, [opening]))
    section = Section(geometry.create_mesh(mesh_sizes=maximum_triangle_area))
    section.calculate_geometric_properties()
    section.calculate_warping_properties()
    area = float(section.get_area())
    if abs(area - area_exact) > area_exact * 1e-10:
        raise ValueError("Площадь сетки не совпадает с площадью принятой геометрии")
    torsion_constant = float(section.get_j())
    return {
        "max_triangle_area_mm2": maximum_triangle_area,
        "elements": len(section.elements),
        "area_mm2": area,
        "j_mm4": torsion_constant,
        "e_implied_mpa": 2 * (1 + POISSON) * MEASURED_RIGIDITY_KNM2 * 1e9 / torsion_constant,
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--wall-mm", type=int, choices=(100, 107), default=100)
    wall = parser.parse_args().wall_mm
    print(json.dumps({"sectionproperties": version("sectionproperties"),
                      "wall_mm": wall,
                      "meshes": [finite_element_torsion(area, wall) for area in (800, 200, 50)]},
                     indent=2, allow_nan=False))
