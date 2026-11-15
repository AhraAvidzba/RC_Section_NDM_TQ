"""Prepare the §114 six-force history and check independent VBA section resultants."""

import argparse
import csv
import hashlib
import json
import math
from pathlib import Path


ORDER = ("N", "Mx", "My", "T", "Qx", "Qy")
KINEMATICS = ("epsilon0", "kappaX_per_mm", "kappaY_per_mm", "twist_per_mm", "mean_shear_x", "mean_shear_y")
AXIAL_LOCATIONS = {(-50.0, -40.0), (50.0, -40.0), (50.0, 40.0), (-50.0, 40.0)}
AREA, LEVER, STAGES = 32000.0, 200.0, 48
TARGET_TOL, INTEGRATION_TOL = 5e-6, 1e-9  # MPa, referred to the same gross area and geometric lever.
CONCRETE = (30000.0, 0.2, 18.0, 30.0, 3.0, 0.002, 0.1, 10.0, 100.0)
STEEL = {"L500": (12.0, 200000.0, 500.0, 2000.0, 0.05), "H240": (8.0, 200000.0, 240.0, 2000.0, 0.05)}


def require(condition, message):
    if not condition:
        raise ValueError(message)


def finite(values):
    numbers = tuple(map(float, values))
    require(all(math.isfinite(value) for value in numbers), "All reported numbers must be finite.")
    return numbers


def wrench_error(first, second):
    require(len(first) == len(second) == 6, "A wrench must have six components.")
    return max(abs(a - b) / AREA / (LEVER if 1 <= i <= 3 else 1.0)
               for i, (a, b) in enumerate(zip(first, second, strict=True)))


def read_reference(path):
    encoded = path.read_bytes()
    record = json.loads(encoded)
    require(tuple(record["resultants_order"]) == ORDER, "OpenSees wrench order is inconsistent.")
    require((record["C_E_MPa"], record["C_nu"]) == CONCRETE[:2], "Unexpected concrete elastic parameters.")
    require(record["concrete_table_sha256"].lower() ==
            "1dc88ceedf59a1e9c7d973554f7698f57ed243a7511e83569c6c232ed5d85880", "Unexpected §97 envelope.")
    geometry = record["section_geometry"]
    require([geometry[key] for key in ("width_x_mm", "depth_y_mm", "prism_length_z_mm", "section_z_mm", "hoop_spacing_mm")]
            == [200, 160, 800, 400, 100] and geometry["mesh"] in ([4, 4, 16], [8, 8, 16], [12, 12, 16])
            and geometry["concrete_origin_xy_mm"] == [0, 0], "Unexpected prism geometry, mesh or comparison point.")
    require(geometry["long_bar_count"] == 4 and
            {tuple(point) for point in geometry["long_bars_xy_mm"]} == AXIAL_LOCATIONS and
            {tuple(point) for point in geometry["hoop_closed_centerline_vertices_mm"]} == AXIAL_LOCATIONS and
            geometry["hoop_levels_from_clamped_end_mm"] == list(range(50, 800, 100)), "Unexpected reinforcement geometry.")
    for material, grade in (("long_steel", "L500"), ("hoop_steel", "H240")):
        entry = record[material]
        diameter, modulus, yield_stress, slope, _ = STEEL[grade]
        require(entry["name"] == grade and (entry["diameter_mm"], entry["E_MPa"], entry["fy_MPa"]) ==
                (diameter, modulus, yield_stress) and entry["post_yield_ratio"] * modulus == slope
                and math.isclose(entry["area_mm2"], math.pi * diameter**2 / 4.0, rel_tol=1e-12), f"Unexpected {grade} law.")
    wanted = record["load_history"]
    require(len(wanted) == STAGES, "OpenSees history must contain exactly 48 accepted stages.")
    external_error = 0.0
    for i, row in enumerate(wanted, 1):
        require(math.isclose(row["factor"], i / STAGES, rel_tol=0.0, abs_tol=1e-13), "Unexpected stage factor.")
        target = finite(row["resultants"])
        below = tuple(a + b for a, b in zip(row["concrete_cut_virtual_work"], row["longitudinal_steel_cut"], strict=True))
        above_data = row["from_above"]
        above = tuple(a + b for a, b in zip(above_data["concrete_cut_virtual_work"], above_data["longitudinal_steel_cut"], strict=True))
        external_error = max(external_error, wrench_error(target, finite(row["resultant_from_end_reactions"])),
                             wrench_error(target, finite(below)), wrench_error(target, finite(above)))
        finite(row["kinematics"][key] for key in KINEMATICS)
    require(external_error <= 5e-7, f"OpenSees independent cut reconstructions differ by {external_error:g} MPa.")
    return record, hashlib.sha256(encoded).hexdigest().upper(), external_error


def prepare(path, target):
    reference, digest, error = read_reference(path)
    with target.open("x", encoding="ascii", newline="") as stream:
        writer = csv.writer(stream, delimiter=";", lineterminator="\r\n")
        writer.writerow(("NDM_PRISM_INPUT_1", digest))
        writer.writerow(("step", "factor", *ORDER))
        for stage, row in enumerate(reference["load_history"], 1):
            writer.writerow((stage, repr(row["factor"]), *(repr(value) for value in row["resultants"])))
    return {"stages": STAGES, "max_open_sees_wrench_error_mpa": error, "reference_sha256": digest, "file": str(target)}


def add_unique(target, key, data):
    require(key not in target, f"Duplicate record {key}.")
    target[key] = data


def read_native(path):
    targets, steps, points, stops, steel, routes = {}, {}, {}, {}, {}, {}
    version = meta = concrete = finish = None
    with path.open(encoding="utf-16", newline="") as stream:
        for row in csv.reader(stream, delimiter=";"):
            if row[0] == "META" and len(row) == 9 and meta is None:
                require(row[1] in ("NDM_PRISM_VBA_1", "NDM_PRISM_VBA_2"), "Unexpected VBA export version.")
                version = row[1]
                meta = row[2:]
            elif row[0] == "CONCRETE" and len(row) == 10 and concrete is None:
                concrete = finite(row[1:])
            elif row[0] == "STEEL" and len(row) == 8:
                add_unique(steel, row[1], finite(row[2:7]))
            elif row[0] == "TARGET" and len(row) == 9:
                add_unique(targets, int(row[1]), finite(row[2:]))
            elif row[0] == "ROUTE" and len(row) == 6 and version == "NDM_PRISM_VBA_2":
                parts, failed = map(int, row[2:4])
                fraction, minimum = finite(row[4:])
                require(0 <= parts <= 16 and 0 <= failed <= 5 and parts + failed > 0 and (parts == 0) == (fraction == 0)
                        and fraction in tuple(i / 16 for i in range(17)) and minimum in (1, 0.5, 0.25, 0.125, 0.0625),
                        "Invalid substep history.")
                add_unique(routes, int(row[1]), {"accepted": parts, "rejected": failed,
                                                 "fraction_reached": fraction, "minimum_attempted_fraction": minimum})
            elif row[0] == "STEP" and len(row) == 17:
                add_unique(steps, int(row[1]), finite(row[2:]))
            elif row[0] == "POINT" and len(row) == 30:
                add_unique(points.setdefault(int(row[1]), {}), int(row[2]), (row[3], finite(row[4:])))
            elif row[0] == "STOP" and len(row) == 6:
                add_unique(stops, int(row[1]), dict(internal_status=row[2], result_code=row[3], reason=row[4], diagnostics=row[5]))
            elif row[0] == "END" and len(row) == 4 and finish is None:
                finish = row[1], int(row[2]), int(row[3])
            else:
                raise ValueError(f"Unexpected or truncated VBA record {row!r}.")
    require(concrete == CONCRETE and steel == STEEL and meta is not None and finish is not None,
            "Missing report completion or different material input.")
    status, accepted, requested = finish
    require(requested == STAGES and 0 <= accepted <= requested, "Incorrect accepted stage count.")
    expected_steps = set(range(1, accepted + 1))
    require(set(steps) == set(points) == expected_steps, "Results do not form a complete accepted prefix.")
    if accepted == requested:
        require(status == "ReachedAll48" and not stops and set(targets) == expected_steps, "Invalid completed run.")
    else:
        require(status == "StoppedUnbalanced" and set(stops) == {accepted + 1} and set(targets) == expected_steps | {accepted + 1},
                "Unverified data appears after the stopped stage.")
    if version == "NDM_PRISM_VBA_2":
        require(set(routes) == set(targets), "Missing substep history.")
        require(all((route["fraction_reached"] == 1) == (stage in expected_steps) for stage, route in routes.items()),
                "The final fraction does not match the accepted stage.")
    return meta, accepted, targets, steps, points, stops, routes


def read_open_sees_steel(path, reference):
    require(hashlib.sha256(path.read_bytes()).hexdigest() == reference["reference_sha256"],
            "The raw OpenSees output does not match the verified reference export.")
    raw = json.loads(path.read_bytes())["nonlinear"]["compression_end_displacement"]["final_extraction"]
    tags, history = raw["steel_geometry"], raw["steel_history"]
    require(len(history) == STAGES, "Missing OpenSees steel history.")
    result = []
    for entry, stage in zip(history, reference["load_history"], strict=True):
        require(math.isclose(entry["factor"], stage["factor"], rel_tol=0.0, abs_tol=1e-13), "Wrong OpenSees steel stage.")
        long_at_cut, hoop_near_cut = [], {350: [], 450: []}
        for tag, stress, strain in entry["truss"]:
            geometry = tags[str(tag)]
            (x, y, z1), (_, _, z2) = geometry["nodes_xyz_mm"]
            if geometry["class"] == "L500" and (x, y) in AXIAL_LOCATIONS and 400 in (z1, z2):
                long_at_cut.append({"xy_mm": (x, y), "z_endpoints_mm": (z1, z2), "stress_mpa": stress, "strain": strain})
            elif geometry["class"] == "H240" and z1 == z2 and z1 in hoop_near_cut:
                hoop_near_cut[z1].append(stress)
        require(len(long_at_cut) == 8 and all(hoop_near_cut.values()), "Incomplete location-labelled steel samples.")
        result.append((long_at_cut, {z: (min(rows), max(rows)) for z, rows in hoop_near_cut.items()}))
    return result


def recover_wrench(points, point_count, spans):
    require(set(points) == set(range(1, point_count + 1)), "Incomplete or repeated quadrature point list.")
    terms = [[] for _ in ORDER]
    areas = {"C": [], "L500": [], "H240": []}
    longitudinal, hoops = [], []
    damage = [0.0, 0.0]
    for grade, values in points.values():
        x, y, area = values[:3]
        strain, stress, direction = values[3:9], values[9:15], values[15:21]
        steel_strain, steel_stress, fy, tension, compression = values[21:]
        require(grade in areas and area > 0.0 and abs(x) <= 100.0 + 1e-9 and abs(y) <= 80.0 + 1e-9,
                "Invalid section quadrature sample.")
        areas[grade].append(area)
        if grade == "C":
            require(not any(direction) and not any((steel_strain, steel_stress, fy)), "Concrete point contains steel data.")
            normal, tx, ty = stress[2], stress[4], stress[5]
            damage = [max(a, b) for a, b in zip(damage, (tension, compression), strict=True)]
        else:
            require(fy == STEEL[grade][2] and
                    abs(steel_strain - math.fsum(a * b for a, b in zip(strain, direction, strict=True))) < 1e-12,
                    "Steel grade or projected strain mismatch.")
            normal, tx, ty = (steel_stress * direction[i] - stress[i] for i in (2, 4, 5))
            plastic = steel_strain - steel_stress / STEEL[grade][1]
            sample = {"xy_mm": (x, y), "direction": direction, "stress_mpa": steel_stress, "strain": steel_strain,
                      "stress_over_yield": abs(steel_stress) / fy, "plastic_strain_inferred": plastic,
                      "strain_state": "PlasticHistory" if abs(plastic) > 1e-12 else "ElasticStrain"}
            if grade == "L500":
                require(direction == (0, 0, 1, 0, 0, 0), "Unexpected longitudinal direction.")
                longitudinal.append(sample)
            else:
                require(not any(direction[i] for i in (2, 4, 5)) and
                        ((abs(abs(y) - 40.0) < 1e-8 and direction[0] == 1) or
                         (abs(abs(x) - 50.0) < 1e-8 and direction[1] == 1)), "Unexpected hoop axis or direction.")
                hoops.append(sample)
        components = (normal, y * normal, x * normal, x * ty - y * tx, tx, ty)
        for term, component in zip(terms, components, strict=True):
            term.append(area * component)
    require(len(areas["C"]) == (4 * spans) ** 2 and len(longitudinal) == 4
            and len(hoops) > 0 and {tuple(row["xy_mm"]) for row in longitudinal} == AXIAL_LOCATIONS,
            "Unexpected concrete or steel support.")
    for key, expected in zip(("C", "L500", "H240"), (AREA, 4 * math.pi * 36, 360 / 100 * math.pi * 16), strict=True):
        require(math.isclose(math.fsum(areas[key]), expected, rel_tol=1e-10), f"Incorrect {key} quadrature area.")
    return tuple(math.fsum(term) for term in terms), longitudinal, hoops, damage


def compare(reference_path, raw_path, native_path):
    source, signature, external_error = read_reference(reference_path)
    meta, accepted, targets, stages, points, stops, routes = read_native(native_path)
    require(meta[0] == signature and int(meta[2]) == 4 and finite(meta[4:6]) == (0, 0),
            "Wrong reference, quadrature support or source origin in VBA output.")
    spans, point_count = int(meta[1]), int(meta[3])
    require(1 <= spans <= 3 and finite(meta[6:])[-1] > 0.0, "Invalid section basis or length scale.")
    steel_reference = read_open_sees_steel(raw_path, source)
    comparisons = []
    target_error = recovered_error = 0.0
    for i, actual_target in sorted(targets.items()):
        expected = source["load_history"][i - 1]
        demand = expected["resultants"]
        require(all(math.isclose(a, b, rel_tol=1e-13, abs_tol=1e-11)
                    for a, b in zip(actual_target, (expected["factor"], *demand), strict=True)), f"Stage {i}: target input differs.")
        if i not in stages:
            continue
        numeric = stages[i]
        reported = numeric[1:7]
        reconstructed, longitudinal, hoops, energies = recover_wrench(points[i], point_count, spans)
        require(abs(numeric[0] - actual_target[0]) <= 1e-13, f"Stage {i}: factor differs.")
        stage_target_error, stage_recovered_error = wrench_error(demand, reconstructed), wrench_error(reported, reconstructed)
        target_error = max(target_error, stage_target_error)
        recovered_error = max(recovered_error, stage_recovered_error)
        require(stage_target_error <= TARGET_TOL and stage_recovered_error <= INTEGRATION_TOL,
                f"Stage {i}: force balance failed, target {stage_target_error:g}, recovered {stage_recovered_error:g} MPa.")
        long_reference, adjacent_hoops = steel_reference[i - 1]
        comparison = {"stage": i, "factor": numeric[0], "target": demand, "vba_integrated": reconstructed,
                      "vba_solver_reported": reported, "target_error_mpa": stage_target_error,
                      "projection_error_mpa": stage_recovered_error,
                      "kinematics": {key: {"vba_local_section": val, "open_sees_gauge_100mm": expected["kinematics"][key],
                                           "difference": val - expected["kinematics"][key]}
                                     for key, val in zip(KINEMATICS, numeric[7:13], strict=True)},
                      "steel_utilization_against_diagnostic_strain_limit": numeric[13:],
                      "concrete_max_dissipation_energy_vba_mpa": energies,
                      "longitudinal_steel_vba_at_cut": longitudinal, "longitudinal_steel_open_sees_adjacent_trusses": long_reference,
                      "hoop_pointwise_comparison_status": "NotComparableNoDiscreteHoopAtCutZ400",
                      "hoop_vba_smeared_at_cut": hoops, "hoop_open_sees_range_at_adjacent_z_levels_mpa": adjacent_hoops}
        comparisons.append(comparison)
    status = ("EquilibratedCommonHistoryDifferentModels" if accepted == STAGES else
              "PartialHistoryNotComparableBeyondAcceptedPrefix" if accepted else "NoCommonEquilibriumComparisonUnavailable")
    return {"status": status, "received_trials": len(targets), "confirmed_steps": accepted, "requested_steps": STAGES,
            "reference_sha256": signature, "basis_spans_per_axis": spans, "substeps": routes,
            "open_sees_mesh": source["section_geometry"]["mesh"],
            "vba_material": {"concrete": CONCRETE, "steel": STEEL, "spatial_material_histories": 3},
            "max_target_balance_error_vba_mpa": target_error if accepted else None,
            "max_independent_vba_projection_error_mpa": recovered_error if accepted else None,
            "max_open_sees_wrench_error_mpa": external_error, "stopping": stops, "cases": comparisons,
            "scope": ["The independent VBA integration checks the common six-force path, not constitutive identity or capacity.",
                      "OpenSees uses a volume model with gross concrete and discrete bonded steel. VBA replaces concrete by steel and smears the hoop.",
                      "OpenSees uses uniaxial tables obtained from NDM with the same calibration parameters. Multiaxial laws and measured kinematics differ.",
                      "VBA provides a local section field while the OpenSees projection uses a 100 mm gauge.",
                      "There is no discrete hoop at z=400 mm. Only location-labelled ranges at z=350 and 450 mm are reported separately."]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="operation", required=True)
    for name in ("prepare", "compare"):
        command = commands.add_parser(name)
        command.add_argument("--reference", type=Path, required=True)
        command.add_argument("--out", type=Path, required=True)
        if name == "compare":
            command.add_argument("--raw-open-sees", type=Path, required=True)
            command.add_argument("--native", type=Path, required=True)
    args = parser.parse_args()
    if args.operation == "prepare":
        summary = prepare(args.reference, args.out)
    else:
        result = compare(args.reference, args.raw_open_sees, args.native)
        with args.out.open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2, ensure_ascii=False, allow_nan=False)
            stream.write("\n")
        summary = {key: result[key] for key in ("status", "confirmed_steps", "requested_steps",
                                               "max_target_balance_error_vba_mpa", "max_independent_vba_projection_error_mpa")}
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
