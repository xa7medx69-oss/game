"""Generate the Physical Social Games wearable pod enclosure.

The model uses millimetres. It intentionally keeps the electronics-specific
dimensions together near the top so the enclosure can be adjusted after the
real PCB, battery, switch, and printer are measured.
"""

from __future__ import annotations

import math
import json
from pathlib import Path

import numpy as np
import trimesh
from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parent
STL_DIR = ROOT / "stl"
PREVIEW_DIR = ROOT / "preview"

SEGMENTS = 128

# Primary envelope
OUTER_DIAMETER = 62.0
OUTER_RADIUS = OUTER_DIAMETER / 2
BASE_HEIGHT = 15.0
LID_HEIGHT = 4.0
WALL = 2.6
BOTTOM = 2.4
INNER_RADIUS = OUTER_RADIUS - WALL

# Button system
BUTTON_FACE_DIAMETER = 47.4
BUTTON_OPENING_DIAMETER = 48.4
BUTTON_TRAVEL = 0.8

# Closure and sealing
LID_SCREW_RADIUS = 27.0
LID_INSERT_HOLE_DIAMETER = 3.2
LID_CLEARANCE_HOLE_DIAMETER = 2.35
BASE_TONGUE_INNER_RADIUS = 26.9
BASE_TONGUE_OUTER_RADIUS = 28.0
BASE_TONGUE_HEIGHT = 1.2

# Measured component envelopes; edit these before a final print
XIAO_BOARD = (21.0, 17.8, 4.0)
BATTERY_MAX = (31.0, 21.5, 6.5)
USB_OPENING = (11.5, 6.5)
POWER_SWITCH_OPENING = (9.0, 4.5)


def moved(mesh: trimesh.Trimesh, xyz) -> trimesh.Trimesh:
    result = trimesh.Trimesh(
        vertices=np.array(mesh.vertices, copy=True),
        faces=np.array(mesh.faces, copy=True),
        process=False,
    )
    if hasattr(mesh.visual, "vertex_colors"):
        result.visual.vertex_colors = np.array(mesh.visual.vertex_colors, copy=True)
    result.apply_translation(xyz)
    return result


def cylinder(radius: float, height: float, z: float = 0.0) -> trimesh.Trimesh:
    mesh = trimesh.creation.cylinder(radius=radius, height=height, sections=SEGMENTS)
    mesh.apply_translation((0, 0, z + height / 2))
    return mesh


def box(size, center) -> trimesh.Trimesh:
    mesh = trimesh.creation.box(extents=size)
    mesh.apply_translation(center)
    return mesh


def union(*meshes: trimesh.Trimesh) -> trimesh.Trimesh:
    result = trimesh.boolean.union(list(meshes), engine="manifold", check_volume=True)
    return result


def difference(base: trimesh.Trimesh, *cuts: trimesh.Trimesh) -> trimesh.Trimesh:
    result = trimesh.boolean.difference([base, *cuts], engine="manifold", check_volume=True)
    return result


def annulus(outer_radius: float, inner_radius: float, height: float, z: float = 0.0):
    return difference(
        cylinder(outer_radius, height, z),
        cylinder(inner_radius, height + 0.4, z - 0.2),
    )


def rounded_plate(width: float, depth: float, height: float, radius: float):
    pieces = [
        box((width - 2 * radius, depth, height), (0, 0, height / 2)),
        box((width, depth - 2 * radius, height), (0, 0, height / 2)),
    ]
    for x in (-width / 2 + radius, width / 2 - radius):
        for y in (-depth / 2 + radius, depth / 2 - radius):
            pieces.append(moved(cylinder(radius, height), (x, y, 0)))
    return union(*pieces)


def lid_fastener_positions():
    for angle in (45, 135, 225, 315):
        radians = math.radians(angle)
        yield LID_SCREW_RADIUS * math.cos(radians), LID_SCREW_RADIUS * math.sin(radians)


def build_base_shell():
    shell = difference(
        cylinder(OUTER_RADIUS, BASE_HEIGHT),
        cylinder(INNER_RADIUS, BASE_HEIGHT - BOTTOM + 1.0, BOTTOM),
        # USB-C access on the +X side, aligned to the XIAO carrier.
        box((8.0, USB_OPENING[0], USB_OPENING[1]),
            (OUTER_RADIUS, 0, 10.6)),
        # Recessed power-switch opening on the front edge.
        box((POWER_SWITCH_OPENING[0], 8.0, POWER_SWITCH_OPENING[1]),
            (-11.0, -OUTER_RADIUS, 8.0)),
    )

    tongue = annulus(
        BASE_TONGUE_OUTER_RADIUS,
        BASE_TONGUE_INNER_RADIUS,
        BASE_TONGUE_HEIGHT,
        BASE_HEIGHT,
    )

    lid_bosses = []
    for x, y in lid_fastener_positions():
        lid_bosses.append(moved(cylinder(3.55, 7.0, 8.0), (x, y, 0)))

    # Blind mounting bosses for the removable strap plate.
    strap_bosses = [moved(cylinder(4.0, 5.5, BOTTOM), (x, 0, 0)) for x in (-17.0, 17.0)]

    # Two low bosses locate and retain the internal electronics carrier.
    carrier_bosses = [moved(cylinder(3.0, 3.2, BOTTOM), (0, y, 0)) for y in (-8.5, 8.5)]

    base = union(shell, tongue, *lid_bosses, *strap_bosses, *carrier_bosses)

    cuts = []
    # Heat-set insert holes open from the top.
    for x, y in lid_fastener_positions():
        cuts.append(moved(cylinder(LID_INSERT_HOLE_DIAMETER / 2, 5.2, 10.8), (x, y, 0)))

    # Blind heat-set insert holes open from the wearer-facing underside.
    for x in (-17.0, 17.0):
        cuts.append(moved(cylinder(3.7 / 2, 5.0, -0.1), (x, 0, 0)))

    # Carrier pilot holes are intentionally small for M2 thread-forming screws.
    for y in (-8.5, 8.5):
        cuts.append(moved(cylinder(1.45 / 2, 3.6, 3.0), (0, y, 0)))

    # Shallow compression groove for the separate TPU perimeter gasket.
    cuts.append(annulus(30.25, 29.0, 0.55, BASE_HEIGHT - 0.55))

    return difference(base, *cuts)


def build_lid():
    lid = difference(
        cylinder(OUTER_RADIUS, LID_HEIGHT),
        cylinder(BUTTON_OPENING_DIAMETER / 2, LID_HEIGHT + 1.0, -0.5),
        # Tongue-and-groove alignment feature on the underside.
        annulus(28.35, 26.55, 1.45, -0.1),
        # Transparent light-ring recess on the top.
        annulus(30.2, 29.15, 1.1, LID_HEIGHT - 1.1),
    )

    cuts = []
    for x, y in lid_fastener_positions():
        cuts.append(moved(cylinder(LID_CLEARANCE_HOLE_DIAMETER / 2, LID_HEIGHT + 1.0, -0.5), (x, y, 0)))
        cuts.append(moved(cylinder(2.05, 1.35, LID_HEIGHT - 1.35), (x, y, 0)))
    return difference(lid, *cuts)


def build_button_cap():
    # The underside flange retains the cap; it can only be removed after the lid is opened.
    flange = cylinder(24.8, 0.9)
    face = cylinder(BUTTON_FACE_DIAMETER / 2, 2.8, 0.9)
    tactile_target = cylinder(17.0, 0.45, 3.7)
    cap = union(flange, face, tactile_target)
    # Three shallow underside pockets reduce mass while leaving a broad load path.
    pockets = []
    for angle in (0, 120, 240):
        r = math.radians(angle)
        pockets.append(moved(cylinder(4.0, 0.45, -0.05), (11 * math.cos(r), 11 * math.sin(r), 0)))
    return difference(cap, *pockets)


def build_button_membrane():
    # Print flat with the plunger upward, then flip during assembly.
    diaphragm = cylinder(28.6, 0.8)
    plunger = cylinder(3.0, 2.6, 0.8)
    membrane = union(diaphragm, plunger)
    holes = []
    for x, y in lid_fastener_positions():
        holes.append(moved(cylinder(1.35, 4.0, -0.2), (x, y, 0)))
    return difference(membrane, *holes)


def build_perimeter_gasket():
    return annulus(30.18, 29.08, 0.8)


def build_light_ring():
    return annulus(30.08, 29.25, 1.0)


def build_bumper():
    # Closed TPU ring stretches over the lower shell and protects drop edges.
    return annulus(32.2, 30.72, 4.2)


def build_strap_mount():
    plate = rounded_plate(48.0, 34.0, 3.0, 6.0)
    cuts = [
        box((29.0, 5.0, 4.0), (0, -10.5, 1.5)),
        box((29.0, 5.0, 4.0), (0, 10.5, 1.5)),
    ]
    for x in (-17.0, 17.0):
        cuts.append(moved(cylinder(1.3, 4.0, -0.5), (x, 0, 0)))
        cuts.append(moved(cylinder(2.35, 1.2, -0.1), (x, 0, 0)))
    return difference(plate, *cuts)


def build_electronics_carrier():
    plate = rounded_plate(52.0, 24.0, 1.5, 6.0)

    # Battery pocket: 31 x 21.5 x 6.5 mm maximum, plus printable clearance.
    battery_center_x = -9.5
    walls = [
        box((1.5, 22.8, 3.2), (battery_center_x - 16.25, 0, 3.1)),
        box((1.5, 22.8, 3.2), (battery_center_x + 16.25, 0, 3.1)),
        box((31.0, 1.5, 3.2), (battery_center_x, -11.4, 3.1)),
        box((31.0, 1.5, 3.2), (battery_center_x, 11.4, 3.1)),
    ]

    # XIAO ESP32-C3 slides toward the USB opening on +X.
    board_center_x = 15.5
    board_width_y = XIAO_BOARD[1] + 0.6
    rails = [
        box((23.0, 1.4, 2.4), (board_center_x, -(board_width_y / 2 + 0.7), 2.7)),
        box((23.0, 1.4, 2.4), (board_center_x, board_width_y / 2 + 0.7, 2.7)),
        box((1.4, board_width_y + 2.8, 2.4), (board_center_x - 11.5, 0, 2.7)),
    ]

    carrier = union(plate, *walls, *rails)

    cuts = [
        # Slots for a removable fabric/TPU battery strap.
        box((3.0, 8.0, 3.0), (battery_center_x - 13.7, 0, 0.75)),
        box((3.0, 8.0, 3.0), (battery_center_x + 13.7, 0, 0.75)),
    ]
    for y in (-8.5, 8.5):
        cuts.append(moved(cylinder(1.15, 3.0, -0.5), (0, y, 0)))
    return difference(carrier, *cuts)


def build_tolerance_coupon():
    """Printer-specific test for screw clearances and heat-set insert holes."""
    coupon = rounded_plate(62.0, 20.0, 4.0, 3.0)
    diameters = (2.2, 2.3, 2.4, 3.0, 3.1, 3.2, 3.3, 3.4)
    cuts = []
    for index, diameter in enumerate(diameters):
        x = -24.5 + index * 7.0
        cuts.append(moved(cylinder(diameter / 2, 5.0, -0.5), (x, 0, 0)))
    return difference(coupon, *cuts)


def validate(name: str, mesh: trimesh.Trimesh):
    if not mesh.is_watertight:
        raise ValueError(f"{name}: mesh is not watertight")
    if not mesh.is_winding_consistent:
        raise ValueError(f"{name}: inconsistent face winding")
    if mesh.volume <= 0:
        raise ValueError(f"{name}: non-positive volume")
    bounds = mesh.bounds
    return {
        "name": name,
        "faces": len(mesh.faces),
        "volume_mm3": round(float(mesh.volume), 1),
        "size_mm": [round(float(v), 2) for v in bounds[1] - bounds[0]],
    }


def validate_design_clearances():
    checks = {
        "cap_radial_clearance_each_side": BUTTON_OPENING_DIAMETER / 2 - BUTTON_FACE_DIAMETER / 2,
        "cap_retention_overlap": 24.8 - BUTTON_OPENING_DIAMETER / 2,
        "screw_head_to_button_opening": LID_SCREW_RADIUS - 2.05 - BUTTON_OPENING_DIAMETER / 2,
        "screw_head_to_outer_edge": OUTER_RADIUS - LID_SCREW_RADIUS - 2.05,
        "tongue_inner_clearance": BASE_TONGUE_INNER_RADIUS - 26.55,
        "tongue_outer_clearance": 28.35 - BASE_TONGUE_OUTER_RADIUS,
        "membrane_material_outside_screw_hole": 28.6 - LID_SCREW_RADIUS - 1.35,
    }
    minimums = {
        "cap_radial_clearance_each_side": 0.35,
        "cap_retention_overlap": 0.5,
        "screw_head_to_button_opening": 0.6,
        "screw_head_to_outer_edge": 1.5,
        "tongue_inner_clearance": 0.25,
        "tongue_outer_clearance": 0.25,
        "membrane_material_outside_screw_hole": 0.2,
    }
    for name, value in checks.items():
        if value < minimums[name]:
            raise ValueError(f"Clearance check failed: {name}={value:.3f} mm")
    return {name: round(value, 3) for name, value in checks.items()}


def set_color(mesh: trimesh.Trimesh, rgba):
    colored = trimesh.Trimesh(
        vertices=np.array(mesh.vertices, copy=True),
        faces=np.array(mesh.faces, copy=True),
        process=False,
    )
    colored.visual.vertex_colors = np.tile(np.asarray(rgba, dtype=np.uint8), (len(colored.vertices), 1))
    return colored


def save_preview(scene: trimesh.Scene, output: Path):
    """Small dependency-free orthographic renderer for the exploded preview."""
    width, height = 1500, 1200
    image = Image.new("RGB", (width, height), "#eef3f7")
    draw = ImageDraw.Draw(image)

    azimuth = math.radians(38)
    elevation = math.radians(57)
    rz = np.array([
        [math.cos(azimuth), -math.sin(azimuth), 0],
        [math.sin(azimuth), math.cos(azimuth), 0],
        [0, 0, 1],
    ])
    rx = np.array([
        [1, 0, 0],
        [0, math.cos(elevation), -math.sin(elevation)],
        [0, math.sin(elevation), math.cos(elevation)],
    ])
    rotation = rx @ rz

    triangles = []
    all_points = []
    for geometry in scene.geometry.values():
        vertices = geometry.vertices @ rotation.T
        all_points.append(vertices[:, :2])
        base = np.array(geometry.visual.vertex_colors[0][:3], dtype=float)
        for face in geometry.faces:
            tri = vertices[face]
            normal = np.cross(tri[1] - tri[0], tri[2] - tri[0])
            norm = np.linalg.norm(normal)
            if norm == 0:
                continue
            normal /= norm
            light = max(0.25, min(1.0, 0.45 + 0.55 * np.dot(normal, np.array([-0.3, -0.4, 0.86]))))
            color = tuple(np.clip(base * light, 0, 255).astype(int))
            triangles.append((float(tri[:, 2].mean()), tri[:, :2], color))

    points = np.vstack(all_points)
    minimum, maximum = points.min(axis=0), points.max(axis=0)
    scale = min((width - 180) / (maximum[0] - minimum[0]), (height - 160) / (maximum[1] - minimum[1]))
    center = (minimum + maximum) / 2

    for _, tri, color in sorted(triangles, key=lambda item: item[0]):
        screen = (tri - center) * scale
        screen[:, 0] += width / 2
        screen[:, 1] = height / 2 - screen[:, 1]
        draw.polygon([tuple(point) for point in screen], fill=color, outline="#25384a")

    draw.text((50, 40), "WEARABLE POD V1 — EXPLODED PRINTED PARTS", fill="#102a43")
    draw.text((50, height - 55), "62 mm body • screw-serviceable • TPU seals • replaceable strap mount", fill="#41566b")
    image.save(output)


def main():
    STL_DIR.mkdir(parents=True, exist_ok=True)
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)

    clearance_report = validate_design_clearances()

    parts = {
        "01-base-shell": build_base_shell(),
        "02-top-lid": build_lid(),
        "03-button-cap": build_button_cap(),
        "04-button-membrane-tpu": build_button_membrane(),
        "05-perimeter-gasket-tpu": build_perimeter_gasket(),
        "06-light-ring-transparent": build_light_ring(),
        "07-drop-bumper-tpu": build_bumper(),
        "08-strap-mount": build_strap_mount(),
        "09-electronics-carrier": build_electronics_carrier(),
        "10-tolerance-coupon": build_tolerance_coupon(),
    }

    report = []
    for name, mesh in parts.items():
        report.append(validate(name, mesh))
        output_path = STL_DIR / f"{name}.stl"
        mesh.export(output_path)
        # Reload the exported STL so validation covers the saved deliverable,
        # not only the in-memory CSG result.
        exported = trimesh.load_mesh(output_path, process=False)
        exported.merge_vertices()
        validate(f"{name}-export", exported)

    colors = {
        "01-base-shell": (20, 49, 73, 255),
        "02-top-lid": (24, 63, 91, 255),
        "03-button-cap": (24, 145, 153, 255),
        "04-button-membrane-tpu": (236, 114, 82, 255),
        "05-perimeter-gasket-tpu": (50, 55, 60, 255),
        "06-light-ring-transparent": (255, 189, 69, 180),
        "07-drop-bumper-tpu": (42, 48, 54, 255),
        "08-strap-mount": (24, 63, 91, 255),
        "09-electronics-carrier": (236, 236, 226, 255),
    }
    exploded_z = {
        "08-strap-mount": -8,
        "07-drop-bumper-tpu": 0,
        "01-base-shell": 7,
        "09-electronics-carrier": 27,
        "05-perimeter-gasket-tpu": 38,
        "04-button-membrane-tpu": 45,
        "02-top-lid": 53,
        "06-light-ring-transparent": 61,
        "03-button-cap": 68,
    }
    scene = trimesh.Scene()
    for name, mesh in parts.items():
        if name not in exploded_z:
            continue
        geometry = moved(set_color(mesh, colors[name]), (0, 0, exploded_z[name]))
        scene.add_geometry(geometry, node_name=name, geom_name=name)

    scene.export(PREVIEW_DIR / "wearable-pod-v1-exploded.glb")
    save_preview(scene, PREVIEW_DIR / "wearable-pod-v1-exploded.png")

    validation_output = {
        "units": "millimetres",
        "design_clearances_mm": clearance_report,
        "parts": report,
    }
    (ROOT / "model-validation.json").write_text(
        json.dumps(validation_output, indent=2),
        encoding="utf-8",
    )

    print("Generated and validated wearable-pod-v1")
    print({"design_clearances_mm": clearance_report})
    for item in report:
        print(item)


if __name__ == "__main__":
    main()
