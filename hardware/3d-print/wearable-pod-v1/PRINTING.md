# Printing Instructions

## Print one tolerance coupon first

Print `10-tolerance-coupon.stl` using the same material, nozzle, layer height, and wall settings planned for the housing.

The eight holes from left to right are:

`2.2, 2.3, 2.4, 3.0, 3.1, 3.2, 3.3, 3.4 mm`

Use the first three to select an M2 clearance hole. Use the last five to select the heat-set-insert pilot hole. Change the constants near the top of `generate_models.py` if your printer needs compensation.

## Materials

| Part | Recommended prototype material | Reason |
|---|---|---|
| Base, lid, cap, strap mount, carrier | PETG | Tough, printable, and suitable for functional mechanical parts |
| Button membrane, gasket, bumper | TPU 95A | Flexible, replaceable, and impact absorbing |
| Light ring | Natural/translucent PETG | Diffuses LEDs better than an opaque shell |

PETG is appropriate for prototypes, but do not leave the pod in a closed vehicle in UAE heat. For higher-temperature testing, evaluate ASA or a suitable PC blend using an enclosed printer. Production material selection requires impact, chemical, sweat, UV, flammability, and regulatory review.

## Baseline slicer settings

- 0.4 mm nozzle.
- 0.20 mm layer height; 0.16 mm for the membrane and gasket if available.
- Four perimeters for rigid structural parts.
- Five top and bottom layers.
- 30–40% gyroid or adaptive-cubic infill.
- 100% infill for the gasket and membrane.
- Elephant-foot compensation enabled after measuring the coupon.
- Seam positioned away from USB, switch, gasket, and button-guide surfaces.
- No automatic scaling.

## Orientation

| STL | Build-plate orientation |
|---|---|
| `01-base-shell` | Flat wearer-facing bottom on the plate; cavity upward |
| `02-top-lid` | Flat underside on the plate |
| `03-button-cap` | Retaining flange on the plate; visible face upward |
| `04-button-membrane-tpu` | Flat diaphragm on the plate; plunger upward; flip during assembly |
| `05-perimeter-gasket-tpu` | Flat |
| `06-light-ring-transparent` | Flat |
| `07-drop-bumper-tpu` | Flat |
| `08-strap-mount` | Flat wearer-facing side on the plate |
| `09-electronics-carrier` | Large base plate on the plate |
| `10-tolerance-coupon` | Flat |

The parts are designed to avoid mandatory support. If a printer struggles with the carrier rails, use support from the build plate only and keep support away from board-contact surfaces.

## Acceptance inspection

Reject and reprint a part if it has layer separation, a warped gasket surface, a cracked boss, sharp strings inside the battery compartment, blocked USB/switch access, a cap that rubs or remains depressed, or strap-slot damage.

PETG guidance: [Prusa PETG material documentation](https://help.prusa3d.com/article/petg_2059).
