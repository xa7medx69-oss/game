# Design Decisions

## Opening and closing method

The selected method is four metal M2 heat-set inserts in the base, four reusable screws through the lid, a radial tongue-and-groove alignment joint, and a separate TPU compression gasket.

| Method considered | Decision |
|---|---|
| Heat-set inserts and screws | Selected: repeatable servicing, controlled clamp force, replaceable gasket, and visible failure inspection |
| Screws directly into printed plastic | Rejected for regular servicing because threads wear and strip |
| Permanent snap fits | Rejected for the alpha because printer variation, fatigue, and hidden cracking make repeated battery access unreliable |
| Quarter-turn bayonet lid | Rejected for the wearable alpha because accidental rotation and uneven gasket compression require more physical validation |
| Adhesive or ultrasonic welding | Rejected because the battery and prototype electronics must remain serviceable |
| Magnets | Rejected because they do not provide controlled sealing force and are undesirable for a wearable attachment system |

## Button retention

The cap is installed from below the lid and has a larger underside flange. It therefore cannot leave the enclosure during play. Removing the four lid screws provides deliberate service access. A separate TPU diaphragm centers the press, limits dust entry, and transfers force to the switch.

## Attachment

The wearable interface is a separate screw-mounted plate rather than a clip molded into the electronics shell. Strap and clip experiments can therefore change without reprinting the sealed enclosure. The first build uses a wide fabric strap with a breakaway overlap; it avoids neck mounting and sharp metal against the body.

## Electronics layout

The battery and XIAO board sit side-by-side on a removable carrier. This avoids stacking the PCB directly over the pouch cell, improves inspection, and gives wires defined routes. The USB connector points toward the side opening. The external antenna must be routed separately after worn-position RF testing.

## Drop and sealing strategy

The TPU lower bumper is sacrificial and replaceable. The perimeter gasket and button diaphragm reduce dust/splash paths, but they do not create a certified ingress rating. A production seal needs material, compression, sweat, temperature-cycle, and laboratory ingress validation.

## Production path

This geometry is optimized for FDM prototype learning. After electronics, human-factors, and play testing stabilize, redesign it for injection molding with appropriate draft, ribs, uniform wall thickness, molded gasket features, and formal DFM review. Do not send this FDM model directly to tooling.
