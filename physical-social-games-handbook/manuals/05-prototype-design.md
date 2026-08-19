> Physical Social Games Business | Manual 05 | Version 0.1 | 2026-08-19

# Prototype Design

## Prototype stages
**P0 - one-button proof:** one node, one hub, one browser. Demonstrate press, acknowledgement, UI state, and reconnect.

**P1 - six wired-on-bench nodes:** six powered nodes on a table. Run 1,000 scripted/manual presses, duplicates, simultaneous events, resets, and low-signal tests.

**P2 - wearable alpha:** battery power, 3D-printed housings, straps/clips, LED and vibration. Conduct controlled walking-speed games.

**P3 - playtest kit:** refined enclosures, labelled IDs, carry tray, charging checklist, host guide, and three complete game modes.

## Enclosure target
Prototype diameter 58-65 mm; depth 20-28 mm; mass below 70 g; 45-50 mm press area; minimum 2 mm rounded shell walls; no sharp external corners; recessed power switch; USB-C access that cannot be hit during play. Use a replaceable fabric/clip interface so the electronics do not need redesign when wear testing changes.

## Attachment
Start with a wide hook-and-loop armband/back strap plus a rounded clip adapter. Avoid strong magnets near medical devices and avoid lanyards or neck mounting. A production attachment should release under abnormal load and must not use a sharp metal edge against the body.

## Build order
1. Flash unique IDs and radio keys onto all boards.
2. Assemble hub and verify all six nodes on USB power.
3. Add one battery-powered node and measure current/temperature.
4. Clone the validated wiring to the other five.
5. Print one enclosure, conduct fit/press/antenna checks, then print the set.
6. Run bench reliability before any active play.
7. Run a safety briefing and low-speed playtest before chase variants.

## Exit criteria
Six nodes run 30 minutes without a reset, every node reconnects within five seconds after power cycling, the host sees battery/signal status, and two unfamiliar hosts can set up using the quick-start instructions.
