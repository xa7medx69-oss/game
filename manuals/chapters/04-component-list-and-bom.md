> Physical Social Games Business | Manual 04 | Version 0.1 | 2026-08-19

# Component List and BOM

## Buy list for one seven-board MVP

| Item | Qty | Planning cost (AED) | Notes |
|---|---:|---:|---|
| Seeed XIAO ESP32-C3 | 7 | 210-315 | Six nodes plus one hub; buy one spare if budget allows. |
| Protected 500 mAh LiPo | 6 | 90-150 | Use reputable packs with correct polarity and connectors. |
| Large momentary switches/caps | 8 | 35-80 | Six plus spares; prototype arcade or low-profile caps. |
| RGB LEDs/rings | 8 | 40-90 | One per node plus spares. |
| Coin vibration motors | 8 | 40-80 | Optional in prototype round one but recommended. |
| Buzzers, MOSFETs, diodes, passives | 1 lot | 35-70 | Breadboard/protoboard stage. |
| Power switches and wire/connectors | 1 lot | 30-60 | Add strain relief and insulation. |
| USB-C charging cable set/charger | 1 | 45-90 | Multi-port quality charger; never use damaged cells. |
| Prototype enclosures/straps/clips | 6 | 90-180 | 3D printing, hook-and-loop, fabric straps. |
| Hub case and optional OLED | 1 | 35-75 | OLED is diagnostic, not required. |

**Expected prototype total:** approximately AED 615-1,190 before tools, shipping, and failed parts. This is a planning range, not a supplier quote.

## Required tools
Temperature-controlled soldering iron, lead-free or appropriately handled solder, flux, flush cutters, wire stripper, heat-shrink, multimeter, USB power meter, bench supply with current limit, small scale/calipers, and access to a 3D printer. A basic logic analyser is useful for debugging.

## Supplier discipline
Record supplier, order date, exact MPN, lot/batch, price, lead time, datasheet URL, and substitute. Do not silently swap batteries, radios, switches, or plastics. For each critical part maintain one approved alternate and document the validation needed after substitution.

## Cost-down ladder
Prototype dev boards first; then one custom PCB; then panelized PCBs and a simple machined/printed enclosure; only after demand validation consider injection moulding and a charging case. Do not spend on tooling before event reliability and repeat-play are proven.
