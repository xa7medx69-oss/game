> Physical Social Games Business | Manual 03 | Version 0.1 | 2026-08-19

# Circuit Design

## Prototype node wiring

| Function | XIAO ESP32-C3 pin | Circuit instruction |
|---|---|---|
| Main press | D1 / GPIO3 | Switch to GND; enable internal pull-up; add 100 nF only if software debounce is insufficient. |
| RGB LED | D2 / GPIO4 | WS2812 data through 330 ohm series resistor; 100 uF across LED supply; limit brightness in firmware. |
| Buzzer | D3 / GPIO5 | Drive piezo or magnetic buzzer through an NPN transistor and base resistor; add flyback protection if inductive. |
| Vibration | D4 / GPIO6 | Logic-level N-MOSFET low-side drive; flyback diode across coin motor; 100 nF motor suppression. |
| Battery sense | D0 / GPIO2 | 220 k/100 k divider to ADC; enable through a transistor if idle leakage matters. |
| Power | BAT and GND | Protected 3.7 V LiPo to battery pads; physical slide switch in series for a true off state. |

## Electrical rules
Keep the 2.4 GHz antenna edge free of copper, battery, motor, steel clip, and fingers. Place motor and buzzer away from the antenna. Add test pads for 3V3, GND, button, LED data, UART TX/RX, and battery. Use keyed battery connectors and strain relief. Do not charge unattended during early prototypes.

## Switch and press design
The large visible cap should transfer force to a small rated tactile switch without side-loading it. Start with 1.5-2.5 N actuation and 0.5-1.0 mm perceived travel. The enclosure must prevent clothing pressure from maintaining the switch. A raised protective rim helps but must not make intentional presses difficult.

## Protection for a production PCB
Include reverse-polarity protection, LiPo over-charge/over-discharge/short protection, USB ESD protection, motor transient suppression, programming pads, a reset method, and a current-measurement link. Use only a certified battery pack and document cell traceability.

## Verification instruction
Before connecting a battery, perform continuity and resistance checks. Power from a current-limited bench supply at 3.8 V with a 150 mA limit, confirm quiescent current, then test each output separately. Record peak current during LED white, motor start, radio transmit, and combined worst case.

## Production handoff outputs
Schematic PDF, ERC-clean source, PCB layout, Gerbers, drill files, pick-and-place, BOM with manufacturer part numbers and alternates, test-point map, programming fixture drawing, and a written bring-up procedure.
