> Physical Social Games Business | Manual 02 | Version 0.1 | 2026-08-19

# Hardware Architecture

## Recommended MVP architecture
Use **six Seeed XIAO ESP32-C3 wearable nodes plus one XIAO ESP32-C3 hub**. Nodes send compact events to the hub over ESP-NOW. The hub runs a local Wi-Fi access point and WebSocket server; a phone or laptop joins the hub network and opens the local web interface. Internet is optional.

```text
[6 wearable nodes] -- ESP-NOW --> [hub] -- local Wi-Fi/WebSocket --> [phone/laptop browser]
```

This is better for the proof-of-concept than connecting six BLE peripherals directly to a phone: the host pairs once with the hub, the game works offline, event timing is centralized, and browser/device Bluetooth differences are removed. ESP32-C3 supports 2.4 GHz Wi-Fi and Bluetooth LE 5, so the same dev board preserves experimentation options. Official overview: https://docs.espressif.com/projects/esp-hardware-design-guidelines/en/latest/esp32c3/product-overview.html

## Node blocks
- XIAO ESP32-C3 module with antenna kept clear of battery and metal clip.
- Large momentary switch under a soft press cap.
- Addressable RGB LED or a 6-8 LED ring for player/status feedback.
- Coin vibration motor driven through a MOSFET.
- Optional piezo buzzer; keep node audio subtle and let the host produce most sound.
- 500 mAh protected LiPo, slide power switch, and USB-C charging through the board.
- Prototype enclosure with breakaway fabric strap or rounded clip.

## Hub blocks
- XIAO ESP32-C3 connected to USB power.
- ESP-NOW receiver and event authority.
- Local Wi-Fi access point, HTTP server, and WebSocket endpoint.
- USB serial debug channel; optional OLED only for diagnostics.

## Protocol
Each node sends `device_id`, `event_type`, `sequence`, `node_time_ms`, battery percentage, and a message authentication/check value. The hub deduplicates by device plus sequence, timestamps receipt, applies the game rule, and replies with an acknowledgement containing the accepted state. Retry at 40, 100, and 220 ms with random jitter, then show a local error pattern.

## Range target
Design for 25 m reliable line-of-sight and 10-15 m through typical villa partitions. Do not market a range until measured. Test the exact antenna orientation with the pod worn against different fabrics and body positions because the human body attenuates 2.4 GHz.

## Consumer architecture after validation
Re-evaluate a Nordic nRF52-class design for lower power and a robust proprietary 2.4 GHz/BLE hub. The nRF52840 supports Bluetooth LE, Bluetooth Mesh, NFC, Thread, and Zigbee, but it raises firmware/tooling and unit-cost complexity. Official product page: https://www.nordicsemi.com/products/nrf52840

## Battery and thermal targets
Active play: 8 hours minimum; shelf-off leakage: under 5% per month; charge time under 2.5 hours; protected cell; enclosure surface below 42 C in normal use. For the MVP, measure actual current in idle, active radio, LED animation, vibration, and power-off states before claiming runtime.

## Reliability rules
- The hub, not the node, owns the official game state.
- Never award points from an unacknowledged local press.
- Every node emits a health beacon every two seconds during a game.
- The UI warns at low battery, weak signal, or missed beacons.
- A reconnecting node requests current state before accepting gameplay input.

## Decision gate
Keep ESP-NOW through the first three playtest rounds. Change radio architecture only if measured packet loss, range, power, platform compatibility, or certification economics fail the targets.
