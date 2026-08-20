# Firmware

ESP32-C3 product firmware belongs here.

Recommended layout:

- `wearable-node/` — button input, feedback, battery monitoring, and ESP-NOW communication.
- `hub/` — device pairing, event validation, game state, Wi-Fi, and WebSocket communication.
- `shared/` — shared packet definitions, IDs, constants, and tests.

Implementation requirements are in [Manual 06](../manuals/chapters/06-firmware-manual.md).
