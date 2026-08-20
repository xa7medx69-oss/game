> Physical Social Games Business | Manual 06 | Version 0.1 | 2026-08-19

# Firmware Manual

## Node firmware modules
Board support; unique identity; input/debounce; event queue; ESP-NOW transport; acknowledgements/retry; LED/haptic patterns; battery monitor; health beacon; configuration storage; diagnostics; signed update placeholder.

## Node state machine
`BOOT -> SELF_TEST -> PAIRING/READY -> ACTIVE -> ELIMINATED/LOCKED -> READY -> LOW_BATTERY -> SHUTDOWN`.

A press is accepted only in ACTIVE, within the configured duration, after debounce, and outside lockout. The node immediately shows a neutral "pending" flash, sends the event, and only shows success after the hub acknowledges it.

## Event packet
```c
struct PressEvent {
  uint8_t version;
  uint8_t device_id;
  uint8_t event_type;
  uint8_t flags;
  uint32_t sequence;
  uint32_t uptime_ms;
  uint16_t battery_mv;
  uint16_t auth_tag;
};
```

## Timing targets
Debounce 25 ms; valid press 80-900 ms; hold event at configurable 1,500 or 3,000 ms; local input lockout 1,000 ms; health beacon 2 s; event acknowledgement target under 120 ms; reconnect target under 5 s.

## Hub firmware
Maintain allow-listed node MACs/IDs, receive and deduplicate packets, translate to normalized game events, persist the last state snapshot, serve local web assets, broadcast WebSocket updates, log diagnostics, and send effect/state commands back to nodes.

## Security baseline
Use per-kit keys, allow-listed devices, a random kit Wi-Fi password, no open debug endpoints in a sale unit, no default cloud credential, monotonic sequence handling, and signed firmware updates before commercial launch. Security is a release requirement even when gameplay data seems harmless.

## Development workflow
Use PlatformIO or ESP-IDF with a reproducible lockfile/toolchain. Keep node and hub projects in one repository with shared packet definitions and tests. Every release gets a semantic version, build hash, protocol version, and migration note.

## Firmware test instruction
Unit-test debounce and state transitions; fuzz malformed packets; simulate duplicate/out-of-order packets; power-cycle during game; run 1,000 presses per node; measure latency distribution; test weak signal; test a dead node; and verify that a stale packet cannot re-eliminate or re-score.
