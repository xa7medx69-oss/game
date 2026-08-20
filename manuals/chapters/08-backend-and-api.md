> Physical Social Games Business | Manual 08 | Version 0.1 | 2026-08-19

# Backend and API

## MVP decision
Keep the authoritative backend on the local hub. Cloud is not required to play. This prevents venue internet from becoming a gameplay dependency and keeps the privacy footprint small.

## Local endpoints
- `GET /api/kit` - identity, firmware, protocol, battery summary.
- `POST /api/session` - create a new game session.
- `POST /api/session/{id}/players` - assign player and device.
- `POST /api/session/{id}/start|pause|resume|end` - controlled transitions.
- `GET /api/session/{id}` - current snapshot.
- `WS /ws/session/{id}` - real-time events and state patches.
- `GET /api/diagnostics/export` - support log without personal history by default.

## Event flow
Hardware events are normalized before the game engine sees them. The engine produces domain events such as `PLAYER_ELIMINATED`, `POINT_STOLEN`, or `SEQUENCE_ADVANCED`. The state reducer is deterministic; replaying the same ordered domain events must produce the same state.

## Cloud phase
Add optional accounts, game catalogue updates, telemetry consent, customer support, purchases, and cross-device history only after the local game is stable. Use regional/legal review for hosting and transfers, encrypt in transit and at rest, and collect the minimum necessary data.

## Operational requirements
Structured logs with timestamps and correlation IDs; no secrets in logs; rate limiting; schema validation; idempotency keys for commands; explicit API versioning; device and app compatibility matrix; safe rollback for updates.
