> Physical Social Games Business | Manual 09 | Version 0.1 | 2026-08-19

# Database and Data Model

## Local model
Use a small embedded store or append-only JSON event log on the hub. Core records: Kit, Device, Player, Session, Team, GameDefinition, GameConfig, RawHardwareEvent, DomainEvent, StateSnapshot, Result, FirmwareVersion.

## Minimal fields
`Device(id, hardware_rev, firmware, battery, last_seen, signal)`; `Session(id, game_id, created_at, started_at, ended_at, status, config_version)`; `Player(id, display_name, device_id, team_id)`; `Event(id, session_id, sequence, source, type, payload, received_at)`.

## Data rules
Display names can be temporary nicknames; delete session names after a configurable period; do not collect birth dates, exact location, contacts, photos, audio, or biometrics. Separate operational telemetry from gameplay history and use explicit consent for cloud analytics.

## Retention default
Raw radio diagnostics: seven days on support builds, disabled or bounded in production. Local session history: last 20 sessions unless the host clears it. Cloud customer profile: only while the account is active plus a documented legal/operational retention period.

## Migration and backup
Version every schema and game definition. Apply forward migrations with a backup/snapshot. If migration fails, preserve the previous playable version and show a clear support code.
