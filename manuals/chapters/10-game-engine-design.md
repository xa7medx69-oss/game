> Physical Social Games Business | Manual 10 | Version 0.1 | 2026-08-19

# Game Engine Design

## Principle
Hardware emits simple physical facts; game definitions decide meaning. A `PRESS` must not be hard-coded as elimination because the same input may steal points, capture a zone, confirm a sequence, rescue a teammate, or complete a mission.

## Engine components
Game manifest; configuration schema; player/team allocator; deterministic reducer; timers; random seed service; rule validators; effect dispatcher; score/winner evaluator; pause/recovery; replay log.

## Normalized inputs
`PRESS`, `HOLD_STARTED`, `HOLD_COMPLETED`, `RELEASE`, `DEVICE_ONLINE`, `DEVICE_OFFLINE`, `BATTERY_LOW`, and host commands. The engine accepts only validated inputs from known devices in the current session.

## Game definition contract
Each game declares player range, age guidance, space and activity level, setup, allowed roles, defaults, event handlers, state shape, feedback effects, winner conditions, timeouts, reconnection policy, safety warning, and analytics events.

## Determinism
Store the random seed at session creation. Timers are derived from the hub monotonic clock. Every state change emits an immutable domain event. This makes bugs reproducible and supports match replay without video or personal tracking.

## New-game checklist
Write a one-sentence fun promise; define player choices; identify the physical action; cap downtime; add a comeback path; specify ambiguous simultaneous events; define disconnection behavior; run a paper prototype; then implement. Reject games whose novelty depends only on changing colors or scores.

## Portfolio balance
Launch library target: one chase game, one team strategy game, one cooperative memory game, one low-mobility social/hidden-mission game, and one reaction game. This makes the kit useful across villas, indoor gatherings, mixed mobility, and short/long sessions.
