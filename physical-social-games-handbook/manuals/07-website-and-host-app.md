> Physical Social Games Business | Manual 07 | Version 0.1 | 2026-08-19

# Website and Host App

## MVP surface
A responsive local web app served by the hub. The host joins the kit Wi-Fi, opens the local address/QR code, creates players, assigns node IDs, selects a game, performs a device check, and starts.

## Required screens
Welcome/connect; create session; player names; button assignment; game library; rule configuration; preflight device check; live game dashboard; pause/abort; winner/results; diagnostics; export log.

## Live dashboard
Show game title and round, timer, player cards, alive/eliminated or score state, team colors, connection and battery warnings, recent event feed, pause, and emergency end. The host must never need to scroll to reach pause/end on a phone.

## Interaction rules
Use large touch targets, high contrast, color plus icon/text status, visible confirmation for destructive actions, and an offline indicator. Do not rely on sound alone. Avoid mandatory accounts or phone numbers for the prototype.

## QR joining
For the MVP, QR opens the host page on the local hub network; players do not need individual phones. Later, optional spectators can join a read-only screen. Do not expose control URLs without session authorization.

## Suggested stack
TypeScript, Vite, React or lightweight vanilla components, CSS variables/design tokens, and WebSocket transport. Build to static assets embedded in the hub filesystem. Use a shared JSON schema for normalized game events.

## Acceptance tests
360 px phone width; latest Chrome/Edge/Safari UI rendering; no internet; reconnect after the phone locks; large player names; six simultaneous events; timer drift; low battery warning; pause/resume; and safe recovery after hub refresh.
