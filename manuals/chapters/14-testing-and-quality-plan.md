> Physical Social Games Business | Manual 14 | Version 0.1 | 2026-08-19

# Testing and Quality Plan

## Test layers
Electrical bring-up; firmware unit/integration; radio performance; battery/charge; mechanical durability; usability/setup; gameplay/fun; safety; environmental; manufacturing end-of-line; regulatory pre-compliance.

## Critical tests
- 1,000 deliberate presses per node: target >=99.5% registered, zero duplicate scoring.
- Six-node burst and simultaneous press test.
- 30-minute continuous session plus four-hour soak.
- Power-cycle node/hub during play and recover without corrupting state.
- Range map in villa rooms, garden, through one/two walls, and while worn.
- Battery runtime and charge-temperature logging across six samples.
- Drop, strap pull, repeated clip use, sweat wipe, and button cycle tests.

## Playtest scorecard
Setup minutes; rules explanation minutes; rounds played; fun 1-5; fairness 1-5; physical comfort 1-5; confusion incidents; false activations; missed activations; safety interventions; quote/request for replay.

## Defect severity
P0: safety, battery, electric shock, data exposure, or uncontrolled heat - stop test and quarantine. P1: wrong winner, lost/duplicate event, unrecoverable disconnect - blocks release. P2: confusing flow or recoverable fault - fix before external beta. P3: polish/cosmetic - prioritize by frequency.

## Release gate
No open P0/P1 defects; all acceptance criteria evidenced; firmware/version locked; six-device serial list recorded; quick-start matches build; risk assessment reviewed; support recovery tested by someone outside the build team.
