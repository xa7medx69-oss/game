> Physical Social Games Business | Manual 01 | Version 0.1 | 2026-08-19

# Product Specification

## Purpose
Define the product to build, who it is for, what the MVP must prove, and what is deliberately postponed.

## Product promise
A reusable box of six wearable electronic controls and one hub that turns a room, garden, majlis, or event space into a library of original real-life games. The physical devices capture deliberate presses; the digital host controls rules, timing, teams, feedback, scoring, and replay.

**MVP decision:** build six wearable nodes plus one hub. The first release is a supervised prototype for adults and older teens, not a children's toy and not a finished waterproof consumer product.

## Flagship game: Pulse Hunt
**Objective:** survive until time expires, or as the Hunter eliminate all Runners.

**Players:** 4-6. One Hunter; 3-5 Runners. For three players, use one Hunter and two Runners with a shorter arena.

**Setup:** charge and assign one pod per player; clip pods high on the upper back; mark a clear play boundary; remove trip hazards; select 5-8 minutes; give Runners a 20-second head start; test every pod before starting.

**Core rules:** the Hunter eliminates a Runner by pressing the Runner's pod once until it confirms with light, vibration, and a sound from the host. Only the device may be touched. No grabbing, blocking doorways, tackling, pushing, hiding in locked rooms, stairs, roads, pools, kitchens, or vehicle areas. Eliminated players freeze, raise a hand, and walk to the safe zone.

**False-press protection:** a valid press is 80-900 ms, followed by a one-second lockout. The hub acknowledges the event. The node changes to red only after acknowledgement; if acknowledgement fails, it vibrates twice and retries.

**Winning:** Runners win if at least one remains alive at zero. The Hunter wins if all are eliminated. Tie-break: fastest complete elimination time.

**Variations:** two Hunters; team rescue with a 3-second teammate hold; three-life mode; shrinking arena announced by the host; silent mode using vibration only; score attack where each confirmed tag is one point and eliminated players re-enter after 20 seconds.

## Seven additional original games

| Game | Core idea | Button use |
|---|---|---|
| Vault Heist | Players steal digital gems while protecting their own vault. | Hold an opponent's pod for 1.5 seconds; a successful steal triggers a cooldown shield. |
| Signal Chain | Cooperative memory and movement challenge. | LEDs cue a changing player sequence; players press their own pods in order before time runs out. |
| Guardian Link | Secret pairs protect one another while completing public tasks. | A partner can cancel one hit by pressing the threatened player's pod within three seconds. |
| Territory Pulse | Teams control physical zones made from detached pods. | Hold a zone pod until its LED changes to the team's color; control scores over time. |
| Echo Match | Fast reaction and pattern recognition. | Players copy color/sound patterns on their own pods; fastest accurate response scores. |
| Silent Mission | Hidden individual objectives inside a normal gathering. | Short and long presses log secret actions; the host reveals completed missions at the end. |
| Relay Rush | Cooperative race against a dynamic route. | The host lights the next player's pod; correct presses pass the signal, mistakes add time. |

## Target user and occasion
Primary: UAE families and friend groups aged 14+, playing at villas, ghabgas, majlis-adjacent spaces, parks, and private gatherings. Secondary: birthdays, youth groups, hospitality venues, and later corporate events. The host should understand setup within three minutes and start a game within five.

## MVP acceptance criteria
- Six nodes connect through one hub without manual repair during a 30-minute session.
- A physical press appears in the UI within 200 ms at 95th percentile in the intended arena.
- At least 99.5% of deliberate valid presses are registered in a 1,000-press bench test.
- No duplicate scoring from one press; every event has a sequence number and acknowledgement.
- A new host can set up Pulse Hunt using only the quick-start page.
- In three playtests, at least 70% of participants request another round and rate fun at 4/5 or higher.

## Not in MVP
Native iOS/Android apps, cloud accounts, children under 14, public matchmaking, GPS, microphones, cameras, precise location tracking, custom injection-moulded tooling, certified waterproofing, and mass-production charging cases.

## Product risks and response
**Accidental activation:** use force/travel tuning, press windows, lockouts, hub acknowledgement, and playtesting. **Unsafe chasing:** strict arena rules, age gate, host briefing, and game modes that do not require running. **Wireless dropouts:** local star topology, retries, health beacons, and no cloud dependency. **One-game novelty:** ship only after three mechanics test well, then expand the library through software.

## Build instruction
Freeze this MVP specification for the first prototype sprint. New ideas go into a backlog unless they affect safety, event reliability, or the fun hypothesis.
