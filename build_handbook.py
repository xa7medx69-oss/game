from pathlib import Path
from datetime import date
import re
import textwrap

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


ROOT = Path(r"C:\Users\Ahmed\Documents\ChatGPT\games")
OUT = ROOT / "physical-social-games-handbook"
MANUALS = OUT / "manuals"
ASSETS = OUT / "assets"
QA = OUT / "qa"
for p in (OUT, MANUALS, ASSETS, QA):
    p.mkdir(parents=True, exist_ok=True)

TODAY = date(2026, 8, 19)
PROJECT = "Physical Social Games Business"
SUBTITLE = "A practical, original, UAE-ready roadmap from playable prototype to launch"

SOURCES = {
    "espressif": "https://docs.espressif.com/projects/esp-hardware-design-guidelines/en/latest/esp32c3/product-overview.html",
    "nordic": "https://www.nordicsemi.com/products/nrf52840",
    "tdra": "https://tdra.gov.ae/en/Services/service-card?sID=Cu4EDmeifjoQ1zAbgYV59qYlswfGAtcsK_IwGX8c0LQ",
    "moiat": "https://moiat.gov.ae/en/services/issue-conformity-certificates-for-regulated-products",
    "business": "https://u.ae/en/information-and-services/business/doing-business-on-the-mainland/steps-to-start-a-business-on-the-mainland",
    "privacy": "https://u.ae/en/about-the-uae/digital-uae/data/data-protection-laws.",
    "ip": "https://u.ae/en/information-and-services/business/intellectual-property",
    "vat": "https://www.tax.gov.ae/en/services/vat.registration.aspx",
}


def dedent(s: str) -> str:
    return textwrap.dedent(s).strip() + "\n"


MANUAL_DATA = [
    ("01", "Product Specification", dedent("""
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
    """)),

    ("02", "Hardware Architecture", dedent(f"""
    # Hardware Architecture

    ## Recommended MVP architecture
    Use **six Seeed XIAO ESP32-C3 wearable nodes plus one XIAO ESP32-C3 hub**. Nodes send compact events to the hub over ESP-NOW. The hub runs a local Wi-Fi access point and WebSocket server; a phone or laptop joins the hub network and opens the local web interface. Internet is optional.

    ```text
    [6 wearable nodes] -- ESP-NOW --> [hub] -- local Wi-Fi/WebSocket --> [phone/laptop browser]
    ```

    This is better for the proof-of-concept than connecting six BLE peripherals directly to a phone: the host pairs once with the hub, the game works offline, event timing is centralized, and browser/device Bluetooth differences are removed. ESP32-C3 supports 2.4 GHz Wi-Fi and Bluetooth LE 5, so the same dev board preserves experimentation options. Official overview: {SOURCES['espressif']}

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
    Re-evaluate a Nordic nRF52-class design for lower power and a robust proprietary 2.4 GHz/BLE hub. The nRF52840 supports Bluetooth LE, Bluetooth Mesh, NFC, Thread, and Zigbee, but it raises firmware/tooling and unit-cost complexity. Official product page: {SOURCES['nordic']}

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
    """)),

    ("03", "Circuit Design", dedent("""
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
    """)),

    ("04", "Component List and BOM", dedent("""
    # Component List and BOM

    ## Buy list for one seven-board MVP

    | Item | Qty | Planning cost (AED) | Notes |
    |---|---:|---:|---|
    | Seeed XIAO ESP32-C3 | 7 | 210-315 | Six nodes plus one hub; buy one spare if budget allows. |
    | Protected 500 mAh LiPo | 6 | 90-150 | Use reputable packs with correct polarity and connectors. |
    | Large momentary switches/caps | 8 | 35-80 | Six plus spares; prototype arcade or low-profile caps. |
    | RGB LEDs/rings | 8 | 40-90 | One per node plus spares. |
    | Coin vibration motors | 8 | 40-80 | Optional in prototype round one but recommended. |
    | Buzzers, MOSFETs, diodes, passives | 1 lot | 35-70 | Breadboard/protoboard stage. |
    | Power switches and wire/connectors | 1 lot | 30-60 | Add strain relief and insulation. |
    | USB-C charging cable set/charger | 1 | 45-90 | Multi-port quality charger; never use damaged cells. |
    | Prototype enclosures/straps/clips | 6 | 90-180 | 3D printing, hook-and-loop, fabric straps. |
    | Hub case and optional OLED | 1 | 35-75 | OLED is diagnostic, not required. |

    **Expected prototype total:** approximately AED 615-1,190 before tools, shipping, and failed parts. This is a planning range, not a supplier quote.

    ## Required tools
    Temperature-controlled soldering iron, lead-free or appropriately handled solder, flux, flush cutters, wire stripper, heat-shrink, multimeter, USB power meter, bench supply with current limit, small scale/calipers, and access to a 3D printer. A basic logic analyser is useful for debugging.

    ## Supplier discipline
    Record supplier, order date, exact MPN, lot/batch, price, lead time, datasheet URL, and substitute. Do not silently swap batteries, radios, switches, or plastics. For each critical part maintain one approved alternate and document the validation needed after substitution.

    ## Cost-down ladder
    Prototype dev boards first; then one custom PCB; then panelized PCBs and a simple machined/printed enclosure; only after demand validation consider injection moulding and a charging case. Do not spend on tooling before event reliability and repeat-play are proven.
    """)),

    ("05", "Prototype Design", dedent("""
    # Prototype Design

    ## Prototype stages
    **P0 - one-button proof:** one node, one hub, one browser. Demonstrate press, acknowledgement, UI state, and reconnect.

    **P1 - six wired-on-bench nodes:** six powered nodes on a table. Run 1,000 scripted/manual presses, duplicates, simultaneous events, resets, and low-signal tests.

    **P2 - wearable alpha:** battery power, 3D-printed housings, straps/clips, LED and vibration. Conduct controlled walking-speed games.

    **P3 - playtest kit:** refined enclosures, labelled IDs, carry tray, charging checklist, host guide, and three complete game modes.

    ## Enclosure target
    Prototype diameter 58-65 mm; depth 20-28 mm; mass below 70 g; 45-50 mm press area; minimum 2 mm rounded shell walls; no sharp external corners; recessed power switch; USB-C access that cannot be hit during play. Use a replaceable fabric/clip interface so the electronics do not need redesign when wear testing changes.

    ## Attachment
    Start with a wide hook-and-loop armband/back strap plus a rounded clip adapter. Avoid strong magnets near medical devices and avoid lanyards or neck mounting. A production attachment should release under abnormal load and must not use a sharp metal edge against the body.

    ## Build order
    1. Flash unique IDs and radio keys onto all boards.
    2. Assemble hub and verify all six nodes on USB power.
    3. Add one battery-powered node and measure current/temperature.
    4. Clone the validated wiring to the other five.
    5. Print one enclosure, conduct fit/press/antenna checks, then print the set.
    6. Run bench reliability before any active play.
    7. Run a safety briefing and low-speed playtest before chase variants.

    ## Exit criteria
    Six nodes run 30 minutes without a reset, every node reconnects within five seconds after power cycling, the host sees battery/signal status, and two unfamiliar hosts can set up using the quick-start instructions.
    """)),

    ("06", "Firmware Manual", dedent("""
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
    """)),

    ("07", "Website and Host App", dedent("""
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
    """)),

    ("08", "Backend and API", dedent("""
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
    """)),

    ("09", "Database and Data Model", dedent("""
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
    """)),

    ("10", "Game Engine Design", dedent("""
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
    """)),

    ("11", "UI and UX Manual", dedent("""
    # UI and UX Manual

    ## Experience target
    Open box to first round in under five minutes. The host should answer only: who is playing, which pod belongs to whom, which game, and any optional rule.

    ## Setup flow
    1. Power hub and six pods.
    2. Scan QR or join the kit network.
    3. Watch six device tiles turn ready.
    4. Tap a physical pod; the matching tile animates; enter player name.
    5. Choose a game and accept the safety briefing.
    6. Run a three-second feedback test and start.

    ## Status language
    Use plain words: Ready, Weak signal, Low battery, Paused, Eliminated, Reconnecting. Pair colors with icons and text. Red means unavailable/eliminated, amber means attention, teal means active/ready, and blue means neutral/system.

    ## Failure recovery
    If a pod drops, pause automatically only when the current game cannot continue fairly. Show which pod failed and a one-tap `Replace device` flow. Preserve the timer and state. Never make the host rebuild all player assignments.

    ## Accessibility
    Minimum 4.5:1 contrast for normal text, 44 px touch targets, scalable type, no meaning by color alone, optional reduced motion, captions/text equivalents for sound, and haptic alternatives. Include low-mobility games and a no-running filter in the library.

    ## Research script
    Observe without coaching: unbox, charge, assign, choose, start, recover one disconnected pod, and explain a result. Record completion time, errors, questions, and confidence. Improve the flow before adding more decoration.
    """)),

    ("12", "Industrial Design Manual", dedent("""
    # Industrial Design Manual

    ## Design intent
    Premium, friendly, obvious to press, comfortable on clothing, and durable enough for repeated gatherings. It should look like a social-game object, not a medical alarm, security tracker, or television-show prop.

    ## Geometry targets
    Consumer goal 55-60 mm diameter, 18-23 mm body depth, under 55 g, 42-48 mm press surface, 1.5-2.5 N activation, rounded R2+ edges, recessed power switch, protected USB-C or contact charging, and a replaceable attachment module.

    ## CMF direction
    Deep navy main shell, soft teal active surface, coral/amber state accents, warm off-white case interior. Matte fine texture hides scratches; glossy light pipe only where needed. Avoid excessive neon and avoid six unrelated toy colors.

    ## Feedback hierarchy
    Host screen carries complex information. Pod feedback stays glanceable: color/state, one short vibration, and rare local sound. Do not expose raw connection codes during normal play.

    ## Durability goals
    Prototype: 1 m drop to wood/carpet, 5,000 presses. Engineering validation: repeated 1.2 m multi-face drops, 100,000 press cycles, sweat/skin-oil wipe tests, 0-45 C operating evaluation, transport vibration, and ingress tests appropriate to the final claim. Never print an IP rating without certified testing.

    ## Human-factors tests
    Different fabrics, abayas/jackets/T-shirts, standing/sitting, left/right hand, reduced grip, accidental leaning, and quick device removal. Test antenna performance while worn, not only on a bench.
    """)),

    ("13", "3D-Printable Enclosure Instructions", dedent("""
    # 3D-Printable Enclosure Instructions

    ## Parametric starting dimensions
    Outer diameter 62 mm; total depth 25 mm; wall 2.0 mm; top-cap travel 0.8 mm; press disk 48 mm; LED light-pipe groove 2.5 mm; PCB standoffs sized to the selected dev board; battery pocket with 0.6-1.0 mm clearance; four M2 heat-set inserts or snap features; clip rail on the back.

    ## Print recommendation
    Print housing in PETG for tougher functional parts; TPU 95A for the press membrane/bumper; 0.2 mm layers; four perimeters; 30-40% infill; orient the back shell flat and the press cap face-up. Keep support away from snap and gasket surfaces.

    ## Assembly sequence
    Deburr; fit heat-set inserts; install switch actuator; place light pipe; mount PCB; route motor and battery without crossing antenna keep-out; secure battery with removable foam tape; add insulation; close shell; fit clip/strap; verify press around the full cap; perform radio and charge tests.

    ## Safety rules
    No exposed screw tips, conductive edges, pinched battery, loose cell, sharp clip, strong neck attachment, or inaccessible power isolation. Reject any printed part with cracks, warped battery pocket, or a cap that can jam down.

    ## CAD deliverables
    Native parametric file, STEP, STL/3MF per part, drawing with critical dimensions/tolerances, exploded assembly, BOM, print orientation sheet, revision table, and a change log tied to playtest findings.
    """)),

    ("14", "Testing and Quality Plan", dedent("""
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
    """)),

    ("15", "Manufacturing Plan", dedent(f"""
    # Manufacturing Plan

    ## Phases
    EVT validates electronics and radio; DVT validates mechanical design, safety, reliability, and pre-compliance; PVT validates tooling, assembly process, fixtures, yield, packaging, and traceability; mass production begins only after signed exit reports.

    ## Supplier package
    Gerbers, drawings, tolerances, golden sample, approved vendor list, controlled BOM, firmware image/hash, programming/test procedure, cosmetic standards, packaging spec, serial/lot format, change-control process, and defect criteria.

    ## End-of-line test
    Program identity/keys; verify current draw; main press; LED colors; vibration/buzzer; radio packet/ack; RSSI at fixture distance; battery voltage; charge detection; firmware/version; label scan; final visual inspection. Store pass/fail by serial and lot.

    ## Yield targets
    Pilot first-pass yield >=90%; production target >=97% after process stabilization. Any supplier substitution requires documented equivalence and targeted revalidation, especially battery, radio module, antenna materials, switch, plastics, and charger.

    ## UAE market gates
    Wireless telecommunications equipment intended for UAE sale requires TDRA review/type approval under the applicable regime; current service requirements can include technical data sheets and accredited radio/EMF/safety reports. Confirm classification before production: {SOURCES['tdra']}

    Products subject to UAE technical regulations may require a MoIAT conformity certificate and accredited test report: {SOURCES['moiat']}

    Do not print conformity marks, IP ratings, age claims, or certifications until formally earned and documented.
    """)),

    ("16", "Packaging and Quick-Start", dedent("""
    # Packaging and Quick-Start

    ## Box contents
    Six numbered wearable pods, one hub, six attachment straps/clips, one charging solution, USB cable, quick-start card, safety card, game library card/QR, warranty/support information, and inventory checklist.

    ## Opening experience
    The lid shows three steps: power, join, play. Devices sit in numbered wells. Status LEDs are visible during charging. The host should not need to remove a lower tray before learning setup.

    ## Quick-start copy
    1. Charge all pods and power the hub.
    2. Place pods high on the upper back using the provided attachment.
    3. Scan the hub QR and join the kit network.
    4. Tap each pod to assign a player.
    5. Choose a game, clear the play area, review the safety rules, and start.

    ## Safety card
    Use only in a clear private play area. No pushing, grabbing, tackling, road/pool/stair play, or device placement on neck/face. Touch only the wearable device. Stop for pain, overheating, damaged battery, cracked housing, or unsafe behavior. MVP use is supervised and 14+.

    ## Sustainability
    Right-size the box, minimize mixed plastics, make paper separable, use replaceable straps, design devices for battery/service access by trained technicians, and avoid a permanently decorative insert that prevents efficient shipping.
    """)),

    ("17", "Branding Manual", dedent("""
    # Branding Manual

    ## Positioning
    A modern physical play platform that makes gatherings active, surprising, and replayable. The brand is original and should never depend on a television programme, copied terminology, name-tag imagery, or recognizable show styling.

    ## Working brand criteria
    Short in Arabic and English, easy to pronounce, broad enough for many game types, available as a UAE trademark/domain/social handle, and not descriptive of only chasing or buttons. Do not select a final name before clearance.

    ## Personality
    Premium but welcoming; energetic without aggression; clever without being complicated; locally relevant without stereotypes. Use concise host language and celebratory feedback.

    ## Visual direction
    Deep navy foundation, teal action color, coral highlight, amber reward, warm off-white background. Rounded geometry echoes the pod without making every graphic a button. Photography shows real groups and safe open spaces, not staged combat poses.

    ## Voice examples
    Prefer: "Choose a game. Clip in. Start the round." "Six ready." "Signal confirmed." Avoid: militaristic language, copied show phrases, humiliation after elimination, and claims such as "unbreakable" or "works everywhere."

    ## Naming sprint
    Generate 50 candidates; remove narrow/copycat names; screen search engines, domains, app stores, trade names, and trademarks; test pronunciation with Arabic/English speakers; shortlist five; obtain professional clearance; then design the identity.
    """)),

    ("18", "Pricing Manual", dedent("""
    # Pricing Manual

    ## Pricing method
    Price from landed cost and customer value, not only competitor markup. Landed cost includes product, hub, battery, accessories, packaging, freight, customs, certification amortization, payment fees, warranty reserve, returns, fulfilment, support, and VAT treatment.

    ## Working model
    If mature landed COGS is AED 220-320, a direct-to-consumer price hypothesis of AED 699-999 gives room for payment, fulfilment, warranty, marketing, and channel discounts. Do not publish this range until willingness-to-pay and supplier quotes are tested.

    ## Validation
    Test three concept prices with the same offer; collect purchase intent and deposits, not compliments; compare family-use vs event-use positioning; test whether included games, carry/charging case, and bilingual content increase value; record refund/objection reasons.

    ## Margin guardrails
    Target DTC gross margin before marketing of 55-65% once scaled. Wholesale requires sufficient margin for the retailer and your support/warranty burden. Hardware subscriptions should not be necessary for basic play.

    ## Offer ladder
    Core six-player kit; accessory straps/replacement pods; larger group expansion; paid premium game packs only after a strong free library; event/corporate bundle with facilitation tools; rental for trial and hospitality.
    """)),

    ("19", "Business Model Manual", dedent("""
    # Business Model Manual

    ## Primary model
    Sell a reusable hardware kit with an included permanent game library. Grow lifetime value through optional expansion pods, accessories, new premium games, gifting, events, and replacement/service—not by disabling core functionality behind a subscription.

    ## Customer segments
    UAE families/friends; gift buyers; party/event organizers; hospitality venues; youth/community groups after safety validation; later corporate team-building providers and regional distributors.

    ## Channels
    Direct online pre-orders, pop-up demonstrations, partnerships with event organizers, curated gift retailers, corporate pilots, and rentals. The product needs live demonstration because the value is experiential.

    ## Unit-economics worksheet
    Revenue minus landed COGS, payment fees, fulfilment, warranty/returns, customer support, and variable content/service cost equals contribution margin. Compare contribution margin against blended customer acquisition cost and return/refund rate.

    ## Defensibility
    Reliable hardware/software integration, original tested game library, brand/community, proprietary event/game engine, industrial design, manufacturing know-how, and accumulated anonymous playtest insights. Game ideas alone are weak protection.
    """)),

    ("20", "Marketing Manual", dedent("""
    # Marketing Manual

    ## Core message
    One compact kit changes what a gathering can become: active games, team challenges, reaction rounds, and low-mobility social missions using the same six devices.

    ## Proof content
    Show setup in under five minutes; a clear 15-second gameplay loop; real reactions; multiple distinct games; offline play; safe play-space setup; kit reset and replay. Avoid implying association with any television show.

    ## Launch content plan
    Founder build story; product reveal; how-it-works animation; three game demos; host tutorial; UAE villa/ghabga use case; reliability test; playtester reactions; pricing/pre-order; delivery/support updates.

    ## Demonstration funnel
    Short video -> landing page -> email/waitlist -> small paid or deposit-backed playtest -> pre-order. Measure click-through, qualified signup, demo attendance, deposit conversion, acquisition cost, and refund intent.

    ## Community loop
    Invite owners to vote on new mechanics, publish seasonal game updates, run host challenges, and reward useful playtest reports. Moderate for safety and never collect children's content without a dedicated compliant process.
    """)),

    ("21", "Launch Strategy", dedent("""
    # Launch Strategy

    ## Stage 1 - proof (weeks 1-4)
    One-node technical proof, six-node bench test, Pulse Hunt paper prototype, and ten customer interviews. Gate: reliable events and clear demand for active social play.

    ## Stage 2 - alpha (weeks 5-8)
    Wearable housings, local host UI, Pulse Hunt plus Signal Chain and Echo Match, five supervised groups. Gate: repeat play, safe behavior, setup under five minutes.

    ## Stage 3 - paid pilot (weeks 9-12)
    Ten kits or facilitated sessions, deposits/fees, support logs, willingness-to-pay, and supplier DFM review. Gate: no open critical defects and credible unit economics.

    ## Stage 4 - engineering validation
    Custom PCB, DVT enclosure, battery/charger validation, pre-compliance testing, TDRA/MoIAT classification, privacy/legal documents, brand clearance, and pilot manufacturer.

    ## Stage 5 - controlled UAE launch
    Limited inventory, local fulfilment, bilingual quick-start/safety material, responsive support, spare stock, and weekly defect/review meeting. Expand only after return, support, and reliability metrics stabilize.

    ## Stop conditions
    Unsafe common behavior, persistent missed/duplicate hits, less than 50% desire to replay after iteration, setup consistently above ten minutes, supplier cost that destroys contribution margin, or regulatory path inconsistent with the model.
    """)),

    ("22", "UAE Business and Compliance", dedent(f"""
    # UAE Business and Compliance

    ## Business setup
    Select the exact activities first: e-commerce/trading of electronic games or toys, software/content, import, and possibly event services. Mainland setup steps include activity, legal form, trade licence, trade name, initial approval, premises, and additional approvals: {SOURCES['business']}

    A free-zone company may face regulated access to mainland sales, so model the intended sales channel and distributor structure before choosing. Obtain professional UAE legal/accounting advice for the specific emirate and activities.

    ## Wireless and product conformity
    Confirm TDRA equipment classification and supplier registration/type-approval requirements before commercial import or sale. Current policy covers entities involved in registration, approval, import, distribution, display, sale, and commercial use of telecommunications equipment. Start here: {SOURCES['tdra']}

    Confirm whether the kit, battery charger, and toy/electronic classification are regulated under MoIAT conformity programmes. A conformity application for regulated products requires a valid trade licence and accredited test report: {SOURCES['moiat']}

    ## Data protection
    UAE Federal Decree-Law No. 45 of 2021 creates obligations for personal-data processing, security, consent/other lawful bases, rights, and cross-border transfers. Minimize data and keep MVP play local: {SOURCES['privacy']}

    ## Tax
    As of this handbook date, UAE resident businesses must register for VAT when taxable supplies/imports exceed AED 375,000 in the prior 12 months or are expected to exceed it in the next 30 days; voluntary threshold is AED 187,500. Verify current rules and your facts: {SOURCES['vat']}

    ## Practical compliance file
    Trade licence/activities; customs/importer setup; TDRA correspondence and certificate; MoIAT classification/certificate; battery transport documents including UN 38.3; accredited safety/EMC/radio reports; supplier declarations; labels/manuals; risk assessment; privacy notice/terms; warranty/returns; incident/recall plan; serial/lot traceability.

    ## Important limitation
    This manual is a planning checklist, not legal, tax, safety-certification, or regulatory advice. Confirm requirements with the relevant UAE authorities and qualified advisors before sale, import, claims, or children's use.
    """)),

    ("23", "Legal and IP Manual", dedent(f"""
    # Legal and IP Manual

    ## Originality boundary
    Do not use the Running Man name, logos, visual identity, character/personality references, episode formats, signature props, promotional comparisons suggesting affiliation, or copied game presentations. Document independent development of names, rules, artwork, hardware, UI, and marketing.

    ## Protection stack
    Trademark for final brand/product names and logos; copyright for software, artwork, manuals, sound, and written game expression; industrial design for distinctive product appearance; patent/utility analysis for genuinely novel technical inventions; contracts and confidentiality for supplier/source assets and unreleased designs.

    UAE's Ministry of Economy and Tourism is the competent authority for trademarks, copyright registrations, patents, and industrial property. UAE guidance also notes that schemes, rules, or methods for playing games are not granted patents as such, while qualifying technical inventions may be evaluated separately: {SOURCES['ip']}

    ## Before naming or launch
    Conduct UAE and relevant international trademark clearance; search company/trade names, domains, app stores, marketplaces, and social handles; document results; obtain professional advice on Nice classes likely covering games/toys, electronics, software, retail, and entertainment services.

    ## Contracts
    Founder IP assignment; contractor work-made/assignment and confidentiality; firmware/software licences and open-source notices; manufacturer NDA, tooling ownership, quality/change controls, and no unauthorized production; playtester consent/release; retailer/distributor terms; privacy/website terms; warranty/returns.

    ## Claims discipline
    Substantiate range, battery life, water resistance, age suitability, safety, durability, and number of included games. Do not claim certification before approval or use "patented" unless accurate for the jurisdiction/status.
    """)),

    ("24", "Future Expansion Roadmap", dedent("""
    # Future Expansion Roadmap

    ## Phase A - prove the platform
    Six nodes, one hub, three excellent games, offline host UI, adult/older-teen supervised use, reliable event protocol, and repeat-play evidence.

    ## Phase B - consumer kit
    Custom low-power PCB, certified battery/charging, robust enclosure and attachment, manufactured case, five-game library, bilingual content, updates, diagnostics, and UAE compliance.

    ## Phase C - library and accessories
    Expansion pods to 8/12 players, floor/territory bases, wrist/vest attachments, spectator mode, seasonal game packs, creator tools with rule validation, and accessibility filters.

    ## Phase D - events and corporate
    Multi-kit orchestration, facilitator dashboard, team analytics that avoid intrusive employee profiling, rental fleet management, quick sanitation/service, venue mode, and certified facilitators.

    ## Phase E - regional/global
    Arabic/English localization quality, country-specific radio/product approvals, distributor/service strategy, regional cloud/data review, climate testing, and localized game research.

    ## Technology options to revisit
    Lower-power Nordic radio architecture; charging contacts/case; secure OTA; larger mesh/multi-hub events; UWB only if precise ranging creates a proven game mechanic; motion sensors only after button simplicity is exhausted. Avoid feature creep that weakens reliability or setup.

    ## North-star metrics
    Median setup under five minutes; session success >99%; repeat play within 30 days; games played per kit; support contacts per 100 sessions; return rate; contribution margin; safety incidents; and percentage of active owners using more than one game.

    ## Next seven actions
    Buy the seven-board BOM; build the one-node proof; run a paper Pulse Hunt; implement acknowledged events; assemble six bench nodes; conduct the 1,000-press test; schedule the first supervised playtest and record the scorecard.
    """)),
]


def copy_generated_assets():
    mapping = {
        Path(r"C:\Users\Ahmed\.codex\generated_images\01a01b74-8c92-71b3-9687-67c01ac3110e\exec-41740bee-1a36-4158-88f0-01b359aaf3b6.png"): ASSETS / "kit-concept.png",
        Path(r"C:\Users\Ahmed\.codex\generated_images\01a01b74-8c92-71b3-9687-67c01ac3110e\exec-e2f19538-d2f4-4fd9-8ffd-a1295b473bac.png"): ASSETS / "gameplay-villa.png",
        Path(r"C:\Users\Ahmed\.codex\generated_images\01a01b74-8c92-71b3-9687-67c01ac3110e\exec-ed5b45f4-bc54-4c7e-9641-83157daf60d2.png"): ASSETS / "wearable-exploded.png",
    }
    for src, dst in mapping.items():
        if not dst.exists():
            dst.write_bytes(src.read_bytes())


def font(size=34, bold=False):
    path = Path(r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf")
    return ImageFont.truetype(str(path), size=size)


NAVY = "#102A43"
TEAL = "#0C7C86"
CORAL = "#E76F51"
AMBER = "#E9A23B"
PALE = "#F4F7FA"
GRAY = "#5C6770"


def diagram_architecture(path: Path):
    img = Image.new("RGB", (1800, 900), "#F7F9FB")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((80, 90, 1720, 790), 34, fill="white", outline="#D5DEE7", width=4)
    d.text((130, 125), "OFFLINE-FIRST MVP ARCHITECTURE", font=font(48, True), fill=NAVY)
    nodes = [(150, 300+i*90, f"Wearable {i+1}") for i in range(6)]
    for x,y,label in nodes:
        d.rounded_rectangle((x,y,x+310,y+62), 20, fill="#E4F4F4", outline=TEAL, width=3)
        d.ellipse((x+20,y+12,x+58,y+50), fill=CORAL)
        d.text((x+78,y+13), label, font=font(27, True), fill=NAVY)
        d.line((x+310,y+31,750,y+31), fill=TEAL, width=3)
    d.rounded_rectangle((750,330,1110,610), 30, fill=NAVY, outline=NAVY)
    d.text((833,395), "CENTRAL", font=font(34, True), fill="white")
    d.text((872,445), "HUB", font=font(52, True), fill="#7FDBD4")
    d.text((802,525), "state + rules", font=font(28), fill="white")
    d.line((1110,470,1340,470), fill=AMBER, width=8)
    d.polygon([(1340,470),(1308,450),(1308,490)], fill=AMBER)
    d.rounded_rectangle((1340,330,1650,610), 30, fill="#FFF4DF", outline=AMBER, width=4)
    d.text((1424,392), "PHONE", font=font(34, True), fill=NAVY)
    d.text((1390,445), "OR LAPTOP", font=font(30, True), fill=NAVY)
    d.text((1383,515), "local web app", font=font(25), fill=GRAY)
    d.text((550,690), "ESP-NOW", font=font(27, True), fill=TEAL)
    d.text((1138,500), "Wi-Fi + WebSocket", font=font(25, True), fill="#A76300")
    d.text((550,746), "No cloud required for gameplay", font=font(29, True), fill=NAVY)
    img.save(path, quality=95)


def diagram_node(path: Path):
    img = Image.new("RGB", (1500, 1050), PALE)
    d = ImageDraw.Draw(img)
    d.text((90, 70), "WEARABLE NODE - FUNCTIONAL BLOCKS", font=font(46, True), fill=NAVY)
    blocks = [
        (120,260,420,390,"PRESS CAP","momentary switch",CORAL),
        (600,220,940,430,"ESP32-C3","input + radio + state",NAVY),
        (1080,180,1380,310,"LED RING","visual state",TEAL),
        (1080,380,1380,510,"VIBRATION","haptic confirm",AMBER),
        (1080,580,1380,710,"BUZZER","optional local cue",CORAL),
        (120,620,420,770,"500 mAh LiPo","protected battery",TEAL),
        (560,690,970,850,"POWER + CHARGE","switch / USB-C / sensing",NAVY),
    ]
    for x1,y1,x2,y2,title,sub,c in blocks:
        d.rounded_rectangle((x1,y1,x2,y2),24,fill="white",outline=c,width=5)
        d.text((x1+24,y1+28),title,font=font(29,True),fill=c)
        d.text((x1+24,y1+75),sub,font=font(23),fill=GRAY)
    lines=[((420,325),(600,325)),((940,275),(1080,245)),((940,330),(1080,445)),((940,385),(1080,645)),((420,695),(560,760)),((765,690),(765,430))]
    for a,b in lines:
        d.line((*a,*b),fill="#8996A3",width=5)
    d.rounded_rectangle((115,900,1385,980),22,fill="#E7EEF5")
    d.text((165,919),"Hub acknowledgement makes the official result; the node never scores itself.",font=font(27,True),fill=NAVY)
    img.save(path, quality=95)


def diagram_roadmap(path: Path):
    img=Image.new("RGB",(1800,900),"white")
    d=ImageDraw.Draw(img)
    d.text((100,80),"12-WEEK MVP PATH",font=font(50,True),fill=NAVY)
    phases=[
        ("WEEKS 1-2","ONE-NODE PROOF","Press -> hub -> browser",TEAL),
        ("WEEKS 3-4","SIX-NODE BENCH","1,000 presses + recovery",CORAL),
        ("WEEKS 5-8","WEARABLE ALPHA","Enclosures + 3 games",AMBER),
        ("WEEKS 9-12","PAID PILOT","10 kits/sessions + deposits",NAVY),
    ]
    x=110
    for i,(wk,title,sub,c) in enumerate(phases):
        w=370
        d.rounded_rectangle((x,260,x+w,650),30,fill=PALE,outline=c,width=6)
        d.text((x+30,300),wk,font=font(27,True),fill=c)
        d.text((x+30,380),title,font=font(31,True),fill=NAVY)
        for j,line in enumerate(textwrap.wrap(sub,22)):
            d.text((x+30,460+j*36),line,font=font(25),fill=GRAY)
        if i<3:
            d.line((x+w,455,x+w+70,455),fill="#9AA5B1",width=6)
            d.polygon([(x+w+70,455),(x+w+45,438),(x+w+45,472)],fill="#9AA5B1")
        x+=430
    d.text((110,740),"Gate each phase on evidence: reliability, safety, repeat play, then economics.",font=font(32,True),fill=NAVY)
    img.save(path, quality=95)


def write_markdown_files():
    for num, title, body in MANUAL_DATA:
        path = MANUALS / f"{num}-{re.sub(r'[^a-z0-9]+','-',title.lower()).strip('-')}.md"
        front = f"> {PROJECT} | Manual {num} | Version 0.1 | {TODAY.isoformat()}\n\n"
        path.write_text(front + body, encoding="utf-8")
    index = [f"# {PROJECT} - Manual Index", "", SUBTITLE, "", "## Files", ""]
    for num,title,_ in MANUAL_DATA:
        fn=f"{num}-{re.sub(r'[^a-z0-9]+','-',title.lower()).strip('-')}.md"
        index.append(f"{num}. [{title}](manuals/{fn})")
    index += ["", "## Start here", "", "Read Manual 01, buy from Manual 04, follow Manual 05, implement Manual 06, and use Manual 14 before active playtests.", ""]
    (OUT/"README.md").write_text("\n".join(index),encoding="utf-8")


def set_cell_shading(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=tcPr.find(qn("w:shd"))
    if shd is None:
        shd=OxmlElement("w:shd"); tcPr.append(shd)
    shd.set(qn("w:fill"),fill)


def set_cell_margins(cell, top=100, start=120, bottom=100, end=120):
    tc=cell._tc; tcPr=tc.get_or_add_tcPr(); tcMar=tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar=OxmlElement("w:tcMar"); tcPr.append(tcMar)
    for m,v in (("top",top),("start",start),("bottom",bottom),("end",end)):
        node=tcMar.find(qn(f"w:{m}"))
        if node is None:
            node=OxmlElement(f"w:{m}"); tcMar.append(node)
        node.set(qn("w:w"),str(v)); node.set(qn("w:type"),"dxa")


def set_table_widths(table, widths_in):
    table.autofit=False
    for row in table.rows:
        for cell,w in zip(row.cells,widths_in):
            cell.width=Inches(w); set_cell_margins(cell)
    tblPr=table._tbl.tblPr
    tblW=tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW=OxmlElement("w:tblW"); tblPr.append(tblW)
    tblW.set(qn("w:w"),"9360"); tblW.set(qn("w:type"),"dxa")
    tblInd=tblPr.find(qn("w:tblInd"))
    if tblInd is None:
        tblInd=OxmlElement("w:tblInd"); tblPr.append(tblInd)
    tblInd.set(qn("w:w"),"120"); tblInd.set(qn("w:type"),"dxa")


def mark_table_header(table):
    trPr = table.rows[0]._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    trPr.append(header)


def add_picture_with_alt(run, path, alt, width=None, height=None):
    shape = run.add_picture(str(path), width=width, height=height)
    shape._inline.docPr.set("descr", alt)
    shape._inline.docPr.set("title", alt)
    return shape


def add_page_field(paragraph):
    paragraph.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    run=paragraph.add_run("Page ")
    fld=OxmlElement("w:fldSimple"); fld.set(qn("w:instr"),"PAGE"); paragraph._p.append(fld)
    for r in paragraph.runs:
        r.font.name="Arial"; r.font.size=Pt(8); r.font.color.rgb=RGBColor(90,103,112)


def add_bullet(doc, text):
    p=doc.add_paragraph(style="List Bullet")
    p.add_run(text)
    return p


def add_numbered_group(doc, items):
    numbering = doc.part.numbering_part.element
    style = doc.styles["List Number"]._element
    style_num = style.pPr.numPr.numId.val
    base_num = next(n for n in numbering.findall(qn("w:num")) if int(n.get(qn("w:numId"))) == style_num)
    abstract_id = base_num.find(qn("w:abstractNumId")).get(qn("w:val"))
    existing = [int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))]
    new_id = max(existing) + 1
    num = OxmlElement("w:num"); num.set(qn("w:numId"), str(new_id))
    abstract = OxmlElement("w:abstractNumId"); abstract.set(qn("w:val"), abstract_id); num.append(abstract)
    override = OxmlElement("w:lvlOverride"); override.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:startOverride"); start.set(qn("w:val"), "1"); override.append(start); num.append(override)
    numbering.append(num)
    for item in items:
        p = doc.add_paragraph(style="List Number")
        p.add_run(item)
        numPr = p._p.get_or_add_pPr().get_or_add_numPr()
        numPr.get_or_add_ilvl().val = 0
        numPr.get_or_add_numId().val = new_id


def add_code(doc, code):
    for line in code.rstrip().splitlines():
        p=doc.add_paragraph()
        p.style=doc.styles["Normal"]
        p.paragraph_format.left_indent=Inches(.25)
        p.paragraph_format.space_after=Pt(0)
        r=p.add_run(line or " "); r.font.name="Consolas"; r.font.size=Pt(8.5); r.font.color.rgb=RGBColor(30,55,75)
        shd=OxmlElement("w:shd"); shd.set(qn("w:fill"),"EEF3F6"); p._p.get_or_add_pPr().append(shd)


def add_markdown(doc, md):
    lines=md.splitlines(); i=0
    while i<len(lines):
        line=lines[i].rstrip()
        if not line:
            i+=1; continue
        if line.startswith("```"):
            buf=[]; i+=1
            while i<len(lines) and not lines[i].startswith("```"):
                buf.append(lines[i]); i+=1
            add_code(doc,"\n".join(buf)); i+=1; continue
        if line.startswith("# "):
            i+=1; continue
        if line.startswith("## "):
            p=doc.add_paragraph(line[3:],style="Heading 2"); p.paragraph_format.keep_with_next=True; i+=1; continue
        if line.startswith("### "):
            p=doc.add_paragraph(line[4:],style="Heading 3"); p.paragraph_format.keep_with_next=True; i+=1; continue
        if line.startswith("| "):
            rows=[]
            while i<len(lines) and lines[i].startswith("|"):
                vals=[v.strip() for v in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-+:?",v) for v in vals): rows.append(vals)
                i+=1
            if rows:
                table=doc.add_table(rows=len(rows),cols=len(rows[0])); table.style="Table Grid"; table.alignment=WD_TABLE_ALIGNMENT.LEFT
                lens=[max(len(r[c]) if c<len(r) else 0 for r in rows) for c in range(len(rows[0]))]; total=max(sum(lens),1)
                widths=[max(.7,6.5*l/total) for l in lens]; scale=6.5/sum(widths); widths=[w*scale for w in widths]
                for ri,row in enumerate(rows):
                    for ci,val in enumerate(row):
                        cell=table.cell(ri,ci); cell.text=val; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
                        if ri==0: set_cell_shading(cell,"DDEBF2")
                        for run in cell.paragraphs[0].runs:
                            run.font.name="Arial"; run.font.size=Pt(8.5); run.bold=(ri==0); run.font.color.rgb=RGBColor(16,42,67)
                set_table_widths(table,widths)
                mark_table_header(table)
                doc.add_paragraph().paragraph_format.space_after=Pt(0)
            continue
        if line.startswith("- "):
            add_bullet(doc,line[2:]); i+=1; continue
        if re.match(r"^\d+\. ",line):
            items=[]
            while i<len(lines) and re.match(r"^\d+\. ",lines[i]):
                items.append(re.sub(r"^\d+\. ","",lines[i])); i+=1
            add_numbered_group(doc,items); continue
        p=doc.add_paragraph()
        parts=re.split(r"(\*\*.+?\*\*|`.+?`)",line)
        for part in parts:
            if part.startswith("**") and part.endswith("**"):
                r=p.add_run(part[2:-2]); r.bold=True
            elif part.startswith("`") and part.endswith("`"):
                r=p.add_run(part[1:-1]); r.font.name="Consolas"; r.font.size=Pt(9)
            else: p.add_run(part)
        i+=1


def style_document(doc):
    sec=doc.sections[0]; sec.top_margin=Inches(0.8); sec.bottom_margin=Inches(0.75); sec.left_margin=Inches(0.85); sec.right_margin=Inches(0.85); sec.header_distance=Inches(.35); sec.footer_distance=Inches(.35)
    styles=doc.styles
    normal=styles["Normal"]; normal.font.name="Arial"; normal.font.size=Pt(10); normal.font.color.rgb=RGBColor(35,48,60); normal.paragraph_format.space_after=Pt(6); normal.paragraph_format.line_spacing=1.15
    for name,size,color,before,after in [("Title",30,NAVY,0,8),("Subtitle",13,GRAY,0,12),("Heading 1",20,NAVY,18,8),("Heading 2",14,TEAL,14,6),("Heading 3",11,CORAL,10,4)]:
        st=styles[name]; st.font.name="Arial"; st.font.size=Pt(size); st.font.bold=name!="Subtitle"; st.font.color.rgb=RGBColor.from_string(color.strip("#")); st.paragraph_format.space_before=Pt(before); st.paragraph_format.space_after=Pt(after)
    for name in ("List Bullet","List Number"):
        st=styles[name]; st.font.name="Arial"; st.font.size=Pt(10); st.paragraph_format.left_indent=Inches(.35); st.paragraph_format.first_line_indent=Inches(-.18); st.paragraph_format.space_after=Pt(3); st.paragraph_format.line_spacing=1.15


def add_cover(doc):
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(6)
    r=p.add_run("BUILD MANUAL • VERSION 0.1"); r.font.name="Arial"; r.font.size=Pt(10); r.bold=True; r.font.color.rgb=RGBColor(12,124,134)
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(4)
    r=p.add_run(PROJECT); r.font.name="Arial"; r.font.size=Pt(28); r.bold=True; r.font.color.rgb=RGBColor(16,42,67)
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(16)
    r=p.add_run(SUBTITLE); r.font.name="Arial"; r.font.size=Pt(13); r.font.color.rgb=RGBColor(92,103,112)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_picture_with_alt(p.add_run(),ASSETS/"kit-concept.png","Concept render of six wearable game buttons, one hub, and an open carry case",width=Inches(6.7))
    cap=doc.add_paragraph("Original concept visualization: six wearable controls plus one hub and carry case."); cap.alignment=WD_ALIGN_PARAGRAPH.CENTER; cap.runs[0].italic=True; cap.runs[0].font.size=Pt(8.5); cap.runs[0].font.color.rgb=RGBColor(92,103,112)
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(10)
    r=p.add_run("Prepared as an execution handbook • UAE focus • "+TODAY.strftime("%d %B %Y")); r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=RGBColor(92,103,112)
    doc.add_page_break()


def add_front_matter(doc):
    doc.add_heading("Executive decision",level=1)
    p=doc.add_paragraph(); r=p.add_run("Build seven devices, not six: "); r.bold=True; p.add_run("six wearable ESP32-C3 nodes and one central hub. Use ESP-NOW node-to-hub and a local Wi-Fi web app. Prove reliability and fun offline before custom electronics, apps, cloud, tooling, or certification spend.")
    doc.add_heading("How to use this handbook",level=2)
    for t in ["Read Manual 01 to freeze scope.","Buy from Manual 04 and build in the order in Manual 05.","Use Manuals 06-11 for firmware/software implementation.","Do not run active playtests until Manual 14's bench and safety gates pass.","Use Manuals 15 and 22-23 before commercial import, claims, or sale."]:
        add_bullet(doc,t)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_picture_with_alt(p.add_run(),ASSETS/"gameplay-villa.png","Six adults safely playing a wearable-button game in a cleared villa garden",width=Inches(6.4))
    cap=doc.add_paragraph("Intended experience: social, premium, active, and safe in a cleared private play area."); cap.alignment=WD_ALIGN_PARAGRAPH.CENTER; cap.runs[0].italic=True; cap.runs[0].font.size=Pt(8.5)
    doc.add_heading("Manual index",level=2)
    table=doc.add_table(rows=1,cols=2); table.style="Table Grid"; table.alignment=WD_TABLE_ALIGNMENT.LEFT
    table.cell(0,0).text="No."; table.cell(0,1).text="Manual"
    for num,title,_ in MANUAL_DATA:
        cells=table.add_row().cells; cells[0].text=num; cells[1].text=title
    for cell in table.rows[0].cells: set_cell_shading(cell,"DDEBF2")
    mark_table_header(table)
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; set_cell_margins(cell)
            for run in cell.paragraphs[0].runs: run.font.name="Arial"; run.font.size=Pt(8.5)
    set_table_widths(table,[.7,5.8])
    doc.add_page_break()


def build_docx():
    doc=Document(); style_document(doc)
    sec=doc.sections[0]
    header=sec.header.paragraphs[0]; header.text="PHYSICAL SOCIAL GAMES • BUILD HANDBOOK"; header.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    for r in header.runs: r.font.name="Arial"; r.font.size=Pt(8); r.bold=True; r.font.color.rgb=RGBColor(92,103,112)
    add_page_field(sec.footer.paragraphs[0])
    add_cover(doc); add_front_matter(doc)
    for idx,(num,title,body) in enumerate(MANUAL_DATA):
        p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(3)
        if idx: p.paragraph_format.page_break_before=True
        r=p.add_run(f"MANUAL {num}"); r.font.name="Arial"; r.font.size=Pt(9); r.bold=True; r.font.color.rgb=RGBColor(231,111,81)
        p=doc.add_paragraph(title,style="Heading 1"); p.paragraph_format.space_before=Pt(0)
        if num=="02":
            p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_picture_with_alt(p.add_run(),ASSETS/"architecture.png","Offline-first architecture: six wearables use ESP-NOW to a hub, which serves a local web app",width=Inches(6.55))
            c=doc.add_paragraph("Figure 1. MVP communication path: local, acknowledged, and independent of venue internet."); c.alignment=WD_ALIGN_PARAGRAPH.CENTER; c.runs[0].italic=True; c.runs[0].font.size=Pt(8.5)
        if num=="03":
            p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_picture_with_alt(p.add_run(),ASSETS/"node-blocks.png","Wearable node functional block diagram showing press, ESP32-C3, LED, vibration, buzzer, battery, and charging",width=Inches(6.25))
            c=doc.add_paragraph("Figure 2. Functional node blocks; the hub remains the authority for state and scoring."); c.alignment=WD_ALIGN_PARAGRAPH.CENTER; c.runs[0].italic=True; c.runs[0].font.size=Pt(8.5)
        if num=="12":
            p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_picture_with_alt(p.add_run(),ASSETS/"wearable-exploded.png","Exploded product concept with cap, silicone actuator, light pipe, PCB, battery, haptic and audio parts, shell, gasket, and clip",height=Inches(6.6))
            c=doc.add_paragraph("Figure 3. Concept assembly: cap, silicone actuator, light pipe, PCB, cell, haptic/audio parts, shell, gasket, and clip."); c.alignment=WD_ALIGN_PARAGRAPH.CENTER; c.runs[0].italic=True; c.runs[0].font.size=Pt(8.5)
        if num=="21":
            p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; add_picture_with_alt(p.add_run(),ASSETS/"roadmap.png","Twelve-week roadmap from one-node proof through six-node bench test, wearable alpha, and paid pilot",width=Inches(6.55))
            c=doc.add_paragraph("Figure 4. Twelve-week path from technical proof to a paid evidence-gathering pilot."); c.alignment=WD_ALIGN_PARAGRAPH.CENTER; c.runs[0].italic=True; c.runs[0].font.size=Pt(8.5)
        add_markdown(doc,body)
    source_heading=doc.add_heading("Source register and limitations",level=1); source_heading.paragraph_format.page_break_before=True
    doc.add_paragraph("Current regulatory and platform facts were checked against the official sources below on 19 August 2026. Component prices are planning estimates and must be replaced with live supplier quotations. This handbook is not legal, tax, medical, product-safety, or certification advice.")
    for label,url in [("Espressif ESP32-C3 overview",SOURCES['espressif']),("Nordic nRF52840 product page",SOURCES['nordic']),("TDRA equipment approval service",SOURCES['tdra']),("MoIAT conformity certificates",SOURCES['moiat']),("UAE mainland business steps",SOURCES['business']),("UAE data protection overview",SOURCES['privacy']),("UAE intellectual property",SOURCES['ip']),("Federal Tax Authority VAT registration",SOURCES['vat'])]:
        add_bullet(doc,f"{label}: {url}")
    out=OUT/"Physical-Social-Games-Business-Handbook.docx"; doc.save(out); return out


def main():
    copy_generated_assets()
    diagram_architecture(ASSETS/"architecture.png")
    diagram_node(ASSETS/"node-blocks.png")
    diagram_roadmap(ASSETS/"roadmap.png")
    write_markdown_files()
    out=build_docx()
    print(out)


if __name__=="__main__":
    main()
