# Flai (Voice AI) — source notes

Dashboard: `app.useflai.com/enterprise/findlay-automotive-group`. Tray's login is already active in Chrome, so the Claude browser extension reaches it without any sign-in step — go straight to `app.useflai.com`, **do not** click Login (that lands on a credential form).

Two views in the top bar: **Service** and **Sales**. Everything below is the **Service** view. Sales has not been read yet.

Data file: `data/flai-aug-2026.mjs`. Slides: Segment 07 of `decks/cda-aug-2026.mjs`.

---

## 🔴 Two rules for this source

1. **Different grain from the P&L.** Flai reports **25 locations**; Findlay has **34 rooftops**. The names are Flai's own and do not map 1:1 — Post Falls appears twice ("CDR Post Falls" and "Jeep Post Falls"), and "Lexus Spokane" has no obvious P&L twin. **Never run these through `rooftops.mjs` canonical()**, and never put a Flai number on a slide with a P&L number.

2. **The three money figures are Flai's estimates**, built on industry benchmarks, not Findlay's P&L — the dashboard itself says so. Any slide showing them must say so too. The call, appointment and transfer counts are hard numbers; the dollars are a model.

---

## Group figures — Jul 16 to Aug 15, 2026 (Service)

| | |
|---|---|
| Calls answered | **15,830** |
| Avg duration | 1:51 |
| Appointments scheduled | **3,152** |
| Booking conversion | **66%** (3,152 of 4,763 bookable) |
| Transfer rate | 53% (8,351 calls handed to a person) |
| **No pickup** | **52% — 4,368 of 8,351 transfers rang out** |
| Est. cost savings | $79,150 *(Flai estimate)* |
| Est. revenue impact | $1,134,720 *(Flai estimate)* |
| Est. profit impact | $703,246 *(Flai estimate)* |

**Reconciliation passed.** The 25 location rows sum exactly to the group tiles on calls (15,830), bookable (4,763), appointments (3,152) and transfers (8,351). Per-location booking rate = appointments ÷ bookable, all 25 within 1pt of the reported figure.

### The finding that carries the segment

**7,479 calls (47%) were resolved end to end without a person.** Of the 8,351 handed off, **4,368 were never picked up** — 52% of transfers and **28% of all inbound**.

Given Tray's ruling below, the framing is: the transfer itself is a *success* — the customer wanted a person and Flai got them one. The failure is entirely on our side of the handoff. That is the actionable number in the segment and it is a hard count, not an estimate.

### Booking rate outliers

Best: Nissan 81% · Toyota Prescott 80% (on 2,300 calls — volume *and* rate) · Toyota Henderson and Toyota Flagstaff 73%.

Worst: **Subaru Las Vegas 6%** — 748 calls, 102 bookable, **6 booked**. That is a setup problem, not performance; flag it for the vendor rather than the store. Then Acura 22% (tiny volume, 19 calls), Jeep Post Falls 37%, CDR Post Falls 41%.

### Data-quality flags seen on the dashboard

- **Avg Res. Time is unreliable.** Honda Henderson reads *2131h 51m* (89 days) and Hyundai Prescott *1029h 55m* (43 days). Do not put this column on a slide until Flai explains it.
- Jaguar Land Rover Las Vegas: Resolution 0%, Avg Res. Time blank.
- Resolution % swings hard on low-volume stores (Audi Reno Tahoe 19%, Lexus Spokane 36%, Nissan 40%) — treat as noisy.

---

## Registry

Added to `metrics.mjs` 2026-08-16: `transfer rate` (up) · `calls answered` (up) · `appointments set` (up) · `booking rate` (up) · `resolution %` (up) · `no pickup %` (down).

**Transfer rate ruling — Tray, Aug 2026: UP IS GOOD.** In his words, *"we interpret that as giving the customer what they want."* A transfer is the correct outcome, not an AI failure. Do not re-raise this. The failure mode to watch is `no pickup %`, which is what happens *after* a good transfer.

---

## Built — Segment 07, four slides

1. `section` separator — Voice AI
2. `content` — "Why we are putting AI on the phone". Tray's argument: overflow and after hours as the floor, speed-to-lead applied to the phone, first-answer stores as the ceiling, staff moved onto exceptions. **Deliberately carries no statistics** — the argument slide sets up the proof slides that follow it.
3. `kpis` — group results, four hard-count tiles (calls · resolved without a person · appointments · booking conversion) plus the 4,368 no-pickup finding as the insight band, framed so the transfer reads as a win and the non-pickup as the gap. Estimate caveat in prose.
4. `ranker` — booking rate by location, scored against the 66% group conversion, GROUP inline at its rank.

## Next on this source

- Read the **Sales** view — never pulled. Likely a second pair if Tray wants it.
- Ask Flai about Avg Res. Time and the Subaru Las Vegas configuration.
- Period is a rolling 30 days set by the date picker top-right. If the segment gets rebuilt for a later meeting, restate the window on the slide — the eyebrow carries it now.
- Now that `transfer rate` is registered, a transfer-rate ranker is available if Tray ever wants a third Flai slide. Not built — slide count is a stated priority.
