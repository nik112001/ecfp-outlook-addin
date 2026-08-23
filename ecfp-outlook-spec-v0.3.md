# eCFP — Email Carbon Footprint Plug-in for Outlook
## Product & Technical Specification (POC)

**Version:** 0.3 (aligned to eCFP Idea Proposal v250726 + AgainstData research)
**Date:** August 3, 2026
**Status:** Proposal for POC build
**Relationship to eCFP proposal:** This spec defines the **endpoint-side POC** described in §3.1 of the eCFP Idea Proposal, implemented as an Outlook add-in. Cloud-side methodology (§3.2 of the proposal) remains an open item and is out of POC scope.

---

## 1. Overview

eCFP is an Outlook add-in that measures, reports, and encourages reduction of email carbon footprint. The POC delivers the endpoint experience: a per-email footprint on read, a live estimate while composing (the "CFF" — current-mail footprint), cumulative personal eCFP tracking with configurable Green/Amber/Red thresholds, a stored-mail footprint view with cleanup nudges, and a lightweight badge system.

### 1.1 Goals

- Prove the endpoint measurement methodology from the eCFP proposal on a real mailbox.
- Make footprint visible at the moment of reading and sending, and cumulatively over time.
- Demonstrate the engagement loop (thresholds → badges) in single-user form.
- Zero backend; all computation client-side.

### 1.2 Non-Goals (POC)

- Cloud-side component, department/org mapping, org token allocation, admin reporting (proposal §3.2, §4.2) — post-POC.
- Windows/Android/iOS native endpoint clients (proposal §3.3) — the Outlook add-in surface covers web + new desktop from one codebase; native clients are post-POC.
- OEM licensing / market mechanics (proposal §6).
- Gmail/other providers (calc engine remains portable).
- Per-hop network measurement (proposal factor #4) — not observable client-side; documented as excluded.
- Precise device power measurement (proposal factor #7) — represented via published embodied/usage coefficients, not live telemetry.

### 1.3 POC Account Scope

Single personal/dev Microsoft account, sideloaded manifest, no admin consent.

---

## 2. Target Experience

Four surfaces, one calculation engine.

### Phase 1 — Read View
Opening an email shows its footprint (e.g., "🌱 4.2 g CO2e") in a pinnable task pane, with breakdown by size, attachments, recipients, and intranet/internet mix. Optional InfoBar one-liner.

### Phase 2 — Compose-Time CFF
Live current-mail footprint in compose, recalculated on recipient/attachment changes and debounced body polling. Color state against thresholds; red state suggests remediation (OneDrive link instead of attachment, trim CC list). Never blocks send.

### Phase 3 — Cumulative eCFP Dashboard
Per the proposal's dual-display requirement, the dashboard shows:
- **Cumulative eCFP** — month-to-date and year-to-date, color-coded against configurable thresholds.
- **Stored-mail footprint** — estimated ongoing storage cost of the mailbox (message count × size × storage coefficient), with a "cleanup opportunity" list (oldest/largest mail, e.g., items > 1 year old or > 5 MB) reflecting the research finding that long-term retention outweighs the one-time cost of deletion.
- Trend chart, top contributors, real-world equivalencies.

### Phase 4 — Engagement Layer (POC-lite)
- Local eCFP token balance: user (or demo config) sets a monthly token budget; each email's footprint draws it down; dashboard and badge tier move together off the same balance, as specified in the proposal.
- Badge tiers (e.g., Platinum / Gold / Silver) computed from threshold performance.
- Cleanup actions (deleting flagged aged mail) credit the balance.
- Org leaderboards and reward redemption require the cloud side — post-POC.

---

## 3. Architecture

### 3.1 Components

```
┌────────────────────────────────────────────────────┐
│ Outlook (web / new Windows / Mac desktop)          │
│  ┌──────────────────────────────────────────────┐  │
│  │ eCFP Add-in (Office.js, static web app)      │  │
│  │  Read pane │ Compose CFF │ Dashboard │ Badges│  │
│  │            └──────┬──────┘                   │  │
│  │         ┌─────────▼─────────┐                │  │
│  │         │ Calc Engine (TS)  │                │  │
│  │         │ + token ledger    │                │  │
│  │         │ + IndexedDB cache │                │  │
│  │         └─────────┬─────────┘                │  │
│  └───────────────────┼──────────────────────────┘  │
└──────────────────────┼─────────────────────────────┘
                       │ HTTPS
          ┌────────────▼──────────────┐
          │ Microsoft Graph API       │
          │ Mail.ReadBasic (metadata) │
          └───────────────────────────┘
```

Add-in hosted as static files (localhost dev; GitHub Pages/Azure Static later). No application backend. The calc engine, thresholds, and token ledger are isolated modules so the future cloud side can consume the same logic.

### 3.2 Key Design Decisions

| Decision | Choice | Rationale |
|---|---|---|
| Endpoint form factor | Office.js add-in (task pane + InfoBar) | Stable official surface; one codebase for web + new desktop; fulfills proposal's endpoint role without native clients |
| Read-mode data | `Office.context.mailbox.item` + Graph size lookup | Instant provisional estimate, authoritative size on Graph response |
| Intranet/internet split (proposal factor #3) | Recipient domain vs. user's own domain | Cheap, deterministic proxy; internal-only mail gets reduced transmission coefficient |
| Storage footprint (proposal factor #5) | Graph mailbox scan: per-message size × age × storage coefficient | Enables cleanup nudges; aligns with retention research |
| Auth | MSAL.js, `Mail.ReadBasic` delegated scope | Metadata only; no admin consent for single account |
| Gamification state | Local token ledger (IndexedDB) | Proposal-faithful single-user loop; cloud sync later |
| Calculation location | Fully client-side | No backend, no privacy liability |

### 3.3 Permissions & Scopes

- Manifest: `read item`.
- Graph delegated: `Mail.ReadBasic` only (metadata incl. size; no bodies).

### 3.4 Data Flows

**Read:** `ItemChanged` → provisional estimate from item recipients/attachments → Graph size refinement → cache → render.
**Compose:** `RecipientsChanged` / `AttachmentsChanged` events + 2 s debounced `body.getAsync` size polling → CFF update.
**Dashboard/storage:** Graph `/me/messages` paged with `$select` minimal fields → per-message send/receive footprint + storage footprint → aggregate, token ledger update, cleanup-candidate list. Incremental sync after first run; honor 429/Retry-After.

---

## 4. Carbon Estimation Model

### 4.1 Principles

- Transparent: coefficients cited in an in-app methodology page.
- Conservative: calibrated to published per-email figures, uncertainty stated.
- Honest: factors we cannot observe (network hops, live device power) are represented by published averages and documented as such — never invented per-message precision.

### 4.2 Proposal Factor Coverage

| # | Proposal factor | POC treatment |
|---|---|---|
| 1 | Attachment size | Direct input (S_MB) |
| 2 | Recipient count | R_factor multiplier |
| 3 | Intranet vs internet recipients | Domain-match discount on transmission term |
| 4 | Network hops | Excluded (unobservable); folded into published averages |
| 5 | Days stored | Storage footprint term (dashboard) |
| 6 | Server location | Excluded from number; roadmap (region carbon intensity) |
| 7 | Device power/embodied | Folded into E_base per published methodology |

### 4.3 Send/Receive Formula (v2)

```
gCO2e_msg = (E_base + E_data × S_MB × I_mix) × R_factor
```

- **E_base** = 0.3 g — short laptop email baseline (device + network + server, incl. amortized embodied carbon)
- **E_data** = 11 g/MB — transmission/processing coefficient
- **S_MB** = message size incl. attachments
- **I_mix** = 1.0 for internet recipients; 0.6 when all recipients share the sender's domain (intranet discount — infrastructure path is shorter; coefficient to be tuned in M1)
- **R_factor** = 1 + 0.25 × (N_recipients − 1), capped at 3.0

### 4.4 Storage Formula (new in v0.3)

```
gCO2e_storage = S_MB × D_years × E_store
```

- **E_store** = 10 g per MB per year (placeholder; to be derived in M1 from data-center energy-per-GB-stored literature)
- Applied per message in the mailbox scan; dashboard shows total ongoing storage footprint and projected annual savings from deleting flagged mail.

### 4.5 Calibration Targets (updated from AgainstData research)

| Email type | Published estimate | Model target |
|---|---|---|
| Filtered spam | ~0.03 g | (not modeled — never reaches user) |
| Short reply, laptop | ~0.3 g | ~0.3 g |
| Short reply, phone | ~0.2 g | ~0.3 g (device split is v2) |
| Long plain-text email | ~17 g | flag: pure text rarely exceeds ~0.1 MB → model gives ~1.4 g; the 17 g figure is dominated by device *time* (write/read minutes), not data size. POC displays a footnote; time-based term is a v2 candidate |
| Image/attachment email | up to ~50 g | ~11–50 g depending on size × recipients |
| Office worker annual (126 msg/day) | ~184 kg/yr | dashboard sanity check at M3 |

> The known tension in the literature — data-size-driven vs. device-time-driven estimates — is documented on the methodology page. POC v1 is size-driven with a cited caveat; a compose/read time term is a tracked v2 item.

### 4.6 Equivalencies

- 1 km driven (avg gasoline car) ≈ 170 g CO2e
- Long email ≈ driving ~100 m; image-heavy email ≈ driving ~300 m
- 1 smartphone charge ≈ 8 g CO2e

---

## 5. UI Specification

### 5.1 Read Pane — per-message figure, color state, breakdown bars (base / data / recipients / intranet mix), methodology link.
### 5.2 Compose — live CFF figure + color state; red-state remediation line. Never blocks send.
### 5.3 Dashboard — cumulative eCFP (MTD/YTD) with threshold color; token balance gauge; badge tier; storage footprint card with top cleanup candidates ("Delete 12 flagged emails → save ~340 g/yr"); trend chart; equivalencies.
### 5.4 Settings — configurable G/A/R thresholds (proposal requirement), monthly token budget, clear data, sign out, methodology page.
### 5.5 Onboarding — explainer (metadata only, local processing) → instant read/compose function → Graph consent on first dashboard open.

This spec satisfies the proposal's Figure 1 placeholder: the architecture diagram (§3.1) and the UI described here (eCFP + CFF with G/A/R coding) can be exported as that figure.

---

## 6. Privacy & Security

- No backend; all processing/storage local.
- `Mail.ReadBasic` — bodies never readable via API.
- Token ledger and badge state stored locally only.
- Clear-data control wipes all local state; MSAL tokens revocable from Microsoft account.

---

## 7. Tech Stack

Office.js (Mailbox 1.10+), unified JSON manifest, TypeScript, React + Fluent UI, Recharts, MSAL.js 2.x, Microsoft Graph JS SDK, IndexedDB (idb), webpack dev server, Vitest for calc-engine/ledger unit tests.

---

## 8. Build Plan & Milestones

| Milestone | Scope | Est. effort |
|---|---|---|
| M0 — Skeleton | Scaffold, sideload, panes render read + compose | 2–4 days |
| M1 — Read + Engine | Calc engine v2 (incl. I_mix), coefficient research & tuning, read pane, Graph size lookup, caching | 6–8 days |
| M2 — Compose CFF | Compose events, body polling, CFF + thresholds | 4–6 days |
| M3 — Dashboard + Storage | MSAL, Graph sync, cumulative eCFP, storage footprint + cleanup list, equivalencies | 6–9 days |
| M4 — Engagement + Polish | Token ledger, badges, configurable thresholds, methodology page, onboarding, demo script | 4–6 days |

**Total: ~5–7 weeks** (one developer).

### 8.1 Success Criteria

- Read pane updates < 1 s on message switch; CFF updates < 2 s on compose changes.
- Calc engine unit tests pass calibration table §4.5.
- 30-day dashboard sync (~1,000 msgs) < 45 s cold / < 2 s cached; annualized figure within order of magnitude of the 184 kg/yr benchmark for a comparable-volume mailbox.
- Storage view lists cleanup candidates with projected savings.
- Badge tier changes when token balance crosses thresholds.
- Network inspection confirms no message body content leaves the client.

---

## 9. Risks & Mitigations

| Risk | Impact | Mitigation |
|---|---|---|
| Size-driven vs time-driven estimate criticism | Medium | Methodology page documents both schools with citations; time term tracked as v2 |
| Storage coefficient weakly sourced | Medium | M1 literature task; display as range until firmed |
| Full-mailbox scan cost for storage view | Medium | Cap initial scan (e.g., 2,000 messages / 12 months), incremental extension |
| Office.js size gaps in read mode | Medium | Graph authoritative; body-length heuristic fallback |
| Gamification feels arbitrary single-user | Low | Frame as demo of the org loop; thresholds/budget configurable |
| Graph throttling | Low | `$select`, `$batch`, incremental sync, Retry-After |
| Future classic Outlook need | Medium | Office.js calls behind adapter interface |

---

## 10. Post-POC Roadmap (mapped to eCFP proposal)

1. **Cloud side (proposal §3.2):** org aggregation service, department/individual mapping, admin dashboards, token allocation by role, leaderboards & recognition (proposal §4.2).
2. **Native endpoint clients (proposal §3.3):** Windows/Android/iOS beyond the add-in surface.
3. **Server-location factor (#6):** data-center region carbon intensity.
4. **Time-based estimate term** for long-email accuracy; device-class split (laptop vs phone).
5. **Reward mechanism definition** (proposal §4.1 open item) once cloud side exists.
6. Gmail port; OEM/enterprise packaging per proposal §6.

---

## 11. TODO (deferred, revisit post-POC)

- [ ] **Time-based estimate term** — add compose/read duration to the model so long plain-text emails align with the ~17 g published figure (currently size-driven only; see §4.5 flag).
- [ ] **Device-class split** — laptop vs phone coefficients (0.3 g vs 0.2 g baseline).
- [ ] **Reconcile per-email range** — proposal doc says 0.03–26 g; sources support up to ~50 g. Update proposal before external circulation.
- [ ] **Firm up E_store coefficient** — replace 10 g/MB/yr placeholder with sourced data-center storage figure (M1 literature task).
- [ ] **Cloud-side methodology** — proposal §3.2 open item; org mapping, token allocation, leaderboards.
- [ ] **Reward mechanism definition** — proposal §4.1 open item (redeemable credits vs recognition tickets).

---

## Appendix A — Sources

1. eCFP Idea Proposal v250726 (internal) — factor model, engagement design, dual-display requirement.
2. AgainstData, "Email Carbon Footprint: Facts, Impact, and Cleanup Tips" (updated Jun 2025) — per-email figures (0.3 g short / 0.03 g spam / ~17 g long / ~50 g attachment), 126 emails/day and ~184 kg/yr office-worker benchmark, storage-vs-deletion energy argument. Note: the article's DIY kWh example (0.3 kg per email) is internally inconsistent with its own per-email figures and is excluded.
3. Berners-Lee, M. *How Bad Are Bananas?* (2020 ed.) — underlying per-email methodology.
4. The Shift Project, *Lean ICT* (2019); Carbon Literacy Project — supporting coefficients.
5. Data-center storage energy literature — to be selected in M1 for E_store.

*Discrepancy log: proposal cites 0.03–26 g per-email range; AgainstData and Berners-Lee support up to ~50 g for attachment-heavy mail. Spec uses 50 g as upper calibration bound — reconcile in the proposal doc before external circulation.*
