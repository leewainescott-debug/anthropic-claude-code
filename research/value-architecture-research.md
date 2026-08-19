# Value Architecture — Research Foundation
**Purpose:** the evidence base for a Value Architecture capability and GTM. This is the research layer under the deck — what clients' problems actually are, what the numbers say, how the market sells against them, and where the whitespace is.
**Compiled:** 19 August 2026 · Sources prioritise 2025–2026 publications. Every stat carries a confidence flag: **[A]** primary-sourced and recent, **[B]** primary but older/classic (use with a date caveat), **[C]** secondary/aggregator (verify before putting in front of a client).

---

## 1. The argument in one page

Organisations don't have an investment problem. They have a **realisation** problem.

The chain of evidence, in plain English:

1. **Money goes in.** Software spend alone grows ~15% in 2026 to $1.4T+ — but most of the growth is price increases and AI, not new value. [A — Gartner]
2. **The business case gets the money, then dies.** ~25% of transformation value is lost at the target-setting/business-case stage — before delivery even starts. [B — McKinsey "Losing from Day One"] Business cases are approval artefacts: signed once, never referenced again.
3. **Delivery "succeeds", value doesn't.** Of ~16,000 projects in Flyvbjerg's database, only 8.5% hit cost and time targets — and only **0.5% hit cost, time AND benefits**. [B — Flyvbjerg, *How Big Things Get Done*, 2023] "On time, on budget" is not the same as value delivered — and the budget was never really the issue.
4. **Nobody is left holding the benefit.** Fewer than 40% of organisations have a formal, repeatable process for tracking whether projects delivered projected benefits after go-live. [B — PMI; verify vintage] Only 39% of organisations mostly/always deliver full project benefits. [A — Wellingtone State of PM 2025/26] The project team disbands at go-live; the benefit owner was never really named. "Post-project amnesia."
5. **Governance tracks spend, not value.** The iron triangle (cost, time, quality/scope) is what gets reported; contribution to business KPIs isn't. [A — APM] KPMG found 83% of M&A deals failed to boost shareholder returns *while integration scorecards showed green* — dashboards measure activity, not value. [B — KPMG via L.E.K.]
6. **2026 is the year the board started asking.** 66% of boards now condition further AI funding on proof of return; 43% of finance leaders are being asked for an ROI number they can't produce. [A — RGP/CFO.com] Board-level AI value reporting is practiced by just 4% of organisations today but expected to be standard by end of 2026. [A — Deloitte] "Boards will stop counting tokens and pilots and start counting dollars" (James Brundage, EY).

**The consequence for the offering:** the client problem is not "we need a strategy." It's *"we keep approving investments whose value we can't see, prove, or steer."* Value Architecture is the discipline that fixes that — it opens every investment conversation with one question: **what value does this drive for the business as a whole, and how do we set ourselves up to actually realise it?** That is a problem clients already feel (board pressure, dead business cases, platforms they pay for but don't use) — not a solution looking for a problem.

---

## 2. What clients' problems actually are

Framed the way a client would say them — each with the evidence behind it.

### "We can't prove our investments paid off."
- 89% of operations leaders say their tech investments haven't fully delivered expected results — while 85% believe they're ahead of competitors on digital. [A — PwC 2026 Digital Trends in Operations, n=767]
- 88% of executives say achieving measurable value from new technology is a challenge. [B — PwC Pulse, 2023 base]
- Only 29% of executives can measure AI ROI confidently, even though 79% see productivity gains — the translation from "feels productive" to P&L is the sticking point. [A — IBM IBV 2025/26]

### "Our business cases are written to get funding, not to run the investment."
- ~25% of value loss happens at target-setting — inflated or vague benefits baked in on day one. [B — McKinsey]
- Optimism bias is so systemic that the UK Treasury Green Book *mandates* uplifts of up to +44% on cost and +20% on schedule for standard business cases — government formally assumes business cases are inflated. [B — HM Treasury]
- Benefit shortfalls of 40–50% are common and up to 75% "not uncommon"; the odds that benefits overrun ever exceed a cost overrun are only ~20%. Costs reliably run away; benefits reliably don't show up. [B — Flyvbjerg]
- The failure cascade for static business cases: unclear targets → implied (not assigned) ownership → no early-warning signals → manual tracking that fades → post-project amnesia. [A — practitioner synthesis, Kiplot 2025/26]

### "We deliver projects, not outcomes."
- Bain's 24,000-initiative database: only ~12% of transformations achieve their original ambition. Transformations with a dedicated Chief Transformation Officer capture **24% more planned value** — dedicated value ownership is the measured differentiator. [A — Bain 2024]
- Even *successful* transformations capture only 67% of maximum possible financial benefit; unsuccessful ones capture 37%. [B — McKinsey]
- Only 39% of organisations mostly/always deliver full project benefits (and benefits management is rated the hardest discipline to embed). [A — Wellingtone 2025/26]
- Note: the famous "70% of transformations fail" stat is a **myth** — it traces to an unscientific 1993 estimate (Hammer & Champy) recycled via Kotter. Don't use it; using Bain's 88%/12% instead (and knowing why) is itself a credibility play in the room.

### "We're paying for platforms we don't fully use."
- Average enterprise: 305 SaaS apps, $55M annual SaaS spend, **46% of licenses unused** — ~$19.8M wasted per organisation per year. [A — Zylo 2025/26 SaaS Management Index; triangulated by Productiv at ~51%]
- Cloud waste rose to **29% of spend in 2026 — the first increase in five years**, driven by AI workloads. [A — Flexera 2026 State of the Cloud, n=753]
- Idle/oversized compute typically eats **30–40%+ of Snowflake and Databricks spend** — consumption pricing means unrealised value is literally money burning on warehouses nobody shut down. [C — cost-optimisation vendor consensus; directionally strong, verify per client]
- Only 42% of companies fully achieve expected cloud value — up just 5 points since 2020. Migration ≠ value capture. [A — Accenture Cloud Outcomes, Jan 2025]
- Salesforce: only 33% of AI initiatives on the platform are meeting expected ROI; only 26% of customers say most of their customer data actually sits in Salesforce. [A — IBM IBV State of Salesforce 2025/26, n=1,222]

### "We bought AI. Where's the value?"
- 75% of leaders rank AI a top-3 priority; only 25% say they're realising significant value. [A — BCG 2025]
- Only ~6% of companies attribute >5% of EBIT to AI; 39% report any EBIT impact at all. [A — McKinsey State of AI, Nov 2025, n=1,993]
- The top 20% of companies capture **74% of all the economic value AI creates** (7.2x the gains of peers). Value concentrates with the disciplined few — this is a power law, not a lottery. [A — PwC AI Performance Study, Apr 2026, n=1,217]
- Gartner predicts **>40% of agentic AI projects will be cancelled by end-2027** — unclear business value, escalating cost, inadequate risk controls. [A — Gartner, standing prediction]
- Why it fails: BCG's 10-20-70 rule — 10% of AI value is the algorithm, 20% tech and data, **70% people and process redesign**. Yet ~80% of companies layer AI onto unchanged workflows. [A — BCG; McKinsey]
- New 2026 concept worth using: the **"verification tax"** — >26% of expected AI productivity gains are lost to humans checking and explaining AI outputs. [A — Sage/IDC via Diginomica 2026]

### "Our deal didn't deliver the synergies we promised the board."
- ~70% of merging companies overestimate synergies at announcement; only ~30% of acquirers hit synergy targets. [A — Bain, 31,000-company analysis]
- Cost synergies capture 70–85% of announced value within 18 months; **revenue synergies capture only 25–35%**, and take ~5 years vs ~2 for cost. [A/B — McKinsey/Bain/BCG consensus]
- Only 43% of companies formally track synergy capture at all — despite synergies being the deal's entire justification. [C — verify]
- PE's problem is sharper in 2026: "**12 is the new 5**" — firms now need ~12% annual EBITDA growth (vs 5% historically) to hit 2.5x MOIC, because multiple expansion collapsed from 59% of returns to under 15%. Holds average 6.6–7 years; 32,000 unsold companies worth $3.8T are waiting for operational value creation to justify the exit. [A — Bain Global PE Report 2026; McKinsey]

---

## 3. Why now — the 2026 market moment

Four forces make this the right year to launch a value capability:

1. **The AI ROI reckoning.** 2026 is the "show me the money" year. 66% of boards condition AI funding on proof of return; 71% of IT leaders believed they had until mid-2026 to prove AI value or face budget/job fallout; only 7% of leaders have "established" AI ROI. [A — CloudZero/Kyriba, cio.com, KPMG Pulse Q2 2026] Gartner now recommends finance sign-off *before* material agentic AI investment — the ROI gate is being institutionalised, and someone has to build it.
2. **Value reporting is about to be table stakes.** Board-level AI value reporting: 4% today → expected standard for large enterprises by end of 2026. [A — Deloitte] Clients who can't report value are about to look exposed at board level.
3. **Waste is rising, not falling.** Cloud waste up for the first time in five years (29%); SaaS waste increasing per 35% of ITAM leaders; AI pricing models making cost visibility worse (only 26% have real-time AI cost visibility — and those who do are 5x more likely to have established ROI). [A — Flexera, KPMG]
4. **PE has run out of other levers.** With multiple expansion gone and exits frozen, operational value creation is the only path to returns — and 58% of PE firms now deploy resources in the first 100 days of ownership, double the prior year. [A — A&M Value Creation Report 2026]

**Measurement itself is the differentiator.** KPMG's Q2 2026 finding is the cleanest version: organisations with full cost visibility are 5x more likely to have established ROI; where the CEO is accountable for AI outcomes, 57% realise meaningful value vs 21% where nobody is. Accountability + visibility — exactly what a value architecture installs — are the two measured predictors of getting value.

---

## 4. How the market sells "value" (competitive landscape)

### The universal product shape
Every serious player sells the same three-layer structure:
1. **Diagnostic / assessment** — baseline the value position, benchmark vs peers (McKinsey Value Assurance 360° Assessment, SAP CVA, Snowflake value assessments).
2. **A value artefact** — value tree / value bridge / value map decomposing a financial North Star into initiatives and KPIs (Bain's "value bridge": e.g. "+$200M run-rate EBITDA, +10 NPS, −25% cycle time"; Deloitte's Enterprise Value Map).
3. **An embedded office** — the anchor SKU: an ongoing governance function that tracks forecast-vs-actual and prevents "value leakage" (Bain Results Delivery Office, PwC Value Realization Managed Services, Capgemini VRO, McKinsey Transformation Office, Accenture Transformation Office powered by Momentum).

Shared enemy narrative: **"value leakage."** Shared conviction: the analytical case is rarely the failure point — adoption, ownership and behaviour change are.

### Firm-by-firm in one line each
| Firm | Offering | The tell |
|---|---|---|
| McKinsey | Value Assurance; Transformation Office | ~25% of global fees now outcomes-based; "value assurance" = finance-embedded tracking of realised vs target |
| BCG | TURN; Performance & Value Acceleration; AI-Powered Transformation Office (2026) | Value = TSR; embedded senior leadership on site; "AI does not create value by default" |
| Bain | Results Delivery® / RDO | Proof stat: best change managers deliver 86%+ of promised results vs 43% for worst (300+ programmes) |
| Deloitte | **Uses the term "Value Architecture"** (capital-allocation context); Enterprise Value Map™; Vision to Value (SAP) | "Impact verified as booked monthly" with Finance; value tree + OKRs, "outcomes rather than activity" |
| PwC | **Value Realization Managed Services — a managed VRO** | The most literal "buy a VRO as a service" offer on the market |
| EY / EY-Parthenon | Long-Term Value (4 pillars); PE value creation | Experimenting with outcomes-based billing on AI engagements |
| Accenture | Value & Benefits Realisation Services; Transformation Office + **Momentum platform** | "Dedicated focus on the 'why', not the 'how', of technology"; hires "Value Realization Specialists" |
| KPMG | Elevate | 300+ value-based analytical modules; strategy + implementation |
| Partners in Performance (now Accenture) | Implementation + value delivery | Contingent fees — money where mouth is; "wiring" value into daily behaviour |
| Kyndryl / Capgemini / Fujitsu | Named VROs | Kyndryl: VROs as "air-traffic controllers of transformation… cutting through the illusion of success defined by activity alone" |

**Naming note:** Deloitte already uses "value architecture" (capital-allocation framing). Not a blocker — but the offering name will need Slalom-distinct framing, or lean into a variant ("Value Architecture & Realisation", or anchor on the VRO/outcome language).

### The competitor nobody talks about: the vendor's own value team
SAP (Value Engineering since 2004), ServiceNow (Now Value + Impact), Salesforce (Business Value Services), Snowflake (Value Engineering) all run in-house value teams. **But they are pre/post-sale motivated and single-vendor** — they exist to justify the vendor's deal, and hand off once it closes. Databricks has no publicly named equivalent (gap). Any Slalom value pitch must position *with* these teams (they bring funding and appetite) while owning what they don't: delivery-phase and post-go-live realisation, cross-platform.

### The whitespace (where Slalom wins)
1. **Value plumbing, not just value architecture.** MBB value offices run on manually reported programme status. Slalom is *inside* the AWS/Salesforce/Snowflake/Databricks build and can wire value tracking directly to platform telemetry — consumption data, adoption data, pipeline data. Live instrumented value, not a reporting layer. Nobody researched does this well.
2. **The vendor VE handoff gap.** Vendor value teams build the pre-sale case and disappear; own the realisation phase they abandon.
3. **Cross-platform.** Every vendor VE team is single-vendor; MBB are vendor-agnostic but not vendor-fluent. Value realisation across AWS + Salesforce + Snowflake + Databricks as one architecture is unowned.
4. **Mid-market.** PwC/Capgemini/McKinsey value offices are priced for mega-programmes. A lightweight, fast-standing, affordable value office is unserved.
5. **Public fixed-price entry.** No competitor publishes a priced, fixed-duration platform value assessment — everything is gated behind a sales conversation. A public "10 days, fixed fee, here's what you get" offer (Slalom already does exactly this with its CRM Health Check on AppSource) is a differentiator in itself.
6. **Verifiable outcome pricing.** Everyone wants outcome-based fees; the blocker is proving value defensibly. Platform telemetry access makes Slalom structurally better placed to underwrite fee-at-risk than a strategy firm.

---

## 5. What good looks like — the mechanics (the "practical, not theoretical" layer)

### Defining value, truly
A defensible working definition:

> **Value architecture is the connective tissue between a strategic ambition and the day-to-day operating numbers that prove it happened.** It names the outcome being bought, decomposes it into the operational drivers and leading indicators that produce it, assigns a named business owner to each, baselines it against real operating data, and keeps the whole structure alive — re-measured, re-owned, re-governed on a fixed cadence — from business-case signature until long after the delivery team has disbanded. It is not a document; it is a standing management discipline.

How value is defined beyond ROI (pick per client, don't preach all of them):
- **Hard vs soft**, or more rigorously the UK public-sector taxonomy: **cashable** (financial), **non-cashable tangible** (quantifiable, hard to monetise), **non-cashable intangible** (identifiable, not quantifiable).
- **Value driver trees** (the MBB standard): board outcome → mathematical drivers → operational leading indicators. "Leading metrics answer *what should we do*; lagging metrics answer *what did that produce*."
- Multi-stakeholder frames when the client needs one: EY Long-Term Value (customer/people/societal/financial), Accenture 360° Value, UK National TOMs for public sector (every social outcome carries a financial proxy value).
- The client-usable one-liner: **value = a business KPI somebody owns, with a baseline, a target, a date, and a named owner in the business.** Everything else is aspiration.

### The established discipline to stand on (credibility anchors)
- **PMI Benefits Realization Management**: Identify → Execute → Sustain.
- **UK Green Book Five Case Model** (Strategic/Economic/Commercial/Financial/Management) — the management case is literally "how will benefits actually be delivered and tracked."
- **Cranfield Benefits Dependency Network** (Ward): investment objectives → business benefits → business changes → enabling changes → technology enablers. The point clients miss: benefits come from *business change*, technology only enables.
- Core artefacts practitioners actually use: **benefits register**, **benefit profiles** (owner in the business, baseline, target, date, measurement method, dis-benefits), **value driver tree**, **benefits realisation plan**.

### The ten mechanisms that make value real
1. Every benefit has **one named owner in the operating business** — never the project team, because the project team disbands and the owner can't.
2. Every benefit is **baselined before the change starts**, against real operating data.
3. A **value driver tree** decomposes the headline number into leading indicators a manager can act on weekly.
4. Value metrics live **inside the dashboards leaders already review** — never a separate benefits spreadsheet (it decays).
5. The business case is **re-baselined at every stage gate** — gates re-validate benefits, not just budget.
6. **Funding releases in tranches** tied to evidenced milestones (venture-style), so capital can be redirected before it's all spent.
7. One senior accountable owner holds **explicit kill/pause authority** (UK SRO model).
8. A **post-implementation review** weeks-to-months after go-live checks realised benefit vs original target.
9. **Smart kills get celebrated**, messengers protected — else sunk-cost bias keeps dead initiatives alive.
10. A fixed cadence: weekly delivery forum, **monthly value review with a Go/Hold/Kill/Accelerate decision**, quarterly executive re-litigation.

### The five anti-patterns to name in front of clients
1. **Benefit-tracking theatre** — dashboards of vanity metrics that spike charts but track nothing anyone signed up to own. (KPMG's M&A finding is the killer illustration: 83% of deals destroyed value while scorecards showed green.)
2. **Post-project amnesia** — team disbands at go-live, nobody left holding the benefit.
3. **The wallpaper business case** — signed at approval, irrelevant by month two. (PE version: the VCP "becomes a wallpaper deck nobody updates after Q2 of Year 1.")
4. **Ownership by implication** — a benefit on a slide with no named owner, no baseline, no consequence.
5. **VRO as police, not coach** — enforcement-flavoured value offices get gamed or ignored; the office exists to help owners hit the number.

### Standing one up (the "how we'd get started")
- A **minimum viable VRO stands up in ~3–6 months**; first 90 days should already report early wins. Start with a maturity assessment of existing governance, keep it ruthlessly simple, track business value not activity, and secure visible executive sponsorship.
- **Where it reports matters:** the strongest models anchor to the **CFO/Finance** ("impact verified as booked monthly" — Deloitte EVD; McKinsey's value assurance is finance-embedded). Finance sign-off is what separates real value from claimed value.
- Build–Operate–Transfer is the natural consulting shape: build it with the client, run it together, hand it over. The exit is the point — a VRO the client can run is a reference; a dependency is a liability.

---

## 6. The four offering theses, mapped to evidence

### Offering 1 — Value Architecture core: reframed business cases + value realisation setup
**Client problem:** business cases get money then die; nobody can prove value; the board is now asking.
**What they buy (practically):**
- *Entry:* a *value case* (not a business case) for one flagship investment — driver tree, baselines, named owners, tracking design. Small, fixed-fee, 3–6 weeks.
- *Follow-on:* value tracking mechanisms, reporting cadence + governance ("living business case" install); or a full VRO stand-up (3–6 months, Build–Operate–Transfer).
**Evidence spine:** Flyvbjerg 0.5%; McKinsey 25%-lost-at-business-case + 67%/37% capture; Bain CTO +24%; Wellingtone 39%; PMI <40% tracking; APM iron-triangle critique.
**Who cares:** CFO first (value verified in the P&L), then transformation lead/sponsor. Deloitte's "booked monthly with Finance" is the standard to match.
**Positioning line to beat:** Kyndryl's "the illusion of success defined by activity alone."

### Offering 2 — AI value realisation
**Client problem:** AI spend under board scrutiny; can't connect spend to outcomes; pilots don't scale; agentic hype meets a 40%-cancellation forecast.
**What they buy:**
- *Entry:* AI value scan — inventory AI spend/initiatives against a value driver tree; kill/fund/fix portfolio view; cost-visibility fix (the KPMG 5x lever). 2–4 weeks.
- *Follow-on:* the ROI gate Gartner says to build (finance sign-off pre-investment); value tracking per use case; workflow-redesign engagements (the BCG "70%").
**Evidence spine:** BCG 75/25 + 10-20-70; McKinsey 6% EBIT; PwC top-20%-capture-74%; KPMG visibility 5x + CEO accountability; Gartner 40% agentic cancellations; verification tax 26%.
**Watch-outs:** use MIT's "95% fail" only with caveats (small sample, not peer-reviewed, contradicted by its own shadow-AI finding); never present Gartner's 30% abandonment *prediction* as fact.
**Whitespace:** no firm has branded an "AI Value Office"/"AI value assurance" product by name. Own the naming.
**Slalom asset:** the existing **AI Value Platform** (TCO/ROI forecasting) — the tool layer the market now demands.

### Offering 3 — Platform value unlock (with AWS, Snowflake, Databricks, Salesforce)
**Client problem:** paying for platforms they use a fraction of; consumption bills rising; told the fix is an upgrade they can't afford.
**What they buy:**
- *Entry:* **Platform Value Assessment** — modules/capabilities used vs owned, consumption efficiency, adoption, trapped-value quantification → a value case + roadmap to unlock value *without major new spend*. Fixed duration (10–15 days — extend the proven CRM Health Check pattern), and **often free or near-free via partner funding**.
- *Follow-on:* the unlock roadmap itself (funded via MAP/LiftOff/consumption incentives), FinOps-style value governance, ongoing value tracking.
**Evidence spine:** Zylo $19.8M/46%; Flexera 29% cloud waste; IBM 33% Salesforce AI ROI; Accenture 42% cloud value; 30–40% idle compute; Gartner/RAND ~80% analytics/AI initiatives fail to deliver outcomes.
**The easy button (partner funding, verify terms current):**
- AWS: MAP assessment funding (~5% of projected ARR; 15–25% ARR credits + partner cash at scale); PoC funding to $25K; Well-Architected reviews free + $5K credits/workload post-remediation.
- Snowflake: Services Funds Program (3–5%); LiftOff funded foundation work.
- Databricks: Feb 2026 Partner Well-Architected Framework explicitly ties **assessments** to tiering/GTM incentives; joint GTM funding.
- Salesforce: FY27 program introduces **consumption funds** rewarding adoption — the natural funding line for "unlock what you already bought."
- Sequencing rule: funded vehicles often can't stack but can chain (PoC → MAP).
**Strategic frame:** vendors *want* this — utilisation drives their renewals and consumption revenue. Slalom's assessment is aligned with the vendor's growth motion but sits on the client's side of the value question. Health checks are commoditised; **quantified trapped value + funded delivery** is the differentiation.

### Offering 4 — Deal & speed-to-value (PMI / PE lens)
**Client problem:** synergies overpromised, undertracked, slow; PE needs 12% EBITDA growth with nowhere to hide; IT integration is the pacing item for synergy capture.
**What they buy:**
- *Entry:* synergy/value tracking stand-up for a live deal (IMO value cell: weekly forum, monthly Go/Hold/Kill/Accelerate value review, quarterly re-litigation); or a 100-day value plan for a new portfolio company.
- *Follow-on:* the tech integration that actually paces synergies (Slalom core delivery), instrumented with live synergy dashboards; standing portfolio value office for PE (recurring, not deal-dependent — holds now average 6.6+ years).
**Evidence spine:** Bain 70% overestimate / 30% hit; cost vs revenue synergy gap (70–85% vs 25–35%); KPMG green-scorecards-while-value-burns; "12 is the new 5"; A&M 58% deploy in first 100 days; AI expectation gap in portcos (70% of GPs expect EBITDA impact, 6% can prove it).
**Slalom wedge:** IT-dependency of synergies (50–60% of synergy initiatives IT-dependent [C — verify]); MBB fill this with productised IP (McKinsey myIMO), not hands-on engineering — the IMO's data pipes and live dashboards are a build job.

### One methodology under all four
Same spine every time — **Frame the value (driver tree, baselines, owners) → Fund it right (staged, evidence-gated, partner-funded where possible) → Track it live (in the platforms, in Finance's numbers) → Govern to realise (cadence, kill authority, re-baselining).** The four offerings are doors into the same room. That's what makes this a capability, not four products.

---

## 7. The stat bank

### Headline stats (client-deck safe, with attribution)
| # | Stat | Source | Conf. |
|---|---|---|---|
| 1 | Only 0.5% of ~16,000 projects hit cost, time AND benefits targets | Flyvbjerg, *How Big Things Get Done*, 2023 | B |
| 2 | 88% of transformations fail original ambition; dedicated value owner → +24% value captured | Bain, 24,000-initiative DB, 2024 | A |
| 3 | Even successful transformations capture only 67% of possible value; ~25% of loss occurs at business-case stage | McKinsey, 2021 (live framework) | B |
| 4 | Only 39% of orgs mostly/always deliver full project benefits | Wellingtone, 2025/26 | A |
| 5 | 89% of ops leaders say tech investments haven't fully delivered | PwC, 2026, n=767 | A |
| 6 | 84% of finance orgs adopted/adopting AI; only 7% report high impact; 62% of CFOs: fewer than a quarter of AI initiatives deliver measurable benefit | Gartner CFO survey, PR Mar 2026 | A |
| 7 | Board-level AI value reporting: 4% today → expected standard by end-2026 | Deloitte, 2025/26 | A |
| 8 | Top 20% of companies capture 74% of AI's economic value (7.2x peers) | PwC AI Performance Study, Apr 2026 | A |
| 9 | 10-20-70: 70% of AI value is people & process redesign | BCG, 2025 | A |
| 10 | Full AI cost visibility → 5x more likely to have established ROI (only 26% have it) | KPMG Global AI Pulse, Q2 2026 | A |
| 11 | >40% of agentic AI projects cancelled by end-2027 (prediction) | Gartner, Jun 2025 | A |
| 12 | 46% of SaaS licenses unused; ~$19.8M wasted per enterprise per year | Zylo SMI 2025/26 | A |
| 13 | Cloud waste rose to 29% in 2026 — first increase in 5 years | Flexera State of the Cloud 2026 | A |
| 14 | Only 33% of Salesforce AI initiatives meet expected ROI | IBM IBV State of Salesforce 2025/26 | A |
| 15 | Only 42% fully achieve expected cloud value (up 5 pts since 2020) | Accenture Cloud Outcomes, Jan 2025 | A |
| 16 | ~70% of acquirers overestimate synergies; only ~30% hit targets; revenue synergies capture 25–35% vs 70–85% for cost | Bain / McKinsey | A/B |
| 17 | "12 is the new 5": PE needs ~12% annual EBITDA growth for 2.5x MOIC | Bain Global PE Report 2026 | A |
| 18 | 58% of PE firms deploy value resources in first 100 days — 2x prior year | A&M Value Creation Report 2026 | A |
| 19 | 66% of boards condition AI funding on proof of return; 43% of CFOs asked for a number they can't produce | RGP / CloudZero cluster, 2026 | A/C |
| 20 | 83% of M&A deals failed to lift shareholder returns while scorecards showed green | KPMG via L.E.K. | B |

### Stats to avoid (or heavily caveat)
- **MIT NANDA "95% of GenAI pilots fail"** — small non-random sample, not peer-reviewed, contradicted by its own shadow-AI finding. If used at all: as a hook, immediately paired with "and here's what the 5% do differently."
- **Gartner "30% of GenAI projects abandoned by end-2025"** — a 2024 *prediction*, never verified as an outcome. Label it a prediction or drop it.
- **"70% of transformations fail"** — no empirical basis (1993 guess, recycled). Use Bain instead.
- **Vendor-aggregator agentic ROI multiples** (171–192%, 540% ROI) — untraceable marketing numbers.
- Vendor TEI studies (Databricks 417% ROI, SAP 251%) — usable only as "the value the vendor claims is available," never as independent evidence.

### Needs verification before client use (secondhand in this pass)
- PMI "<40% track benefits post-go-live" (confirm vintage/publication); "30–50% of M&A value lost to slow integration" (attributed McKinsey); the 79%/70%/63% no-plan-at-signing stats; 50–60% of synergy initiatives IT-dependent; "43% formally track synergy capture"; BCG Platinion 70%; ~80%-of-features-unused; Salesforce "37% fully embraced." Direct fetches to mckinsey.com, bain.com, bcg.com, flexera.com and most consulting domains were proxy-blocked this session — all findings from those domains came via search snippets. A verification pass is recommended for anything going verbatim onto a client-facing slide.

---

## 8. Open questions for the storyline (decisions before slides)

1. **Name & claim.** Deloitte owns "value architecture" language in a CFO context. Keep the name and out-position on practicality/plumbing, or differentiate the name?
2. **Anchor buyer.** The evidence says CFO/Finance is where value gets verified — but Slalom's relationships are often CIO/CDO-side. Which door does the deck lead with, and does the client-partner enablement story bridge the two?
3. **Lead offering.** The platform value assessment is the easiest sell (funded, fixed, concrete waste to find) but the least differentiated category; the VRO is the most differentiated but the biggest ask. Lead with the assessment as the wedge and the VRO as the destination?
4. **Tooling.** The market now expects a demoable value-tracking layer (Accenture Momentum, McKinsey instrumentation). Is the play to extend Slalom's AI Value Platform, build a lightweight value dashboard accelerator on client platforms (Snowflake/Databricks-native), or stay tool-agnostic initially?
5. **Commercial posture.** Fee-at-risk/outcome pricing is the market direction and Slalom's telemetry access makes it credible — but it needs internal appetite. Signal it in the deck or hold it back?
6. **Internal enablement scope.** The stated goal is client partners bringing you into value conversations. Does the deck need a companion one-pager per offering (the "when you hear X, call me" cheat sheet) as part of the GTM?

---

## Appendix: source landscape
Primary 2025–2026 anchors: Bain (transformation DB 2024; Global PE Report 2026; M&A Midyear 2026), McKinsey (State of AI Nov 2025; M&A survey Jan 2025), BCG (AI Impact Gap 2025; AI-Powered Transformation Office 2026), PwC (AI Performance Study Apr 2026; Digital Trends in Ops 2026; M&A Integration Survey), Deloitte (AI ROI paradox 2025/26), KPMG (Global AI Pulse Q2 2026), Gartner press releases (CFO/AI Mar 2026; agentic Jun 2025; IT spend Jul 2026), Flexera (State of the Cloud 2026; ITAM 2025), Zylo SMI 2025/26, IBM IBV (State of Salesforce 2025/26; AI ROI), FinOps Foundation (State of FinOps 2026), A&M (Value Creation Report 2026), L.E.K. (serial acquirers 2025), Wellingtone (State of PM 2025/26), Accenture (Cloud Outcomes Jan 2025; 360° Value Report 2025). Classic anchors: Flyvbjerg 2023; McKinsey "Losing from Day One" 2021; PMI BRM guide; Cranfield BDN; HM Treasury Green Book; Steve Jenner *Managing Benefits*.
