# RAW: Platform Unlock Mechanics, Honest Funding, Strategy-Led vs Commodity (Agent Output, 2026-08-19)

## Key mechanics found, named and real

- Salesforce offers Optimizer, which is a free self-service scan of a customer's org, and its partners offer free health checks too, though those are really just commodity lead generation, taking anywhere from 24 hours to 5 days and working off a checklist. Salesforce also runs Signature and Premier Success Plans itself, which the client pays for, and these produce a Customer Success Score built from Product Adoption, Expertise, and Technical Health.
- ServiceNow offers a HealthScan and Configuration Review, delivered by consultants, and its Now Value methodology moves through vision, strategic drivers, business outcomes, and success measurement in that order. ServiceNow's Impact tiers are priced as a percentage of the subscription and the client pays for them, and UpperEdge has flagged that they carry hidden costs. GuideVision has a good line worth reusing: projects finish "on deadline, within budget... but after the project is over, a disenchantment often sets in that the value isn't delivering."
- SAP retired its Pathfinder tool and replaced it with the Signavio Process Insights discovery edition, which is free and usage-based and went generally available in February 2024. Signavio's process mining compares how a process was designed against how it is actually executed, and SAP pairs its Value Lifecycle Manager with Signavio to run process-driven value management. SAP's Preferred Success offering, on the other hand, is paid for by the client.
- Microsoft's FastTrack is genuinely free for any customer with 150 or more seats, for the life of the subscription, which makes it the cleanest example of a vendor actually funding this work itself. Microsoft 365 also has a self-service Adoption Score built from the platform's own usage data.
- Workday's Adoption Kit is really just a content library rather than an actual assessment, and Workday has no native adoption tool of its own; any adoption work is delivered by a partner, and the client pays for it.
- OneStream's MindStream Health Check is a one week diagnostic, run as separate sessions for IT and platform admins on one side and finance and operations users on the other, and it produces a Technical Findings document plus an Improvement Roadmap split into quick wins and longer term items. Riveron, HollandParker, and SC&H all offer their own optimisation services on top of this, and OneStream also has a partnership with PwC.
- Snowflake has a Well-Architected Framework built around five pillars, of which cost is only one, and it comes with a scorecard, a set of Blueprints, and, as of 2026, additional lenses plus a CoCo AI review; it is also embedded as a custom lens inside AWS's own Well-Architected tooling. The framework is built around capability rather than being a cost audit.
- Databricks introduced a Partner Well-Architected Framework in February 2026, aimed at ISVs, and customers get their own Well-Architected Review delivered through partners. Aimpoint's "Strategy Accelerator" combines use case discovery with data readiness work and carries it through to a roadmap, deployment, and value tracking, while deviq offers a "4 Weeks to Production" engagement.
- For AWS, the relevant mechanic is the Well-Architected Review. MAP is not relevant here, since MAP funds migration work rather than acting as a value or usage mechanism.

## Gap evidence, business flavoured

- Nexthink's Soft-WASTE research, covering more than 6 million environments in 2023, found that 49.96% of installed software goes completely unused.
- Zylo's 2026 research found that the average enterprise runs 305 SaaS apps, with about 46% of licensed seats sitting unused, and that business units control 70% of SaaS spend against just 26% controlled by IT; that split is exactly why duplication happens, because nobody has visibility over the whole portfolio.
- Gartner argues that application rationalisation has to move "Beyond Cost Savings," in its own document titled that (document 6780334), which means Gartner itself is signalling the shift from a cost conversation to a strategy conversation.
- WalkMe's 2025 research found that firms underestimate their actual technology usage by 1,600%, and that their sample of companies wasted $104 million in 2024.
- The AFP's 2025 research found that more than 90% of finance practitioners use spreadsheets weekly for planning and reporting, and that 96% of FP&A teams use spreadsheets for planning even though they already own an ERP or EPM platform. This is real evidence for the pattern where Excel sits alongside the platform instead of being replaced by it.
- TDEOS's 2026 research on "Shadow CRM" found that Excel and Notion trackers often hold data that people trust more than the official CRM, and that this gets fixed simply by turning on a module the company already owned; this is worth using as a pattern to describe, without naming any specific client.
- Surety Systems frames this well, without placing blame: "the failure mode is quiet: nothing breaks, the system just sits underutilized while people work around it."
- OpenText's 2026 work on the ALA licence trap is the actual diagnostic mechanic to use: it reconciles three separate views of a platform, what was contracted, what was deployed or activated, and what is actually used, and the gaps between those three views cut both ways, creating both audit exposure and reclaimable value.

## Conversation mechanics: no insult, no upsell

- Sobel's advice is that the client has to own the new way of seeing the problem themselves, so the job is to walk them there step by step rather than push it on them.
- FusionSpan makes the point that underutilisation is simply the default outcome for any organisation, and it happens without anyone doing anything wrong: "people fit the new tool into current process without disturbing anything, which underutilises it."
- Avoid the word "wrong" altogether, and use collaborative, discovery-style language instead.
- Other firms' positioning language is worth knowing, even if we don't copy it. GuideVision says its work "maximises your ServiceNow investment." Riveron promises to "yield the greatest return on your investment." HollandParker says it will "activate exponential ROI." Prudent talks about "systems of execution rather than isolated systems of record." Surety's quiet-failure line above is the strongest of the set.

## Funding reality, honest

- Salesforce's FY27 Catalyst program pays partners for post-sale activation and consumption work, rather than paying clients directly, with a $1 billion target behind it, so it is not a grant the client ever sees. Beyond the free commodity checks, the client pays.
- ServiceNow's Impact offering is paid for by the client, priced as a percentage of the subscription, and its MDF and SIF funds are partner-facing and conditional on the deal.
- For SAP, only the automated tooling is free; any human-delivered adoption service is paid for by the client.
- Microsoft's FastTrack is genuinely free for 150 or more seats.
- For both Workday and OneStream, no vendor funding was found; the client pays.
- Snowflake's Partner Services Fund is gated by partner tier and certification, though its Well-Architected Framework tool itself is free.
- For Databricks, the most relevant fund is the Velocity "Activate" SPIFF, aimed at early adoption inside existing customers, and it is gated by partner tier.
- For AWS, MAP is not relevant, but the Well-Architected Review is sometimes free when delivered through an AWS Solutions Architect.
- The plain read across all of this is that by default the client pays for a short, sharply scoped assessment, because the value case carries the cost on its own; partner incentives exist mainly to offset the partner's own cost of delivering the work, rather than functioning as grants handed to the client. Microsoft's FastTrack and SAP's automated tooling are the only two things in this list that are genuinely free.

## Strategy-led versus commodity

- A commodity engagement looks like a licence audit or a free health check; it works off a checklist, covers configuration, security, and data quality, feels procurement flavoured, and exists mainly as lead generation.
- A strategy-led engagement starts from the outcome logic of the organisation's own strategy, the way Fujitsu's ResultsChain does. It talks about value leakage rather than wasted spend, the way PwC's VRO does. It prioritises by relevance to business outcomes rather than by licence dollars. And it ends in an ongoing operating cadence rather than a one-off report.
- ISG's credibility angle is independence, since it doesn't sell licences itself.

## A grounded two to three week engagement shape

In week one, the diagnostic work reconciles what was contracted against what was activated against what is actually used, runs the vendor's own native usage tooling as an objective baseline, and splits sessions between platform admins and IT on one side and business users on the other, so the team can see where workarounds exist, where Excel sits next to the platform, and why.

In week two, the work maps to business direction. An executive workshop covers where the business is actually going over the next 12 to 24 months, and the capability inventory gets cross-referenced against that specific direction rather than against a generic maturity checklist, so that everything gets prioritised by its relevance to a real business outcome.

In week three, the roadmap and handoff happen. Every item is tied to a named business outcome, and each one is sorted into either "turn on now," meaning no new build is needed, or "needs a scoped enablement phase," which is the natural point where the engagement converts into further work. The week closes by proposing a lightweight measurement cadence, so the client ends up with an ongoing operating discipline rather than a one-off report.
