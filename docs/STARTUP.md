# Zenith Open Source Projects: ExcelAgento startup thesis

This document turns the product direction into a shared, honest startup narrative. It is a
positioning and launch guide, not a claim that roadmap items already exist in the application.

## The company and the product

- **Company / umbrella:** Zenith Open Source Projects
- **Flagship product:** ExcelAgento
- **Category:** open-source, plain-English workspace software for spreadsheet-heavy work
- **Core promise:** people can describe a spreadsheet outcome in ordinary English and get a
  reviewable, reversible result without handing control of their workbook to an opaque agent.

### One-line positioning

> ExcelAgento is the Zenith workspace that turns messy spreadsheet instructions into verified,
> reversible work.

### Short pitch

Teams still lose hours cleaning exports, reconciling tabs, normalizing dates, and preparing weekly
reports. ExcelAgento lets a person upload a workbook, explain the desired outcome, inspect the
agent's plan, and approve only the change they understand. Deterministic operations, previews,
invariant checks, and undo make the useful path safer than asking a general chatbot to rewrite a
file.

## Who we serve first

The first wedge is not “everyone who uses Excel.” It is people whose workbooks are operational
interfaces and whose mistakes are expensive:

1. **Operations and finance generalists** who clean recurring exports before a close, review, or
   handoff.
2. **Analysts and founders** who need quick, explainable transformations without building a data
   pipeline for every one-off request.
3. **Small teams** that want an open, inspectable alternative to sending sensitive workbooks into
   an unbounded SaaS automation black box.

### The job to be done

> “Take this messy workbook and make it usable for the next decision, while showing me what
> changed and giving me a safe way back.”

## Product wedge

The first repeatable workflow is:

1. Load an `.xlsx`, `.xls`, or `.csv` file.
2. Say what needs to happen in plain English.
3. Let the local engine and agents inspect the workbook and prepare a bounded plan.
4. Review the affected cells and the assumptions.
5. Confirm the operation, export the result, and keep the history boundary for undo.

The product should win one weekly spreadsheet task before it tries to become a general-purpose
workspace. Supported transformations and statistics are more valuable than a vague promise of
autonomy.

## The Zenith trust standard

“Trust” must be observable in the product, not only in the brand language:

- **Local-first:** parsing, deterministic cleaning, analysis, history, and checkpoints run in the
  browser by default.
- **Bounded intelligence:** model output is untrusted input; only cataloged operations can reach
  the engine.
- **Preview before mutation:** the user sees the proposed operation and affected range before an
  apply action.
- **Explicit destructive gates:** operations that can remove or overwrite data require confirmation.
- **Transactional history:** a successful plan commits once and exposes a reversible boundary.
- **Honest failure:** invalid, unsupported, stale, or interrupted work is surfaced instead of being
  presented as success.
- **Open implementation:** the engine, agent boundary, fixtures, and evaluation suite are visible
  to contributors and adopters.

These are current product principles. They should remain true as hosted features are added.

## What exists today

ExcelAgento already has a strong technical foundation for this wedge:

- a pure TypeScript workbook engine with schema validation and invariants;
- `.xlsx` / `.xls` / `.csv` import and export in the web app;
- deterministic local planning and formula-aware analysis;
- optional bring-your-own-key provider adapters;
- multi-agent planning for complex requests;
- reviewable missions, history, checkpoints, usage visibility, and browser-local memory;
- fixtures, unit tests, evaluation cases, and a Vercel deployment path.

The current application is a browser-first product, not an unattended server queue. Marketing and
documentation should keep that boundary clear.

## Business model hypothesis

The open-source project is the trust and distribution layer. Revenue should pay for convenience,
team coordination, and support without making the safe core opaque.

### Near-term options to test

- **Free open-source core:** local workbook operations and single-user workflows remain available
  without a hosted model key.
- **Hosted team layer:** shared workspaces, policy controls, audit retention, and managed provider
  access are a future paid surface; they are not promises of the current browser-only release.
- **Support and implementation:** paid onboarding, private deployment help, and reliability reviews
  can serve teams that need confidence before adopting the project.
- **Managed AI convenience:** an optional hosted model plan can remove BYOK setup while preserving
  the same engine guardrails and usage transparency.

The first pricing decision should follow evidence from recurring workflows and support demand, not
become a reason to weaken local-first behavior.

## Launch message

Use the following message consistently across the website, README, and launch posts:

> **Tell your workbook what needs to change.**
> ExcelAgento by Zenith turns plain-English spreadsheet instructions into reviewable,
> reversible work. Open source, local-first by default, and designed so the agent has to earn the
> right to touch the sheet.

### Proof to show, not just say

1. A short screen recording of “normalize dates and remove duplicates” with the preview visible.
2. A before/after fixture where the affected cells and the undo boundary are easy to inspect.
3. The exact engine operation and evaluation case behind a common request.
4. A plain-language data-flow explanation for BYOK providers.
5. A visible example of the system refusing an unsupported or unsafe request.

## First 90 days

### 0–30 days: earn trust with one workflow

- Choose one recurring workflow (for example, cleaning weekly exports) and measure completion.
- Publish three synthetic before/after examples and the matching prompts.
- Keep the landing page outcome-led and link each trust statement to product behavior or source.
- Collect feedback from operators, analysts, and small-team leads rather than generic “AI users.”

### 31–60 days: improve repeatability

- Turn the most common successful prompts into named workflows with visible assumptions.
- Track successful plan rate, cancellation rate, stale-preview rate, and export completion.
- Add regression fixtures for every workflow that becomes part of the public promise.
- Document the data path and open-source contribution route.

### 61–90 days: test the paid boundary

- Interview teams that need shared policy, audit, or managed model access.
- Prototype the smallest hosted convenience that does not compromise the local core.
- Publish a transparent pricing experiment and a self-hosting path before asking for larger trust.

## Product metrics

The north-star signal is **verified workbook outcomes**, not chat volume. Track:

- workbooks reaching at least one confirmed operation;
- repeat users returning to complete the same class of task;
- confirmed operations that are later undone;
- unsupported or unsafe requests correctly blocked;
- time from upload to exported result;
- provider spend and model usage when BYOK is enabled;
- issue reports and contribution activity for the open-source project.

Avoid optimizing for autonomous actions, token count, or impressive demos that hide the review step.

## Guardrails for the brand

- Say **“built for reliability”** or point to a concrete guarantee; do not claim perfection.
- Say **“open-source core”** only where the relevant code and license are actually available.
- Say **“local-first”** with the provider/data-flow caveat when a user connects an external model.
- Never imply that an AI result is certified financial, legal, tax, or accounting advice.
- Keep the product name **ExcelAgento** and the parent brand **Zenith Open Source Projects**
  distinct so future Zenith workspaces can share the same trust standard.

## Immediate next step

Use the current landing page to recruit a small set of design partners around one weekly workbook
task. The next product change should be chosen from those observed tasks and protected by a fixture,
an evaluation case, and a user-visible acceptance test.
