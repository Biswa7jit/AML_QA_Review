# AML Quality Assurance Review (Excel)

A no-code AML Quality Assurance (QA) review program — 10 fictional
analyst case files, each independently scored against an 8-point QA
checklist using a critical-error methodology. Built entirely in Excel
formulas. No Python, no SQL, no VBA, no add-ins.

This project connects most directly to hands-on QA experience: independent
testing of case decisions — not detection, not screening, but
*auditing the auditors* — is one of the five pillars of an AML
compliance program.

## Why this project

A QA function doesn't just check whether an analyst closed a case
correctly. It checks whether the case *file, on its own,* defensibly
supports that closure — the same standard an examiner or auditor
would apply months later with no memory of the case. This project
applies a scoring methodology to AML case review that will look
familiar to anyone who has run a QA program in another domain: a
weighted checklist, a small set of zero-tolerance critical criteria
that override the numeric score, and a rollup dashboard that surfaces
which specific dimension analysts struggle with most — not just an
aggregate pass rate.

## What's in this repo

```
AML_QA_Review.xlsx                  ← the main workbook (open this)
README.md
data/
  qa_scoring_rubric.csv             ← the 8 criteria and critical flags, standalone CSV
  case_summaries.csv                ← what each analyst did, standalone CSV
  qa_review_ratings.csv             ← the QA reviewer's raw Pass/Partial/Fail ratings + comments
  qa_review_output.csv              ← computed scores and outcomes, exported as CSV
screenshots/
  qa_review.png
  qa_dashboard.png
```

## Workbook structure

| Tab | Purpose |
|---|---|
| `README` | Methodology (same content as this file, for anyone who only opens the workbook) |
| `QA_Scoring_Rubric` | The 8 criteria, which 3 are critical (zero-tolerance) items, and the outcome logic |
| `Case_Summaries` | What each analyst actually did — the alert, the investigation, the disposition reached |
| `QA_Review` | The actual QA assessment — every case scored against all 8 criteria, with a reviewer comment, and live formulas for score/outcome |
| `QA_Dashboard` | Rollup metrics: overall pass rate, pass rate by individual criterion, and a per-analyst breakdown |

Open `QA_Review` and click any cell to see the live formula behind
the score and outcome — nothing is hardcoded.

## The 8 QA criteria

| # | Criterion | Critical? |
|---|---|---|
| Q1 | KYC information reviewed correctly? | No |
| Q2 | Customer risk rating appropriate? | No |
| Q3 | Transaction red flags identified? | **Yes** |
| Q4 | Investigation sufficiently documented? | No |
| Q5 | RFI appropriate? | No |
| Q6 | Adverse media properly assessed? | No |
| Q7 | Escalation required? | **Yes** |
| Q8 | Final disposition supported by evidence? | **Yes** |

## The scoring methodology

- Each of the 8 criteria is scored **Pass (1) / Partial (0.5) / Fail
  (0)** for every case.
- **Three criteria are marked CRITICAL.** A Fail on any one of these
  overrides the numeric score entirely and produces `FAIL - Critical
  Error` — the same design as a zero-tolerance error category in a
  call-center or operations QA scorecard. A missed red flag or an
  unsupported disposition isn't a few-points deduction; it's the kind
  of finding that gets a case reopened, regardless of how well the
  rest of the file is written.
- Outside of a critical failure, the average score across all 8
  criteria determines the outcome: **90%+ is PASS, 75–89% is PASS
  WITH COACHING, and below 75% is FAIL.**

## Key formulas used

All standard Excel — no add-ins, no VBA:

- **`COUNTIF` across a horizontal range** — counting how many of a
  case's 8 criteria scored Pass vs. Partial, to compute the numeric
  score in one formula rather than eight nested terms.
- **`OR`** — checking the three critical criteria for a Fail,
  independent of the numeric score.
- **Nested `IF`** — translating the critical-failure flag and score
  into the final plain-English outcome.
- **`COUNTIFS` / `AVERAGEIF` on the Dashboard sheet** — pass-rate
  rollups sliced by criterion and by analyst.
- **Conditional formatting** — green/amber/red/dark-red by outcome
  severity, so the scorecard is scannable at a glance.

## Results

Running the QA program against the 10 synthetic cases produces:

| Outcome | Count | % of Cases |
|---|---|---|
| PASS | 3 | 30% |
| PASS WITH COACHING | 3 | 30% |
| FAIL | 1 | 10% |
| FAIL - Critical Error | 3 | 30% |

![QA Review scorecard](screenshots/qa_review.png)

**The dashboard's per-criterion breakdown is the more useful finding
for a real QA program** — it isn't just "3 analysts failed," it's
*why*:

![QA Dashboard](screenshots/qa_dashboard.png)

"Investigation sufficiently documented?" has the lowest pass rate in
the sample (30%) — lower even than the critical criteria. That's a
realistic and important pattern: analysts often reach the *correct*
conclusion without leaving a file that proves it, which is invisible
in a simple pass/fail case count but is exactly what a criterion-level
rollup is built to surface. A QA program that only tracks overall
pass rate would miss that documentation, specifically, is the
recurring gap — not analytical judgment.

## The two failure modes, and why they're scored differently

- **QA-007** (missed a linked-account layering pattern visible in the
  case tool's own link-analysis panel) and **QA-009** (a confirmed
  regulatory-action adverse media hit dismissed with a one-line note)
  are both `FAIL - Critical Error` — a red flag was missed or evidence
  was mishandled, and no amount of good paperwork elsewhere in the
  file offsets that.
- **QA-010** scores an almost identically low 44%, but is scored a
  plain `FAIL`, not critical — because on independent re-review, the
  underlying transaction pattern likely didn't warrant escalation.
  The deficiency is that the file doesn't *demonstrate* that; it's a
  documentation failure, not a missed risk. Collapsing both cases into
  a single "FAIL" bucket would treat a paperwork gap and a missed
  layering pattern as the same severity of problem, which they aren't
  — and a QA program that can't tell the two apart risks both
  over-reacting to weak files and under-reacting to genuinely missed
  risk.

## A note on the sample data

The 10 cases are deliberately weighted toward findings (3 clean
passes, 3 pass-with-coaching, 1 fail, 3 critical failures) to
showcase the range of things a QA review should catch — a correct
escalation with no documented reasoning, an unnecessary RFI that
added friction without changing the outcome, an RFI response accepted
without cross-checking it against the KYC file already on hand, and a
clearly missed red flag sitting in a standard part of the case tool.
This is not meant to represent a real program's typical pass rate.

## Limitations (stated honestly)

- Synthetic data only — 10 cases, sized for a readable demonstration.
  A live QA program samples a statistically meaningful percentage of
  closed cases on a recurring cycle, not a one-time review.
- The specific weighting (Pass=1/Partial=0.5/Fail=0, 90%/75%
  thresholds, which 3 criteria are critical) is a policy design
  choice for this project, not a universal QA standard — real
  programs calibrate this against their own regulatory findings and
  risk appetite.
- No tracking of QA findings over time (trend analysis, whether a
  given analyst's error rate is improving after coaching) — this
  workbook produces a point-in-time review only.
- A single QA reviewer's ratings and comments are shown; real programs
  often include a calibration process across multiple reviewers to
  keep scoring consistent.

## About

Built by Biswajit Das, a Quality Analyst with 3+ years of QA
experience (auditing task execution against policy and compliance
standards) and CAMS certification, as a portfolio piece connecting
that QA background directly to AML case review.
