# Variance Reduction Techniques for Retirement Solvency Simulation

Monte Carlo simulation project applying and comparing variance reduction techniques to estimate the solvency probability of a retiree following the 4% withdrawal rule.

## The Question

If a retiree withdraws 4% of their portfolio annually (adjusted for inflation), what is the probability their savings outlast them? This project estimates that solvency rate via Monte Carlo simulation and then asks a harder question: how do you get a stable, trustworthy estimate, not one that swings by a percentage point every time you rerun it?

## Approach

A baseline simulation draws 10,000 random retiree paths (random lifetime × random monthly market returns) and tracks portfolio balance to determine solvency. Because every draw is random, repeated runs of the baseline produced solvency estimates ranging as much as **1.17 percentage points** across 15 runs; too much noise for confident financial advice.

Four variance reduction techniques were implemented and benchmarked against the baseline, each isolated in its own worksheet:

| Technique | Idea |
|---|---|
| **Antithetic Variates** | Pair every random market path with its mirror image so a lucky draw is automatically offset by an unlucky one, canceling noise on averaging. |
| **Control Variates** | Use a related quantity with a known expected value (market performance vs. expectation, and optionally lifetime vs. expectation) to correct the raw estimate for how "lucky" a given run's draws were. |
| **Stratified Sampling** | Divide the range of possible lifetimes into equal segments and force exactly one draw per segment, guaranteeing short and long lifetimes are always represented. |
| **Antithetic + Stratified (combined)** | Antithetic variates address market-return noise; stratified sampling addresses lifetime noise — combining both targets both sources of variance at once. |

## Results

All three independent variance reduction approaches converged on a solvency estimate of **~96.8%**, versus the baseline's **95.66%** — agreement across independent methods is strong evidence the baseline was mildly underestimating solvency, and that ~96.8% is close to the true value.

- **Antithetic variates** — simplest to implement, strong variance reduction, best effort-to-payoff ratio.
- **Control variates** — most accurate and most flexible (extendable to multiple control variables), but the most technically involved.
- **Antithetic + Stratified** — best overall balance: near-best stability with implementation complexity well below control variates.

Full methodology, formulas, and the general-audience explanation are in the written report; all formulas and results are in the workbook.

## Repo Contents

- [`report/Monte_Carlo_Variance_Reduction_Report.docx`](report/Monte_Carlo_Variance_Reduction_Report.docx) — full write-up: plain-language overview plus the technical methodology, assumptions, and results.
- [`model/Variance_Reduction_With_Monte_Carlo_Simulation__Senior_Capstone_.xlsm`](model/Variance_Reduction_With_Monte_Carlo_Simulation__Senior_Capstone_.xlsm) — Excel/VBA workbook with one sheet per method (`Var1` baseline, `AntitheticVariate`, `ControlVariate`, `StratifiedSampling`, `Antithetic+Stratified`) plus a `Results` summary sheet. Enable macros to run the simulations.

## Tools

Excel, VBA (simulation engine and automation), Monte Carlo methods, variance reduction techniques (antithetic variates, control variates, stratified sampling).
