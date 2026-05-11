# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a research project studying job preferences and wage expectations among Ghanaian university graduates. The baseline survey data feeds a future pilot RCT. The pipeline is: raw Excel data → Stata cleaning/analysis → LaTeX report.

## Commands

**Compile the LaTeX report:**
```bash
pdflatex -output-directory Report Report/report.tex
# Run twice to resolve cross-references and bookmarks
```

**Run the Stata cleaning and analysis script:**
```bash
stata -do "Code/d1 cleaning.do"
# Or open in Stata GUI and run with Ctrl+D / Cmd+Shift+D
```

## Repository Structure

- `Data/` — raw Excel survey data (`Coding -General - Cega Project.xlsx`, sheet `Original data-kobocollect`)
- `Code/d1 cleaning.do` — sole Stata script: cleans data and runs all regressions
- `Report/report.tex` — LaTeX report summarising findings

## Stata Script Architecture (`d1 cleaning.do`)

The script runs sequentially with no functions or sub-files:

1. **Import** — loads from the Excel sheet; columns arrive as generic letter names (A, B, C…)
2. **Rename** — maps Excel columns to descriptive variable names (`wageWTAmin`, `wageWTAmax`, `jobfactors_*`, etc.)
3. **Wage cleaning** — converts free-text wage entries (hourly/daily/monthly/ranges/written words) to monthly GHC equivalents using constants `hrs_pm ≈ 173.33` and `dys_pm ≈ 21.67`; creates `wageWTAmin_monthly` and `wageWTAmax_monthly`
4. **Categorical cleaning** — encodes graduate status, drops non-graduates (`grad_code == 3`), creates binary indicators (`female`, `married`, `child`)
5. **Sample restriction** — drops observations missing wages or any key job-factor/demographic variable; final n = 52
6. **Analysis** — cross-tab of social job factors, then four robust OLS regressions of `wageWTAmin_monthly` and `wageWTAmax_monthly` on all `jobfactors_*` (excluding salary), with and without age/female controls

**Key variable naming conventions:**
- `jobfactors_*` — binary indicators for each job-choice factor (0/1)
- `wageWTA*_monthly` — cleaned monthly GHC wage variables
- `jobfactors_salary` is intentionally excluded from all regressions (commented out in the `keep` statement) to avoid circularity with the wage outcomes

## LaTeX Report Notes

- Compile with `pdflatex` (not `xelatex`/`lualatex`); uses standard CTAN packages
- Table 1 counts/shares for non-social job factors contain `\textit{[x]}` placeholders to be filled from Stata `tab` output
- Sample attrition counts (`\textit{[x]} dropped`) in Section 1 also need filling from Stata output
- Cross-reference `\ref{sec:notes}` in Table 2 note points to the `\paragraph{Salary excluded...}` label

## Writing Style

See `JMP.tex` in the Meaningful-Work repo for the author's writing style. Key rules:
- No em dashes (`---`); use commas or parentheses instead
- Significance stars ordered `*** p<0.01, ** p<0.05, * p<0.1` with commas
- Table notes as `\scriptsize Notes:` in a `\begin{minipage}{\linewidth}` block
