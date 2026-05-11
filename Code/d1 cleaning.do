/* ============================================================================
   d1 cleaning.do
   ----------------------------------------------------------------------------
   Purpose  : Load, clean, and prepare survey data on job preferences among
              Ghanaian university graduates; then run preliminary regressions.
   Input    : Coding -General - Cega Project.xlsx  (sheet: Original data-kobocollect)
   Output   : Clean dataset in Stata memory; regression results printed.

   ============================================================================ */

/* ─── SECTION 1: Environment setup ─────────────────────────────────────────── */

clear all
// Remove any dataset currently held in Stata's memory.
// Starting with a clean slate prevents old data from accidentally mixing
// with the data we are about to load.

cd "/Users/remybeauregard/Documents/GitHub/Low-Paid-Jobs/Data"
// Set the "working directory": the default folder where Stata looks for files.
// Any filename used without a full path is read from (and saved to) this folder.

capture mkdir "../Output"
// Create the Output directory if it does not already exist.
// "capture" suppresses the error that would otherwise occur if the folder
// already exists. All log files and any other generated output go here.

capture log close
log using "../Output/d1_cleaning.log", replace text
// Open a plain-text log file that records every line of Stata output for
// this run. "replace" overwrites the previous log; "text" saves as .log
// (not binary .smcl). After running this script, check this file to verify
// any numbers that appear in the report before reporting them.

import excel "Coding -General - Cega Project.xlsx", sheet("Original data-kobocollect") clear
// Load data from the specified Excel workbook, reading the worksheet named
// "Original data-kobocollect". The "clear" option discards whatever is
// currently in memory before importing.
// After import, each Excel column becomes a Stata variable. Because the first
// spreadsheet row holds column headers rather than data, Stata cannot use
// those headers as variable names directly and instead assigns generic letter
// names: A, B, C, D, … (one letter per column).


/* ─── SECTION 2: Drop uninformative columns and the duplicate header row ─────── */

drop A B D E F
// Remove columns A, B, D, E, F — these held spreadsheet metadata or
// submission identifiers (e.g. timestamps, form IDs) not needed for analysis.

drop in 1
// "drop in 1" removes the very first observation (row 1 of the data).
// When Stata imported the Excel file the original header row became the first
// data row, filled with question-text labels rather than real values.
// Deleting it leaves only genuine respondent data.


/* ─── SECTION 3: Rename variables from Excel column letters to meaningful names ─ */
// The generic letter names (C, G, H, …) are replaced with short, descriptive
// names that reflect each survey question's content.
// "la var" attaches a longer human-readable label that appears in output tables.

rename C graduate
la var graduate "Which of the following best describes you?"
// Records the respondent's qualification status (e.g. recent graduate,
// current student, etc.).

rename G seekingwork
la var seekingwork "Are you are looking for work in your field or not?"

rename H wageWTAmin
la var wageWTAmin "What is the minimum wage you will be willing to take?"
// WTA = "Willingness To Accept". This is the respondent's reservation wage
// floor: the lowest monthly pay they would accept before turning a job down.

rename I wageWTAmax
la var wageWTAmax "What is the maximum wage you will be willing to take?"
// The respondent's wage ceiling — the highest pay they consider appropriate
// (above which they may feel over-qualified or the offer implausible).

rename J acceptMW
la var acceptMW "If you are offered a minimum wage for a job offer will take it?"

rename M fieldofstudy
rename N graduationyear
rename P age
rename Q gender
rename R maritalstatus
rename S livingarrangement
rename T children
rename U nationalserviceDATE  // Date national service was completed
                               // (compulsory in Ghana before formal employment)
rename V nationalservicePAY   // Monthly stipend received during national service

// ── Job-factor variables ────────────────────────────────────────────────────
// Each variable below corresponds to a survey item asking whether that factor
// influences the respondent's job-choice decision.
// Typical coding: 1 = this factor matters to me, 0 = it does not.

rename Z  jobfactors_salary
rename AA jobfactors_location
rename AB jobfactors_familyopinion   // Does family opinion influence job choice?
rename AC jobfactors_peeropinion     // Does peer/friend opinion influence job choice?
rename AD jobfactors_hoursflexibility
rename AE jobfactors_qualifications  // Match between the job and own qualifications
rename AF jobfactors_promotion       // Promotion prospects
rename AG jobfactors_profdev         // Professional development opportunities
rename AH jobfactors_training        // On-the-job training


/* ─── SECTION 4: Convert job-factor variables from text to numbers ──────────── */
// When Excel data is imported, numeric-looking values sometimes arrive as plain
// text ("string variables"). Statistical commands such as regression require
// numeric variables. "destring" strips stray non-numeric characters and
// converts each variable to a proper number. "replace" overwrites the string
// version in place rather than creating a separate copy.

foreach v in jobfactors_salary jobfactors_location jobfactors_familyopinion ///
             jobfactors_peeropinion jobfactors_hoursflexibility              ///
             jobfactors_qualifications jobfactors_promotion                  ///
             jobfactors_profdev jobfactors_training {
    destring `v', replace
    // The loop body runs once per variable name in the list.
    // `v' is replaced by the current name on each iteration, so Stata executes
    // "destring jobfactors_salary, replace", then "destring jobfactors_location,
    // replace", and so on through all nine job-factor variables.
}


/* ─── SECTION 5: Wage variable cleaning ────────────────────────────────────── */
// Problem: respondents entered wages as free text in many different formats:
//   "GHC 2,000/month"    "15 Ghana Cedis per hour"    "1500-2000"
//   "Two thousand"        "3000 GHS and above"
// Goal: convert every entry to a single comparable unit — monthly GHC
// (Ghanaian Cedis) — so wages can be used as outcome variables in regressions.
//
// Conversion constants (40-hr week, 5-day week; standard Ghanaian assumption):
//   Hours per month = 40 hrs/wk × 52 wks ÷ 12 months ≈ 173.33
//   Days  per month =  5 days/wk × 52 wks ÷ 12 months ≈ 21.67

local hrs_pm = 40 * 52 / 12
// "local" defines a temporary named value called hrs_pm. Writing `hrs_pm'
// anywhere below pastes in the computed number (173.333…).

local dys_pm =  5 * 52 / 12
// Days per month (21.667). Both locals persist only for the duration of this
// do-file run.

foreach var in wageWTAmin wageWTAmax {
    // Run the entire cleaning block first for wageWTAmin, then for wageWTAmax.

    capture confirm variable `var'
    if _rc continue
    // "capture" runs "confirm variable `var'" and suppresses any error message.
    // If the variable does not exist, _rc (the return/error code) is non-zero
    // and "continue" jumps to the next loop iteration, skipping this variable.
    // This is a safety guard in case a column is unexpectedly absent.

    * If the variable imported as numeric, convert to string first
    capture confirm string variable `var'
    if _rc tostring `var', replace
    // Check whether the variable is already stored as text.
    // If it is NOT a string (_rc ≠ 0), "tostring" converts it to text.
    // We need a text variable so we can inspect and manipulate each cell's
    // characters before extracting the number.

    gen str500 _s = lower(trim(`var'))
    // Create a temporary working variable "_s" (capacity: 500 characters).
    // "lower()" converts all letters to lowercase ("GHC" → "ghc"),
    // "trim()" removes leading and trailing blank spaces.
    // Standardizing case and spacing prevents mismatches caused by
    // inconsistent capitalization in free-text responses.

    * Detect time unit before stripping keywords
    gen byte _isHour = regexm(_s, "hour")
    // Creates a 0/1 indicator: 1 if the response contains the word "hour",
    // 0 otherwise. "regexm" performs a regular-expression search anywhere
    // in the text. "byte" declares a compact integer type sufficient for 0/1.

    gen byte _isDay  = regexm(_s, "\bday\b")
    // Same for "day". The "\b" symbols are word-boundary anchors: they ensure
    // "day" matches only as a whole word, not inside "sunday" or "today".
    // We capture time-unit flags BEFORE stripping unit words so we retain
    // the information needed to apply the correct conversion multiplier.

    * Written-word numbers
    replace _s = "1000" if regexm(_s, "one\s*thousand")
    replace _s = "2000" if regexm(_s, "two\s*thousand")
    replace _s = "3000" if regexm(_s, "three\s*thousand")
    // "\s*" means "zero or more whitespace characters", so both
    // "one thousand" and "onethousand" are matched.

    * Strip currency labels and unit words (longest patterns first)
    // After this loop, only the bare number (or range) should remain in _s.
    // Patterns are listed longest-first to prevent partial matches:
    // e.g. "ghana cedis" must be removed before "cedi" so we do not leave
    // a stray "s" after stripping "cedi" from "ghana cedis".
    foreach p in "ghana cedis" "ghana cedi" "gh cedis" "gh cedi" ///
                 "ghc" "ghs" "cedis" "cedi" "gh"                  ///
                 "per hour" "per day" "per month"                  ///
                 "/month" "/hour" "/day"                           ///
                 "and above" "above"                               ///
                 "month" "hour" "day" {
        replace _s = subinstr(_s, `"`p'"', "", .)
        // subinstr(original, find, replace_with, max_replacements)
        // "." as the last argument means "replace all occurrences".
        // The outer `"…"' quoting lets us embed literal double-quote characters
        // inside a Stata string.
    }

    * Remove thousand-separator commas and tidy whitespace
    replace _s = subinstr(_s, ",", "", .)
    // e.g. "1,500" → "1500"

    replace _s = strtrim(itrim(_s))
    // "itrim" collapses internal runs of multiple spaces to one;
    // "strtrim" removes any remaining leading or trailing spaces.

    * Handle ranges (e.g. "1500-2000"): take midpoint
    // Some respondents entered a range. We represent this as the average of
    // the two endpoints.
    gen double _lo = real(substr(_s, 1, strpos(_s, "-") - 1)) if strpos(_s, "-") > 1
    gen double _hi = real(substr(_s, strpos(_s, "-") + 1, .))  if strpos(_s, "-") > 1
    // "strpos(_s, "-")" returns the character position of the first hyphen.
    //   If > 1, a range was entered (a leading minus sign would sit at position 1).
    // "substr(text, start, length)" extracts a substring; "." as length means
    //   "to the end of the string".
    // "real()" converts the extracted text to a number.

    * Parse and convert to monthly
    gen double `var'_monthly = real(_s)
    // Attempt to read the cleaned text as a plain number. Entries that still
    // contain non-numeric characters (or are blank) become missing (.).

    replace `var'_monthly = (_lo + _hi) / 2 if !missing(_lo)
    // For range entries, overwrite the parsed value with the midpoint.

    replace `var'_monthly = `var'_monthly * `hrs_pm' if _isHour
    // If the respondent stated an hourly wage, multiply by hours per month.

    replace `var'_monthly = `var'_monthly * `dys_pm' if _isDay
    // If the respondent stated a daily wage, multiply by days per month.

    la var `var'_monthly "Monthly GHC equivalent of `var'"

    drop _s _isHour _isDay _lo _hi
    // Remove all temporary working variables before the next loop iteration.
}


/* ─── SECTION 6: Categorical variable cleaning ──────────────────────────────── */

encode graduate, g(grad_code)
// "encode" creates a new numeric variable (grad_code) where each unique text
// value of "graduate" is assigned a consecutive integer (1, 2, 3, …).
// Stata stores the original text as a "value label" so output still shows text,
// but the variable is numeric — required for use in regressions.

drop if grad_code==3 // "I am not a graduate"
// Remove observations whose graduation-status category is coded 3.
// Category 3 corresponds to non-graduates; the analysis targets university
// graduates only.

replace graduationyear="2025" if graduationyear=="2025 (Expected)"
// Standardize the text "2025 (Expected)" to a plain four-digit year so that
// destring can convert it to a number.

destring graduationyear, replace
destring age, replace
// Convert graduation year and age from text strings to numeric values
// so they can be used as controls in regression models.

g female = gender=="Female"
// Create a binary indicator: 1 for female respondents, 0 for all others.
// The expression (gender=="Female") evaluates to 1 (true) or 0 (false)
// for each observation.

g married = maritalstatus=="Married"
// Binary indicator: 1 = currently married, 0 = single / divorced / other.

g child=children=="Yes"
// Binary indicator: 1 = has at least one child, 0 = does not.

encode livingarrangement,g(living_code)
// Convert living-arrangement text to a labelled numeric variable,
// using the same approach as "encode graduate" above.


/* ─── SECTION 7: Keep only variables needed for analysis ────────────────────── */
// "keep" discards every variable NOT listed here. This removes raw/uncleaned
// duplicates and any columns that were never renamed, leaving a tidy dataset.

keep graduate seekingwork acceptMW fieldofstudy graduationyear age           ///
     nationalserviceDATE nationalservicePAY                                  ///
     gender maritalstatus livingarrangement children                         ///
     wageWTAmin wageWTAmax                                                   ///
     wageWTAmin_monthly wageWTAmax_monthly                                   ///
     jobfactors_location jobfactors_familyopinion                            ///
     jobfactors_peeropinion jobfactors_hoursflexibility                      ///
     jobfactors_qualifications jobfactors_promotion jobfactors_profdev       ///
     jobfactors_training                                                      ///
     grad_code female married child living_code // jobfactors_salary
// Note: jobfactors_salary is commented out and therefore NOT retained.
// It is excluded from the analysis — likely because salary is mechanically
// related to the wage outcome variables, which would introduce circularity.


/* ─── SECTION 8: Drop observations with insufficient data ────────────────────── */

drop if (missing(wageWTAmin_monthly) & missing(wageWTAmax_monthly)) | ///
        missing(age, graduationyear, grad_code,     ///
                jobfactors_location, jobfactors_familyopinion,         ///
                jobfactors_peeropinion, jobfactors_hoursflexibility,   ///
                jobfactors_qualifications, jobfactors_promotion,       ///
                jobfactors_profdev, jobfactors_training)
// Remove any respondent that meets either condition:
//   (a) Has no valid monthly wage for BOTH the minimum AND maximum
//       (i.e. both wageWTAmin_monthly and wageWTAmax_monthly are missing), OR
//   (b) Is missing any core demographic or job-factor variable needed for
//       regression.
// "|" means OR; "&" means AND. "missing(x, y, z)" returns true if ANY of
// x, y, or z is missing. After this step the dataset contains 48 observations.


/* ─── SECTION 9: Descriptive tabulations ────────────────────────────────────── */

// ── 9a. Summary statistics for demographic and wage variables ─────────────────
// This block computes summary statistics and writes them directly into
// "../Report/sumstats_table.tex", which is \input into the report's Table 1.
// Re-run this script whenever the data change to keep the table current.
//
// "file open" opens a file for writing; "write replace" overwrites any
// existing file. "file write" sends text to the file. "quietly" suppresses
// Stata's screen output. "summarize" computes summary statistics; results
// are stored in r(mean), r(sd), r(min), r(max).

// Stata writes a complete \begin{table}...\end{table} block to sumstats_table.tex.
// In report.tex, this is \input at document level (not inside a tabular), which
// avoids LaTeX's restriction on \multicolumn inside \input'd files.

// sumstats_table.tex uses tabular*{\linewidth} so the table fills text width
// and the minipage note below matches width automatically. No extra packages needed.
// Columns: Mean | SD | p5 | p95 | Min | Max  (7 cols total including label).
// p5 and p95 for wage variables are the winsorisation thresholds; summarize,
// detail provides all percentiles in r(p5) and r(p95).

// _char(36) is the ASCII dollar sign ($). Using _char(36) directly in file write
// prevents Stata from treating $ as a global-macro prefix (e.g. $n → obs count).

file open sumf using "../Report/sumstats_table.tex", write replace

file write sumf "\begin{table}[h!]" _n
file write sumf "\centering" _n
file write sumf "\begin{threeparttable}" _n
file write sumf "\caption{Summary Statistics (" _char(36) "n = 48" _char(36) ")}" _n
file write sumf "\label{tab:sumstats}" _n
file write sumf "\begin{tabular}{lcccccc}" _n
file write sumf "\toprule" _n
file write sumf "\textbf{Variable} & \textbf{Mean} & \textbf{SD} & " ///
               _char(36) "p_{5}" _char(36) " & " _char(36) "p_{95}" _char(36) ///
               " & \textbf{Min} & \textbf{Max} \\" _n
file write sumf "\midrule" _n
file write sumf "\multicolumn{7}{l}{\textit{Demographics}} \\[2pt]" _n

// Age: 1 decimal for mean/SD, whole years for percentiles/min/max
quietly summarize age, detail
file write sumf "Age & " %5.1f (r(mean)) " & " %5.1f (r(sd)) ///
            " & " %5.0f (r(p5)) " & " %5.0f (r(p95)) ///
            " & " %5.0f (r(min)) " & " %5.0f (r(max)) " \\" _n

// Binary indicators: 2 decimal places throughout; Has children is last in the
// Demographics block so its row ends with \\[4pt] to space before Wages.
foreach v in female married child {
    quietly summarize `v', detail
    if "`v'" == "female"  local lab "Female"
    if "`v'" == "married" local lab "Married"
    if "`v'" == "child"   local lab "Has children"
    local row_end = cond("`v'" == "child", " \\[4pt]", " \\")
    file write sumf "`lab' & " %5.2f (r(mean)) " & " %5.2f (r(sd)) ///
                " & " %5.2f (r(p5)) " & " %5.2f (r(p95)) ///
                " & " %5.0f (r(min)) " & " %5.0f (r(max)) "`row_end'" _n
}

file write sumf "\multicolumn{7}{l}{\textit{Wages (monthly GHC)}} \\[2pt]" _n

// Wages: 0 decimal places; p5 and p95 are the winsorisation bounds
foreach v in wageWTAmin_monthly wageWTAmax_monthly {
    quietly summarize `v', detail
    if "`v'" == "wageWTAmin_monthly" local lab "Minimum WTA"
    if "`v'" == "wageWTAmax_monthly" local lab "Maximum WTA"
    file write sumf "`lab' & " %6.0f (r(mean)) " & " %6.0f (r(sd)) ///
                " & " %6.0f (r(p5)) " & " %6.0f (r(p95)) ///
                " & " %6.0f (r(min)) " & " %6.0f (r(max)) " \\" _n
}

file write sumf "\bottomrule" _n
file write sumf "\end{tabular}" _n
file write sumf "\begin{tablenotes}[flushleft]" _n
file write sumf "\scriptsize" _n
file write sumf "\item \textit{Notes:} Female, Married, and Has children are binary " ///
               "(0/1) indicators; their means equal the share of the sample with that " ///
               "characteristic. " _char(36) "p_5" _char(36) " and " ///
               _char(36) "p_{95}" _char(36) " for wage variables are the winsorization " ///
               "thresholds used in Table~\ref{tab:regsw}. Wages standardized to " ///
               "monthly GHC equivalents using a 40-hour, 5-day working week." _n
file write sumf "\end{tablenotes}" _n
file write sumf "\end{threeparttable}" _n
file write sumf "\end{table}" _n

file close sumf

// Additional descriptive tabs (for reference; not written to the report)
tab seekingwork       // Share actively seeking work in their field.
tab acceptMW          // Share willing to accept the minimum wage if offered.
tab livingarrangement // Distribution of living arrangements.
tab fieldofstudy      // Distribution of fields of study.

// ── 9b. Job-factor prevalence ─────────────────────────────────────────────────
// Tabulate each job-factor indicator separately (salary excluded; see Section 7).
// Each tab produces a two-row frequency table showing how many respondents
// have the factor coded 0 (does not matter) vs. 1 (matters).
// The row for "1" gives the count and percentage needed for Table 2 of the report.

foreach v in jobfactors_location      ///
             jobfactors_familyopinion ///
             jobfactors_peeropinion   ///
             jobfactors_hoursflexibility ///
             jobfactors_qualifications ///
             jobfactors_promotion     ///
             jobfactors_profdev       ///
             jobfactors_training {
    tab `v'
    // Prints a frequency table for the variable named in `v'.
    // The loop iterates through all eight job-factor indicators in turn.
}


/* ─── SECTION 10: Winsorization ─────────────────────────────────────────────── */
// Winsorisation replaces values beyond a given percentile threshold with the
// value at that threshold, limiting (but not removing) extreme observations.
// This is applied at the 5th and 95th percentiles to address extreme reported
// wages that may reflect measurement or reporting noise.
//
// For each wage variable, _pctile computes the 5th and 95th percentiles and
// stores them in r(r1) and r(r2). Values outside these bounds are capped.
// A new _w variable is created so the original values are preserved.

foreach v in wageWTAmin_monthly wageWTAmax_monthly {
    quietly _pctile `v', p(5 95)
    local lo = r(r1)
    local hi = r(r2)
    gen double `v'_w = `v'
    replace `v'_w = `lo' if `v'_w < `lo' & !missing(`v'_w)
    replace `v'_w = `hi' if `v'_w > `hi' & !missing(`v'_w)
    la var `v'_w "`v' winsorized at 5th/95th percentile (monthly GHC)"
}


/* ─── SECTION 11: Regressions and report table generation ───────────────────── */
// Runs all regressions (original and winsorised outcomes) and auto-generates
// two LaTeX regression tables. Recompiling the report after re-running this
// script updates the tables — no manual transcription needed.
// Significance: * p<0.10  ** p<0.05  *** p<0.01 (robust HC1 standard errors).

// ── Helper: returns significance cell for one coefficient ─────────────────────
// Restores estimates `estname', tests `varname', and returns r(cell):
//   "+***" / "-**" / "ns" etc. — plain text, no dollar signs, safe for
//   use in file write without triggering Stata global-macro expansion.

capture program drop sig_cell
program define sig_cell, rclass
    args estname varname
    capture quietly estimates restore `estname'
    if _rc {
        return local cell "ns"
        exit
    }
    capture quietly test `varname'
    if _rc {
        return local cell "ns"
        exit
    }
    local p = r(p)
    local b = _b[`varname']
    local pm = cond(`b' >= 0, "+", "-")
    if `p' >= 0.10 {
        return local cell "ns"
        exit
    }
    local stars = cond(`p' < 0.01, "***", cond(`p' < 0.05, "**", "*"))
    return local cell "`pm'`stars'"
end

// ── Variable display labels ────────────────────────────────────────────────────
local lab_jobfactors_profdev          "Professional development"
local lab_jobfactors_familyopinion    "Family opinion"
local lab_jobfactors_peeropinion      "Peer opinion"
local lab_jobfactors_location         "Location"
local lab_jobfactors_hoursflexibility "Hours flexibility"
local lab_jobfactors_qualifications   "Qualifications match"
local lab_jobfactors_promotion        "Promotion prospects"
local lab_jobfactors_training         "Training"

local regvars "jobfactors_profdev jobfactors_familyopinion jobfactors_peeropinion jobfactors_location jobfactors_hoursflexibility jobfactors_qualifications jobfactors_promotion jobfactors_training"

// ── Regressions: original wages ────────────────────────────────────────────────

reg wageWTAmin_monthly jobfactors*, robust
estimates store m1

reg wageWTAmin_monthly age female jobfactors*, robust
estimates store m2

reg wageWTAmax_monthly jobfactors*, robust
// N = 44; four respondents provided a minimum but not a maximum wage.
estimates store m3

reg wageWTAmax_monthly age female jobfactors*, robust
estimates store m4

// ── Regressions: winsorised wages ──────────────────────────────────────────────

reg wageWTAmin_monthly_w jobfactors*, robust
estimates store m1w

reg wageWTAmin_monthly_w age female jobfactors*, robust
estimates store m2w

reg wageWTAmax_monthly_w jobfactors*, robust
estimates store m3w

reg wageWTAmax_monthly_w age female jobfactors*, robust
estimates store m4w

// ── Write regression tables ────────────────────────────────────────────────────
// Stata local macro names are limited to 32 characters. Variable names such as
// jobfactors_hoursflexibility (27 chars) would exceed that limit if used as part
// of a cell key (e.g. cell_jobfactors_hoursflexibility_m1w = 36 chars). Instead,
// cells are stored using numeric indices: c_i_j where i = var index (1-8) and
// j = model index (1-4). Maximum name length is 6 chars (e.g. c_8_4).

// Map numeric index to full variable name (for sig_cell and label lookups)
local var1 jobfactors_profdev
local var2 jobfactors_familyopinion
local var3 jobfactors_peeropinion
local var4 jobfactors_location
local var5 jobfactors_hoursflexibility
local var6 jobfactors_qualifications
local var7 jobfactors_promotion
local var8 jobfactors_training

// The forvalues loop (s=1: original wages; s=2: winsorised wages) generates
// regtable_table.tex and regtable_w_table.tex respectively.

forvalues s = 1/2 {

    // Table-specific settings
    if `s' == 1 {
        local outfile "../Report/regtable_table.tex"
        local caption "Summary of OLS Regression Results"
        local tlabel  "tab:regs"
        local m1_ m1
        local m2_ m2
        local m3_ m3
        local m4_ m4
    }
    else {
        local outfile "../Report/regtable_w_table.tex"
        local caption "OLS Regression Results: Winsorized WTA (5th/95th Percentile)"
        local tlabel  "tab:regsw"
        local m1_ m1w
        local m2_ m2w
        local m3_ m3w
        local m4_ m4w
    }

    // Compute all 32 significance cells using numeric indices i (var) and j (model)
    // c_i_j stores "+***", "-**", or "ns" — all well within the 32-char name limit.
    forvalues i = 1/8 {
        local v = "`var`i''"
        forvalues j = 1/4 {
            local mj = "`m`j'_'"
            sig_cell `mj' `v'
            local c_`i'_`j' = r(cell)
        }
    }

    // Retrieve N for each model using the same j index
    forvalues j = 1/4 {
        local mj = "`m`j'_'"
        quietly estimates restore `mj'
        local N_`j' = e(N)
    }

    // Partition variable indices into significant (any model) and non-significant
    local sig_idx ""
    local nonsig_idx ""
    forvalues i = 1/8 {
        local any = 0
        forvalues j = 1/4 {
            if "`c_`i'_`j''" != "ns" {
                local any = 1
            }
        }
        if `any' {
            local sig_idx "`sig_idx' `i'"
        }
        else {
            local nonsig_idx "`nonsig_idx' `i'"
        }
    }

    // Write the complete table file
    file open tf using "`outfile'", write replace
    file write tf "\begin{table}[h!]" _n
    file write tf "\centering" _n
    file write tf "\begin{threeparttable}" _n
    file write tf "\caption{`caption'}" _n
    file write tf "\label{`tlabel'}" _n
    file write tf "\begin{tabular}{l*{4}{>{\centering\arraybackslash}p{1.6cm}}}" _n
    file write tf "\toprule" _n
    file write tf " & \multicolumn{2}{c}{\textbf{Min WTA}} & \multicolumn{2}{c}{\textbf{Max WTA}} \\" _n
    file write tf "\cmidrule(lr){2-3}\cmidrule(lr){4-5}" _n
    file write tf "\textbf{Job factor} & (1) & (2) & (3) & (4) \\" _n
    file write tf "\midrule" _n
    file write tf "\multicolumn{5}{l}{\textit{Significant predictors}} \\[2pt]" _n

    // Write each row cell-by-cell. _char(36) writes a literal $ to produce
    // $+$*** / $-$*** formatting (math-mode sign + superscript stars).

    local n_sig : word count `sig_idx'
    local k = 0
    foreach i of local sig_idx {
        local k = `k' + 1
        local v = "`var`i''"
        local lab "`lab_`v''"
        file write tf "`lab'"
        forvalues j = 1/4 {
            file write tf " & "
            if "`c_`i'_`j''" == "ns" {
                file write tf "ns"
            }
            else {
                local sg = substr("`c_`i'_`j''", 1, 1)
                local st = substr("`c_`i'_`j''", 2, .)
                file write tf _char(36) "`sg'" _char(36) "`st'"
            }
        }
        if `k' == `n_sig' {
            file write tf " \\[4pt]" _n
        }
        else {
            file write tf " \\" _n
        }
    }

    if "`nonsig_idx'" != "" {
        file write tf "\multicolumn{5}{l}{\textit{Non-significant predictors}} \\[2pt]" _n
        foreach i of local nonsig_idx {
            local v = "`var`i''"
            local lab "`lab_`v''"
            file write tf "`lab'"
            forvalues j = 1/4 {
                file write tf " & "
                if "`c_`i'_`j''" == "ns" {
                    file write tf "ns"
                }
                else {
                    local sg = substr("`c_`i'_`j''", 1, 1)
                    local st = substr("`c_`i'_`j''", 2, .)
                    file write tf _char(36) "`sg'" _char(36) "`st'"
                }
            }
            file write tf " \\" _n
        }
    }

    file write tf "\midrule" _n
    file write tf "Controls (age, female) & & \checkmark & & \checkmark \\" _n
    file write tf _char(36) "N" _char(36) " & `N_1' & `N_2' & `N_3' & `N_4' \\" _n
    file write tf "\bottomrule" _n
    file write tf "\end{tabular}" _n
    file write tf "\begin{tablenotes}[flushleft]" _n
    file write tf "\footnotesize" _n
    file write tf "\item \textit{Note:} " _char(36) "+" _char(36) " / " ///
                  _char(36) "-" _char(36) " indicates coefficient direction; ns = not significant. *** " ///
                  _char(36) "p" _char(36) "<0.01, ** " _char(36) "p" _char(36) ///
                  "<0.05 (robust standard errors). Salary excluded from all models" ///
                  " (see Section~\ref{sec:notes})." _n
    file write tf "\end{tablenotes}" _n
    file write tf "\end{threeparttable}" _n
    file write tf "\end{table}" _n
    file close tf

}

log close
