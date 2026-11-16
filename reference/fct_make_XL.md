# Export the inputs and all calculations to one Excel workbook with live formulas

Writes a single XLSX file where the calculation chain of mocaredd is
reproduced as Excel formulas, so every result can be traced back to the
input cells and recalculated in Excel when an input changes. The
workbook contains:

- the input tabs (`setup`, `time`, `area`, `carbon`, or the v1
  equivalents) and any other tab of the input file, unchanged;

- **Calc EF**: carbon elements aggregated into carbon stocks per land
  use, then into emission factors per land use transition, with value,
  standard error (SE) and uncertainty (U\\

- **Calc ERRs**: activity data and emission factors aggregated into
  emissions per transition, per time period (and REDD+ activity),
  reference and monitoring levels, and emission reductions, with value,
  SE and U\\

- **Calc MCS**: one row per source of simulations, in blocks (AD, carbon
  elements, carbon stocks and EF, E_REF / E_MON / ER_MON), with the
  input or Approach 1 result (value, SE, U\\ median, confidence interval
  and U\\

- **Calc Formulas**: every formula used, one per row, with a short
  description;

- **Sims**: the Monte Carlo simulations used by **Calc MCS** (one column
  per row of **Calc MCS**), written as values.

The Approach 1 formulas follow
[`fct_arithmetic_mean3()`](https://openforis.github.io/mocaredd/reference/fct_arithmetic_mean3.md),
so after recalculation the workbook gives the same carbon stocks,
emission factors, emissions and emission reductions. Simulations are
random draws and are written as values: the simulation statistics in
**Calc MCS** are formulas on the **Sims** tab, but they do not change
when an input changes.

U\\ \\z\\ the two-sided normal quantile for `setup$conf_level`. The
simulation U\\
[`fct_calc_res()`](https://openforis.github.io/mocaredd/reference/fct_calc_res.md).

Formulas are written without cached results. Excel and LibreOffice
calculate them when the file is opened, but tools that read cached
values (e.g. `readxl`) see empty cells until the file has been opened
and saved.

## Usage

``` r
fct_make_XL(
  .checked_data,
  .mcs = NULL,
  .path_in = NULL,
  .path_out,
  .sims_sheet = TRUE
)
```

## Arguments

- .checked_data:

  The list returned by
  [`fct_checkinput()`](https://openforis.github.io/mocaredd/reference/fct_checkinput.md).

- .mcs:

  Optional list of simulations, as `rv$mcs` in the app: `sim_trans`
  (output of
  [`fct_combine_mcs_E()`](https://openforis.github.io/mocaredd/reference/fct_combine_mcs_E.md))
  and, optionally, `sim_REF`, `sim_MON`
  ([`fct_combine_mcs_P()`](https://openforis.github.io/mocaredd/reference/fct_combine_mcs_P.md))
  and `sim_ER`
  ([`fct_combine_mcs_ER()`](https://openforis.github.io/mocaredd/reference/fct_combine_mcs_ER.md)).
  Missing period level simulations are calculated from `sim_trans`. If
  `NULL`, the simulation columns of **Calc MCS** are left empty.

- .path_in:

  Path to the input XLSX file that was checked. All its tabs are copied
  and the calculation tabs refer to its cells. If `NULL`, the four
  tables of `.checked_data` are written instead.

- .path_out:

  Path of the XLSX file to write.

- .sims_sheet:

  If `TRUE` (default), simulations are written to the **Sims** tab and
  their statistics are Excel formulas. If `FALSE`, the statistics are
  written as values and the **Sims** tab is omitted (smaller file).

## Value

The path of the XLSX file, invisibly.

## Examples

``` r
if (FALSE) { # \dontrun{
path    <- system.file("extdata/mocaredd-templatev2-simple.xlsx", package = "mocaredd")
checked <- fct_checkinput(.path = path)

set.seed(checked$data$setup$ran_seed)
sim_trans <- fct_combine_mcs_E(.checked_data = checked)

fct_make_XL(
  .checked_data = checked,
  .mcs          = list(sim_trans = sim_trans),
  .path_in      = path,
  .path_out     = tempfile(fileext = ".xlsx")
)
} # }
```
