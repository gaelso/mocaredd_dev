#' Export the inputs and all calculations to one Excel workbook with live formulas
#'
#' @description
#' Writes a single XLSX file where the calculation chain of mocaredd is reproduced
#' as Excel formulas, so every result can be traced back to the input cells and
#' recalculated in Excel when an input changes. The workbook contains:
#' \itemize{
#'   \item the input tabs (\code{setup}, \code{time}, \code{area}, \code{carbon},
#'         or the v1 equivalents) and any other tab of the input file, unchanged;
#'   \item \strong{Calc EF}: carbon elements aggregated into carbon stocks per land
#'         use, then into emission factors per land use transition, with value,
#'         standard error (SE) and uncertainty (U\%), IPCC Approach 1;
#'   \item \strong{Calc ERRs}: activity data and emission factors aggregated into
#'         emissions per transition, per time period (and REDD+ activity), reference
#'         and monitoring levels, and emission reductions, with value, SE and U\%;
#'   \item \strong{Calc MCS}: one row per source of simulations, in blocks (AD,
#'         carbon elements, carbon stocks and EF, E_REF / E_MON / ER_MON), with the
#'         input or Approach 1 result (value, SE, U\%) next to the simulation
#'         median, confidence interval and U\%;
#'   \item \strong{Calc Formulas}: every formula used, one per row, with a short
#'         description;
#'   \item \strong{Sims}: the Monte Carlo simulations used by \strong{Calc MCS}
#'         (one column per row of \strong{Calc MCS}), written as values.
#' }
#'
#' The Approach 1 formulas follow \code{fct_arithmetic_mean3()}, so after
#' recalculation the workbook gives the same carbon stocks, emission factors,
#' emissions and emission reductions. Simulations are random draws and are
#' written as values: the simulation statistics in \strong{Calc MCS} are formulas
#' on the \strong{Sims} tab, but they do not change when an input changes.
#'
#' U\% is \eqn{z \cdot SE / |X|}, as a fraction formatted in percent, with
#' \eqn{z} the two-sided normal quantile for \code{setup$conf_level}. The
#' simulation U\% is \eqn{(upper - lower) / 2 / |median|}, as in
#' \code{fct_calc_res()}.
#'
#' Formulas are written without cached results. Excel and LibreOffice calculate
#' them when the file is opened, but tools that read cached values (e.g.
#' \code{readxl}) see empty cells until the file has been opened and saved.
#'
#' @param .checked_data The list returned by \code{fct_checkinput()}.
#' @param .mcs Optional list of simulations, as \code{rv$mcs} in the app:
#'             \code{sim_trans} (output of \code{fct_combine_mcs_E()}) and,
#'             optionally, \code{sim_REF}, \code{sim_MON} (\code{fct_combine_mcs_P()})
#'             and \code{sim_ER} (\code{fct_combine_mcs_ER()}). Missing period
#'             level simulations are calculated from \code{sim_trans}. If
#'             \code{NULL}, the simulation columns of \strong{Calc MCS} are left empty.
#' @param .path_in Path to the input XLSX file that was checked. All its tabs are
#'             copied and the calculation tabs refer to its cells. If \code{NULL},
#'             the four tables of \code{.checked_data} are written instead.
#' @param .path_out Path of the XLSX file to write.
#' @param .sims_sheet If \code{TRUE} (default), simulations are written to the
#'             \strong{Sims} tab and their statistics are Excel formulas. If
#'             \code{FALSE}, the statistics are written as values and the
#'             \strong{Sims} tab is omitted (smaller file).
#'
#' @return The path of the XLSX file, invisibly.
#'
#' @importFrom rlang .data
#'
#' @examples
#' \dontrun{
#' path    <- system.file("extdata/mocaredd-templatev2-simple.xlsx", package = "mocaredd")
#' checked <- fct_checkinput(.path = path)
#'
#' set.seed(checked$data$setup$ran_seed)
#' sim_trans <- fct_combine_mcs_E(.checked_data = checked)
#'
#' fct_make_XL(
#'   .checked_data = checked,
#'   .mcs          = list(sim_trans = sim_trans),
#'   .path_in      = path,
#'   .path_out     = tempfile(fileext = ".xlsx")
#' )
#' }
#'
#' @export
fct_make_XL <- function(.checked_data, .mcs = NULL, .path_in = NULL, .path_out, .sims_sheet = TRUE){

  ## !!! FOR TESTING ONLY
  # .checked_data <- fct_checkinput(.path = path)
  # .mcs <- list(sim_trans = fct_combine_mcs_E(.checked_data))
  # .path_in <- path ; .path_out <- "test.xlsx" ; .sims_sheet <- TRUE
  ## !!!

  setup  <- .checked_data$data$setup
  time   <- .checked_data$data$time
  area   <- .checked_data$data$area
  carbon <- .checked_data$data$carbon
  is_v2  <- isTRUE(.checked_data$template_version == 2)

  if (any(carbon$c_period != "ALL")) {
    stop("fct_make_XL() requires c_period = 'ALL' for all carbon elements, as fct_combine_mcs_E().")
  }

  ## TOC ####
  ## 1. Workbook, input tabs and cell references of the inputs
  ## 2. Preparation: carbon elements, carbon fraction, degradation (as fct_arithmetic_mean3())
  ## 3. Calc EF: carbon stocks per land use, emission factors per land use pair
  ## 4. Calc ERRs: emissions per transition, period, levels and emission reductions
  ## 5. Calc MCS (and Sims): inputs and Approach 1 results vs. simulations
  ## 6. Calc Formulas
  ## 7. Save


  ##
  ## 1. Workbook and input tabs ################################################
  ##

  in_sheet <- if (is_v2) {
    c(setup = "setup", time = "time", area = "area", carbon = "carbon")
  } else {
    c(setup = "user_inputs", time = "time_periods", area = "AD_lu_transitions", carbon = "c_stocks")
  }

  sh <- c(ef = "Calc EF", err = "Calc ERRs", mcs = "Calc MCS", form = "Calc Formulas", sims = "Sims")

  if (!is.null(.path_in)) {
    wb <- openxlsx2::wb_load(.path_in)
  } else {
    wb <- openxlsx2::wb_workbook()
    for (s in names(in_sheet)) {
      wb$add_worksheet(in_sheet[[s]])
      wb$add_data(sheet = in_sheet[[s]], x = .checked_data$data[[s]], na.strings = NULL)
    }
  }

  ## Re-running on an output file: drop previous calculation tabs and names
  for (s in intersect(sh, wb$get_sheet_names())) wb$remove_worksheet(sheet = s)
  ## Workbook names used by the formulas; existing ones (e.g. #REF! leftovers in
  ## the templates) are overwritten in step 3.1
  xl_names <- c("conf_level", "z_score", "ad_annual")

  ## Position of each table in its tab: column letter per column name, Excel row per table row
  key_col <- c(setup = "conf_level", time = "year_start", area = "trans_area", carbon = "c_value")
  in_map  <- purrr::imap(in_sheet, function(s, tab) {
    xl_map_table(wb, s, .checked_data$data[[tab]], key_col[[tab]])
  })

  ## Absolute reference to an input cell, e.g. 'carbon'!$F$5
  in_ref <- function(tab, col, i) {
    m <- in_map[[tab]]
    if (!col %in% names(m$col)) stop("Column '", col, "' not found in tab '", in_sheet[[tab]], "'.")
    paste0(xl_sheet(in_sheet[[tab]]), "!$", m$col[[col]], "$", m$row[i])
  }


  ##
  ## 2. Preparation (mirrors fct_arithmetic_mean3()) ###########################
  ##

  ## + 2.1. Carbon elements, with their row in the input tab ####
  carbon <- carbon |>
    dplyr::mutate(
      xl_i   = dplyr::row_number(),
      c_unit = if (is_v2) .data$c_unit else setup$c_unit,
      c_lu   = if ("c_lu" %in% names(carbon)) .data$c_lu else .data$c_lu_id
    )
  carbon_pools <- carbon |> dplyr::filter(.data$c_element != "CF")

  ## + 2.2. Carbon fraction: v2 'carbon' row, v1 'user_inputs' ####
  if (is_v2) {
    i_cf   <- which(carbon$c_element == "CF")
    has_cf <- length(i_cf) > 0
    cf_ref <- if (has_cf) c(value = in_ref("carbon", "c_value", i_cf[1]), se = in_ref("carbon", "c_se", i_cf[1]))
  } else {
    has_cf <- isTRUE(setup$c_unit == "DM") && !is.null(setup$c_fraction)
    cf_ref <- if (has_cf) c(value = in_ref("setup", "c_fraction", 1), se = in_ref("setup", "c_fraction_se", 1))
  }

  ## + 2.3. Degradation: degraded land uses, their intact land use, pools affected ####
  lu_dg  <- unique(carbon_pools$c_lu_id[carbon_pools$c_element == "DG_ratio"])
  has_dg <- length(lu_dg) > 0
  lu_intact_of <- if (has_dg) {
    out <- stringr::str_remove(lu_dg, setup$dg_ext)
    if (is_v2) out <- stringr::str_replace(out, "[^A-Za-z]$", "")
    stats::setNames(out, lu_dg)
  } else {
    character(0)
  }
  dg_pool    <- stringr::str_trim(stringr::str_split(setup$dg_pool, ",")[[1]])
  dg_partial <- has_dg && !identical(dg_pool, "ALL")

  ## + 2.4. Land uses and elements used as columns ####
  lu_tbl  <- carbon_pools |> dplyr::distinct(.data$c_lu_id, .keep_all = TRUE) |> dplyr::select("c_lu_id", "c_lu")
  el_cols <- intersect(c("AGB", "BGB", "RS", "DW", "LI", "SOC", "ALL", "DG_ratio"), unique(carbon_pools$c_element))

  ## + 2.5. Periods and period types ####
  periods   <- time |> dplyr::filter(.data$period_no %in% area$trans_period)
  ptypes    <- unique(stats::na.omit(periods$period_type))
  ref_types <- ptypes[stringr::str_detect(ptypes, "REF")]
  mon_types <- ptypes[stringr::str_detect(ptypes, "MON")]
  ptype_of  <- stats::setNames(time$period_type, time$period_no)

  ## Formulas listed in 'Calc Formulas', filled along the way
  form_rows <- list()
  add_form  <- function(step, formula, description, sheet, example = NA_character_) {
    form_rows[[length(form_rows) + 1]] <<- dplyr::tibble(
      Step = step, Formula = formula, Description = description,
      `Calculated in` = sheet, `Example (Excel)` = example
    )
  }


  ##
  ## 3. Calc EF ###############################################################
  ##

  wb$add_worksheet(sh[["ef"]])
  xl_title(wb, sh[["ef"]],
    "Calc EF - Carbon stocks and emission factors (IPCC Approach 1)",
    "Carbon elements are aggregated into carbon stocks per land use (tC/ha), then into emission factors per land use transition (tCO2e/ha)."
  )

  ## + 3.1. Settings, with workbook names used by all formulas ####
  set_df <- dplyr::tibble(
    Parameter = c("Confidence level", "z score", "AD annual"),
    Value     = c(paste0("=", in_ref("setup", "conf_level", 1)),
                  "=NORMSINV(1-(1-conf_level)/2)",
                  paste0("=", in_ref("setup", "ad_annual", 1))),
    Name      = xl_names,
    Note      = c("From the setup tab.",
                  "Two-sided normal quantile for the confidence level.",
                  "From the setup tab. TRUE if activity data are annual, FALSE if they are period totals.")
  )
  b_set <- xl_block(wb, sh[["ef"]], "Settings", set_df, start_row = 5,
                    num_fmt = c(Value = "0.000"), input_cols = "Value")
  for (k in seq_along(xl_names)) {
    wb$add_named_region(sheet = sh[["ef"]], dims = paste0("B", b_set$rows[k]), name = xl_names[k], overwrite = TRUE)
  }
  add_form("Settings", "z_score = NORMSINV(1 - (1 - conf_level) / 2)",
           "Two-sided normal quantile for the confidence level of the setup tab.", sh[["ef"]],
           paste0(xl_sheet(sh[["ef"]]), "!B", b_set$rows[2], ": ", set_df$Value[2]))
  add_form("Settings", "U% = z_score * SE / |X|",
           "Uncertainty: half-width of the confidence interval relative to the estimate X, shown in percent.",
           paste(sh[c("ef", "err", "mcs")], collapse = ", "))
  add_form("Uncertainty rules", "X = A * B  ->  SE_X = sqrt((B * SE_A)^2 + (A * SE_B)^2) = X * sqrt((SE_A / A)^2 + (SE_B / B)^2)",
           "IPCC Approach 1 (2006 GL, Vol. 1, Ch. 3, Eq. 3.1): product of independent terms, relative errors combine in quadrature.", "All")
  add_form("Uncertainty rules", "X = A + B (or A - B)  ->  SE_X = sqrt(SE_A^2 + SE_B^2)",
           "IPCC Approach 1 (2006 GL, Vol. 1, Ch. 3, Eq. 3.2): sum or difference of independent terms, absolute errors combine in quadrature.", "All")
  add_form("Uncertainty rules", "Shared inputs are treated as independent",
           "Approach 1 ignores correlations from shared inputs (e.g. CF, a carbon stock used in several transitions); the Monte Carlo simulations (Approach 2) account for them.", "All")

  ## + 3.2. Carbon stock per land use ####
  cs_cols <- c("lu_id", "land use", as.vector(rbind(el_cols, paste0(el_cols, "_se"))))
  if (has_cf) cs_cols <- c(cs_cols, "CF", "CF_se")
  if (has_dg) cs_cols <- c(cs_cols, "lu_intact", "C_intact_DG", "C_intact_DG_se", "C_intact_noDG", "C_intact_noDG_se")
  cs_cols <- c(cs_cols, "C", "C_se", "C_U", "formula C", "formula C_se")

  cs_start <- b_set$next_row
  cs_r1    <- cs_start + 2
  cs_row   <- stats::setNames(cs_r1 + seq_len(nrow(lu_tbl)) - 1, lu_tbl$c_lu_id)
  cs_let   <- stats::setNames(openxlsx2::int2col(seq_along(cs_cols)), cs_cols)
  cs_cell  <- function(col, lu) paste0(cs_let[[col]], cs_row[[lu]])

  cs_df <- matrix(NA_character_, nrow = nrow(lu_tbl), ncol = length(cs_cols), dimnames = list(NULL, cs_cols)) |>
    dplyr::as_tibble()
  cs_df$lu_id      <- lu_tbl$c_lu_id
  cs_df$`land use` <- lu_tbl$c_lu

  cs_forms <- list()  ## formula text per land use, for 'Calc Formulas'

  ## Carbon stock formula from a set of elements of a land use
  make_forms <- function(els, units) {
    if (length(els) == 0) return(c(C = "0", C_se = "0"))
    c(
      C    = fct_make_formula2( .c_el = els, .c_unit = if (is_v2) units else setup$c_unit, .version = if (is_v2) 2 else 1),
      C_se = fct_make_formula_U(.c_el = els, .c_unit = if (is_v2) units else setup$c_unit, .version = if (is_v2) 2 else 1)
    )
  }

  for (i in seq_len(nrow(lu_tbl))) {
    lu  <- lu_tbl$c_lu_id[i]
    els <- carbon_pools |> dplyr::filter(.data$c_lu_id == lu)

    ## Input values and SE
    for (k in seq_len(nrow(els))) {
      cs_df[[els$c_element[k]]][i]                <- paste0("=", in_ref("carbon", "c_value", els$xl_i[k]))
      cs_df[[paste0(els$c_element[k], "_se")]][i] <- paste0("=", in_ref("carbon", "c_se",    els$xl_i[k]))
    }

    varmap <- c(
      stats::setNames(vapply(els$c_element, function(e) cs_cell(e, lu), character(1)), els$c_element),
      stats::setNames(vapply(els$c_element, function(e) cs_cell(paste0(e, "_se"), lu), character(1)), paste0(els$c_element, "_se"))
    )
    if (has_cf) varmap <- c(varmap, CF = cs_cell("CF", lu), CF_se = cs_cell("CF_se", lu))

    if (!lu %in% lu_dg) {

      ## Intact or non-degraded land use: formulas from fct_make_formula2() / fct_make_formula_U()
      pools <- els |> dplyr::filter(.data$c_element != "DG_ratio")
      f     <- make_forms(pools$c_element, pools$c_unit)
      uses_cf <- any(stringr::str_detect(f, "\\bCF\\b"))
      if (uses_cf && !has_cf) stop("Land use '", lu, "' has elements in dry matter but no carbon fraction (CF).")
      if (uses_cf) {
        cs_df$CF[i]    <- paste0("=", cf_ref[["value"]])
        cs_df$CF_se[i] <- paste0("=", cf_ref[["se"]])
      }
      cs_df$C[i]    <- paste0("=", xl_translate(f[["C"]],    varmap))
      cs_df$C_se[i] <- paste0("=", xl_translate(f[["C_se"]], varmap))
      cs_df$`formula C`[i]    <- paste0("C = ", f[["C"]])
      cs_df$`formula C_se`[i] <- paste0("C_se = ", f[["C_se"]])
      cs_forms[[lu]] <- c(type = "stock", f)

      ## Intact land use of a degraded land use, only some pools degraded:
      ## stock of the degraded pools (RS follows AGB) and of the other pools
      if (dg_partial && lu %in% lu_intact_of) {
        is_dgp <- pools$c_element %in% dg_pool | (pools$c_element == "RS" & "AGB" %in% dg_pool)
        f_dg   <- make_forms(pools$c_element[is_dgp],  pools$c_unit[is_dgp])
        f_nodg <- make_forms(pools$c_element[!is_dgp], pools$c_unit[!is_dgp])
        cs_df$C_intact_DG[i]      <- paste0("=", xl_translate(f_dg[["C"]],      varmap))
        cs_df$C_intact_DG_se[i]   <- paste0("=", xl_translate(f_dg[["C_se"]],   varmap))
        cs_df$C_intact_noDG[i]    <- paste0("=", xl_translate(f_nodg[["C"]],    varmap))
        cs_df$C_intact_noDG_se[i] <- paste0("=", xl_translate(f_nodg[["C_se"]], varmap))
        cs_forms[[paste0(lu, "..DG")]]   <- c(type = "intact_DG",   f_dg)
        cs_forms[[paste0(lu, "..noDG")]] <- c(type = "intact_noDG", f_nodg)
      }

    } else {

      ## Degraded land use: DG_ratio * C_intact_DG + C_intact_noDG
      intact <- lu_intact_of[[lu]]
      cs_df$lu_intact[i] <- intact
      if (dg_partial) {
        cs_df$C_intact_DG[i]      <- paste0("=", cs_cell("C_intact_DG",      intact))
        cs_df$C_intact_DG_se[i]   <- paste0("=", cs_cell("C_intact_DG_se",   intact))
        cs_df$C_intact_noDG[i]    <- paste0("=", cs_cell("C_intact_noDG",    intact))
        cs_df$C_intact_noDG_se[i] <- paste0("=", cs_cell("C_intact_noDG_se", intact))
      } else {
        cs_df$C_intact_DG[i]      <- paste0("=", cs_cell("C",    intact))
        cs_df$C_intact_DG_se[i]   <- paste0("=", cs_cell("C_se", intact))
        cs_df$C_intact_noDG[i]    <- "=0"
        cs_df$C_intact_noDG_se[i] <- "=0"
      }
      f <- c(
        C    = "DG_ratio * C_intact_DG + C_intact_noDG",
        C_se = "sqrt((C_intact_DG * DG_ratio_se)^2 + (DG_ratio * C_intact_DG_se)^2 + C_intact_noDG_se^2)"
      )
      varmap_dg <- c(
        DG_ratio = cs_cell("DG_ratio", lu), DG_ratio_se = cs_cell("DG_ratio_se", lu),
        C_intact_DG   = cs_cell("C_intact_DG", lu),   C_intact_DG_se   = cs_cell("C_intact_DG_se", lu),
        C_intact_noDG = cs_cell("C_intact_noDG", lu), C_intact_noDG_se = cs_cell("C_intact_noDG_se", lu)
      )
      cs_df$C[i]    <- paste0("=", xl_translate(f[["C"]],    varmap_dg))
      cs_df$C_se[i] <- paste0("=", xl_translate(f[["C_se"]], varmap_dg))
      cs_df$`formula C`[i]    <- paste0("C = ", f[["C"]])
      cs_df$`formula C_se`[i] <- paste0("C_se = ", f[["C_se"]])
      cs_forms[[lu]] <- c(type = "degraded", f)
    }

    cs_df$C_U[i] <- xl_u(cs_cell("C", lu), cs_cell("C_se", lu))
  }

  b_cs <- xl_block(
    wb, sh[["ef"]], "1. Carbon stock per land use (tC/ha)", cs_df, start_row = cs_start,
    num_fmt    = c(stats::setNames(rep("#,##0.000", length(cs_cols)), cs_cols), C_U = "0.0%"),
    input_cols = c(el_cols, paste0(el_cols, "_se"), if (has_cf) c("CF", "CF_se"))
  )

  ## Carbon stock formulas for 'Calc Formulas', one row per distinct formula
  cs_ftbl <- dplyr::tibble(
    lu   = stringr::str_remove(names(cs_forms), "\\.\\.(noDG|DG)$"),
    type = purrr::map_chr(cs_forms, "type"),
    C    = purrr::map_chr(cs_forms, "C"),
    C_se = purrr::map_chr(cs_forms, "C_se"),
    col  = dplyr::case_when(.data$type == "intact_DG" ~ "C_intact_DG", .data$type == "intact_noDG" ~ "C_intact_noDG", TRUE ~ "C")
  )
  cs_type_txt <- c(
    stock       = "Carbon stock (tC/ha) of",
    intact_DG   = "Carbon stock of the pools affected by degradation (dg_pool) of",
    intact_noDG = "Carbon stock of the pools not affected by degradation of",
    degraded    = "Carbon stock of degraded land uses (DG_ratio of the intact stock, plus pools not degraded):"
  )
  cs_ftbl |>
    dplyr::summarise(lus = paste(.data$lu, collapse = ", "), lu1 = dplyr::first(.data$lu),
                     .by = c("type", "C", "C_se", "col")) |>
    purrr::pwalk(function(type, C, C_se, col, lus, lu1) {
      lhs  <- if (type == "stock" || type == "degraded") "C" else col
      desc <- paste(cs_type_txt[[type]], lus)
      add_form("Carbon stock", paste0(lhs, " = ", C), desc, sh[["ef"]],
               paste0(xl_sheet(sh[["ef"]]), "!", cs_cell(col, lu1), ": ", cs_df[[col]][match(lu1, cs_df$lu_id)]))
      add_form("Carbon stock", paste0(lhs, "_se = ", C_se), paste("Standard error (Approach 1) -", tolower(desc)), sh[["ef"]],
               paste0(xl_sheet(sh[["ef"]]), "!", cs_cell(paste0(col, "_se"), lu1), ": ", cs_df[[paste0(col, "_se")]][match(lu1, cs_df$lu_id)]))
    })
  add_form("Carbon stock", "C_U = z_score * C_se / |C|", "Uncertainty of the carbon stock.", sh[["ef"]],
           paste0(xl_sheet(sh[["ef"]]), "!", cs_cell("C_U", lu_tbl$c_lu_id[1]), ": ", cs_df$C_U[1]))
  if (has_cf) {
    add_form("Carbon stock", "CF applies to elements in dry matter (DM); RS takes the unit of AGB",
             "Elements in carbon (C) are added without CF. In formulas, AGB * (1 + RS) is AGB plus BGB derived from the root-to-shoot ratio.", sh[["ef"]])
  }

  ## + 3.3. Emission factor per land use pair ####
  ef_pairs <- area |>
    dplyr::distinct(.data$lu_initial_id, .data$lu_final_id) |>
    dplyr::mutate(
      ef_id       = paste(.data$lu_initial_id, .data$lu_final_id, sep = "*"),
      is_dg_trans = .data$lu_final_id %in% lu_dg &
        .data$lu_initial_id == dplyr::coalesce(unname(lu_intact_of[.data$lu_final_id]), "")
    )

  ef_cols <- c("EF_id", "lu_initial_id", "lu_final_id", "method", "C_i", "C_i_se", "C_f", "C_f_se")
  if (has_dg) ef_cols <- c(ef_cols, "DG_ratio", "DG_ratio_se", "C_intact_DG", "C_intact_DG_se")
  ef_cols <- c(ef_cols, "dC", "dC_se", "EF", "EF_se", "EF_U")

  ef_start <- b_cs$next_row
  ef_r1    <- ef_start + 2
  ef_row   <- stats::setNames(ef_r1 + seq_len(nrow(ef_pairs)) - 1, ef_pairs$ef_id)
  ef_let   <- stats::setNames(openxlsx2::int2col(seq_along(ef_cols)), ef_cols)
  ef_cell  <- function(col, id) paste0(ef_let[[col]], ef_row[[id]])

  ef_df <- matrix(NA_character_, nrow = nrow(ef_pairs), ncol = length(ef_cols), dimnames = list(NULL, ef_cols)) |>
    dplyr::as_tibble()

  for (i in seq_len(nrow(ef_pairs))) {
    id <- ef_pairs$ef_id[i]; li <- ef_pairs$lu_initial_id[i]; lf <- ef_pairs$lu_final_id[i]
    c_ <- function(col) ef_cell(col, id)
    ef_df$EF_id[i]         <- id
    ef_df$lu_initial_id[i] <- li
    ef_df$lu_final_id[i]   <- lf
    ef_df$C_i[i]    <- paste0("=", cs_cell("C",    li))
    ef_df$C_i_se[i] <- paste0("=", cs_cell("C_se", li))
    ef_df$C_f[i]    <- paste0("=", cs_cell("C",    lf))
    ef_df$C_f_se[i] <- paste0("=", cs_cell("C_se", lf))
    if (ef_pairs$is_dg_trans[i]) {
      ## Intact -> degraded: (1 - DG_ratio) * C_intact_DG, product rule, so the
      ## intact stock is not counted twice
      ef_df$method[i]         <- "degradation: (1 - DG_ratio) * C_intact_DG"
      ef_df$DG_ratio[i]       <- paste0("=", cs_cell("DG_ratio",       lf))
      ef_df$DG_ratio_se[i]    <- paste0("=", cs_cell("DG_ratio_se",    lf))
      ef_df$C_intact_DG[i]    <- paste0("=", cs_cell("C_intact_DG",    lf))
      ef_df$C_intact_DG_se[i] <- paste0("=", cs_cell("C_intact_DG_se", lf))
      ef_df$dC[i]    <- paste0("=(1-", c_("DG_ratio"), ")*", c_("C_intact_DG"))
      ef_df$dC_se[i] <- paste0("=SQRT((", c_("C_intact_DG"), "*", c_("DG_ratio_se"), ")^2+((1-", c_("DG_ratio"), ")*", c_("C_intact_DG_se"), ")^2)")
    } else {
      ef_df$method[i] <- "stock difference: C_i - C_f"
      ef_df$dC[i]     <- paste0("=", c_("C_i"), "-", c_("C_f"))
      ef_df$dC_se[i]  <- paste0("=SQRT(", c_("C_i_se"), "^2+", c_("C_f_se"), "^2)")
    }
    ef_df$EF[i]    <- paste0("=", c_("dC"),    "*44/12")
    ef_df$EF_se[i] <- paste0("=", c_("dC_se"), "*44/12")
    ef_df$EF_U[i]  <- xl_u(c_("EF"), c_("EF_se"))
  }

  b_ef <- xl_block(
    wb, sh[["ef"]], "2. Emission factor per land use transition (tCO2e/ha)", ef_df, start_row = ef_start,
    num_fmt = c(stats::setNames(rep("#,##0.000", length(ef_cols)), ef_cols), EF_U = "0.0%")
  )

  ef_ex <- function(col, dg) {
    i <- which(ef_pairs$is_dg_trans == dg)[1]
    if (is.na(i)) return(NA_character_)
    paste0(xl_sheet(sh[["ef"]]), "!", ef_cell(col, ef_pairs$ef_id[i]), ": ", ef_df[[col]][i])
  }
  add_form("Emission factor", "dC = C_i - C_f", "Carbon stock change between the initial (i) and final (f) land use (tC/ha).", sh[["ef"]], ef_ex("dC", FALSE))
  add_form("Emission factor", "dC_se = sqrt(C_i_se^2 + C_f_se^2)", "Standard error of the stock change (sum rule).", sh[["ef"]], ef_ex("dC_se", FALSE))
  if (any(ef_pairs$is_dg_trans)) {
    add_form("Emission factor", "dC = (1 - DG_ratio) * C_intact_DG",
             "Intact to degraded land use: loss of the degraded pools, written as a product so the intact stock is not counted twice.", sh[["ef"]], ef_ex("dC", TRUE))
    add_form("Emission factor", "dC_se = sqrt((C_intact_DG * DG_ratio_se)^2 + ((1 - DG_ratio) * C_intact_DG_se)^2)",
             "Standard error of the stock change for degradation (product rule).", sh[["ef"]], ef_ex("dC_se", TRUE))
  }
  add_form("Emission factor", "EF = dC * 44/12", "Emission factor (tCO2e/ha), 44/12 converts carbon to CO2.", sh[["ef"]], ef_ex("EF", FALSE))
  add_form("Emission factor", "EF_se = dC_se * 44/12", "Standard error of the emission factor.", sh[["ef"]], ef_ex("EF_se", FALSE))
  add_form("Emission factor", "EF_U = z_score * EF_se / |EF|", "Uncertainty of the emission factor.", sh[["ef"]], ef_ex("EF_U", FALSE))

  xl_finish_sheet(wb, sh[["ef"]], n_cols = max(length(cs_cols), length(ef_cols)), first_width = 30)


  ##
  ## 4. Calc ERRs ##############################################################
  ##

  wb$add_worksheet(sh[["err"]])
  xl_title(wb, sh[["err"]],
    "Calc ERRs - Emissions and emission reductions (IPCC Approach 1)",
    "Activity data (ha) and emission factors (tCO2e/ha) are combined into emissions (tCO2e/yr), aggregated to time periods, reference and monitoring levels, then emission reductions."
  )

  ## + 4.1. Emissions per transition ####
  tr_cols <- c("trans_id", "trans_period", "period_type", "redd_activity", "lu_initial_id", "lu_final_id",
               "period_length", "AD", "AD_se", "AD_U", "EF", "EF_se", "E", "E_se", "E_year", "E_year_se", "E_year_U")
  tr_start <- 5
  tr_r1    <- tr_start + 2
  tr_rows  <- tr_r1 + seq_len(nrow(area)) - 1
  tr_let   <- stats::setNames(openxlsx2::int2col(seq_along(tr_cols)), tr_cols)
  tr_rng   <- function(col) paste0("$", tr_let[[col]], "$", min(tr_rows), ":$", tr_let[[col]], "$", max(tr_rows))

  t_i <- match(area$trans_period, time$period_no)
  c_  <- function(col) paste0(tr_let[[col]], tr_rows)
  ef_ids <- paste(area$lu_initial_id, area$lu_final_id, sep = "*")

  tr_df <- dplyr::tibble(
    trans_id      = area$trans_id,
    trans_period  = area$trans_period,
    period_type   = dplyr::coalesce(ptype_of[area$trans_period], ""),
    redd_activity = area$redd_activity,
    lu_initial_id = area$lu_initial_id,
    lu_final_id   = area$lu_final_id,
    period_length = paste0("=", in_ref("time", "year_end", t_i), "-", in_ref("time", "year_start", t_i), "+1"),
    AD            = paste0("=", in_ref("area", "trans_area", seq_len(nrow(area)))),
    AD_se         = paste0("=", in_ref("area", "trans_se",   seq_len(nrow(area)))),
    AD_U          = xl_u(c_("AD"), c_("AD_se")),
    EF            = paste0("=", xl_sheet(sh[["ef"]]), "!", ef_let[["EF"]],    ef_row[ef_ids]),
    EF_se         = paste0("=", xl_sheet(sh[["ef"]]), "!", ef_let[["EF_se"]], ef_row[ef_ids]),
    E             = paste0("=", c_("AD"), "*", c_("EF")),
    E_se          = paste0("=SQRT((", c_("EF"), "*", c_("AD_se"), ")^2+(", c_("AD"), "*", c_("EF_se"), ")^2)"),
    E_year        = paste0("=IF(ad_annual,", c_("E"),    ",", c_("E"),    "/", c_("period_length"), ")"),
    E_year_se     = paste0("=IF(ad_annual,", c_("E_se"), ",", c_("E_se"), "/", c_("period_length"), ")"),
    E_year_U      = xl_u(c_("E_year"), c_("E_year_se"))
  )

  b_tr <- xl_block(
    wb, sh[["err"]], "1. Emissions per land use transition", tr_df, start_row = tr_start,
    num_fmt = c(period_length = "0", AD = "#,##0", AD_se = "#,##0", AD_U = "0.0%", EF = "#,##0.000", EF_se = "#,##0.000",
                E = "#,##0", E_se = "#,##0", E_year = "#,##0", E_year_se = "#,##0", E_year_U = "0.0%"),
    input_cols = c("AD", "AD_se")
  )
  tr_ex <- function(col) paste0(xl_sheet(sh[["err"]]), "!", tr_let[[col]], tr_r1, ": ", tr_df[[col]][1])
  add_form("Emissions", "period_length = year_end - year_start + 1", "Length of the time period (years), from the time tab.", sh[["err"]], tr_ex("period_length"))
  add_form("Emissions", "E = AD * EF", "Emissions of a land use transition (tCO2e), activity data (ha) times emission factor (tCO2e/ha).", sh[["err"]], tr_ex("E"))
  add_form("Emissions", "E_se = sqrt((EF * AD_se)^2 + (AD * EF_se)^2)", "Standard error of the emissions (product rule).", sh[["err"]], tr_ex("E_se"))
  add_form("Emissions", "E_year = IF(ad_annual, E, E / period_length)", "Annual emissions (tCO2e/yr): E if activity data are annual, otherwise E divided by the period length.", sh[["err"]], tr_ex("E_year"))
  add_form("Emissions", "E_year_se = IF(ad_annual, E_se, E_se / period_length)", "Standard error of the annual emissions.", sh[["err"]], tr_ex("E_year_se"))

  ## + 4.2. Emissions per time period and REDD+ activity ####
  pa <- area |>
    dplyr::distinct(.data$trans_period, .data$redd_activity) |>
    dplyr::arrange(match(.data$trans_period, time$period_no))
  pa_start <- b_tr$next_row
  pa_rows  <- pa_start + 2 + seq_len(nrow(pa)) - 1
  pa_df <- dplyr::tibble(
    trans_period  = pa$trans_period,
    period_type   = dplyr::coalesce(ptype_of[pa$trans_period], ""),
    redd_activity = pa$redd_activity,
    E_year        = paste0("=SUMIFS(", tr_rng("E_year"), ",", tr_rng("trans_period"), ",A", pa_rows, ",", tr_rng("redd_activity"), ",C", pa_rows, ")"),
    E_year_se     = paste0("=SQRT(SUMPRODUCT((", tr_rng("trans_period"), "=A", pa_rows, ")*(", tr_rng("redd_activity"), "=C", pa_rows, ")*", tr_rng("E_year_se"), "^2))"),
    E_year_U      = xl_u(paste0("D", pa_rows), paste0("E", pa_rows))
  )
  b_pa <- xl_block(wb, sh[["err"]], "2. Annual emissions per time period and REDD+ activity", pa_df, start_row = pa_start,
                   num_fmt = c(E_year = "#,##0", E_year_se = "#,##0", E_year_U = "0.0%"))

  ## + 4.3. Emissions per time period ####
  pe_start <- b_pa$next_row
  pe_rows  <- pe_start + 2 + seq_len(nrow(periods)) - 1
  pe_df <- dplyr::tibble(
    trans_period  = periods$period_no,
    period_type   = dplyr::coalesce(periods$period_type, ""),
    period_length = paste0("=", in_ref("time", "year_end", match(periods$period_no, time$period_no)), "-",
                           in_ref("time", "year_start", match(periods$period_no, time$period_no)), "+1"),
    E_year        = paste0("=SUMIFS(", tr_rng("E_year"), ",", tr_rng("trans_period"), ",A", pe_rows, ")"),
    E_year_se     = paste0("=SQRT(SUMPRODUCT((", tr_rng("trans_period"), "=A", pe_rows, ")*", tr_rng("E_year_se"), "^2))"),
    E_year_U      = xl_u(paste0("D", pe_rows), paste0("E", pe_rows))
  )
  b_pe <- xl_block(wb, sh[["err"]], "3. Annual emissions per time period", pe_df, start_row = pe_start,
                   num_fmt = c(period_length = "0", E_year = "#,##0", E_year_se = "#,##0", E_year_U = "0.0%"))
  pe_rng <- function(col) paste0("$", col, "$", min(pe_rows), ":$", col, "$", max(pe_rows))
  add_form("Aggregation", "E_period = sum(E_year)", "Annual emissions of a time period (and REDD+ activity): sum over its transitions.", sh[["err"]],
           paste0(xl_sheet(sh[["err"]]), "!D", pe_rows[1], ": ", pe_df$E_year[1]))
  add_form("Aggregation", "E_period_se = sqrt(sum(E_year_se^2))", "Standard error of the period emissions (sum rule).", sh[["err"]],
           paste0(xl_sheet(sh[["err"]]), "!E", pe_rows[1], ": ", pe_df$E_year_se[1]))

  ## + 4.4. Reference and monitoring levels ####
  lv_types <- c(ref_types, mon_types)
  lv_start <- b_pe$next_row
  lv_rows  <- stats::setNames(lv_start + 2 + seq_along(lv_types) - 1, lv_types)
  lv_df <- dplyr::tibble(
    result       = ifelse(lv_types %in% mon_types, paste0("E-", lv_types), lv_types),
    period_type  = lv_types,
    total_length = paste0("=SUMIFS(", pe_rng("C"), ",", pe_rng("B"), ",B", lv_rows, ")"),
    E            = paste0("=SUMPRODUCT((", pe_rng("B"), "=B", lv_rows, ")*", pe_rng("D"), "*", pe_rng("C"), ")/C", lv_rows),
    E_se         = paste0("=SQRT(SUMPRODUCT((", pe_rng("B"), "=B", lv_rows, ")*(", pe_rng("E"), "*", pe_rng("C"), ")^2))/C", lv_rows),
    E_U          = xl_u(paste0("D", lv_rows), paste0("E", lv_rows)),
    E_lower      = paste0("=D", lv_rows, "-z_score*E", lv_rows),
    E_upper      = paste0("=D", lv_rows, "+z_score*E", lv_rows)
  )
  b_lv <- xl_block(wb, sh[["err"]], "4. Reference and monitoring emission levels (tCO2e/yr)", lv_df, start_row = lv_start,
                   num_fmt = c(total_length = "0", E = "#,##0", E_se = "#,##0", E_U = "0.0%", E_lower = "#,##0", E_upper = "#,##0"))
  add_form("Aggregation", "E_level = sum(E_period * period_length) / sum(period_length)",
           "Reference (FREL) or monitoring emission level: length-weighted mean of the annual emissions of its time periods.", sh[["err"]],
           paste0(xl_sheet(sh[["err"]]), "!D", lv_rows[1], ": ", lv_df$E[1]))
  add_form("Aggregation", "E_level_se = sqrt(sum((E_period_se * period_length)^2)) / sum(period_length)",
           "Standard error of the emission level (sum rule, constant weights).", sh[["err"]],
           paste0(xl_sheet(sh[["err"]]), "!E", lv_rows[1], ": ", lv_df$E_se[1]))
  add_form("Aggregation", "lower = X - z_score * SE ; upper = X + z_score * SE", "Confidence interval bounds (Approach 1).", sh[["err"]],
           paste0(xl_sheet(sh[["err"]]), "!G", lv_rows[1], ": ", lv_df$E_lower[1]))

  ## + 4.5. Emission reductions ####
  er_start <- b_lv$next_row
  er_rows  <- stats::setNames(er_start + 2 + seq_along(mon_types) - 1, mon_types)
  frel_row <- lv_rows[[ref_types[1]]]
  er_df <- dplyr::tibble(
    result      = paste0("ER-", mon_types),
    period_type = mon_types,
    E_REF       = paste0("=D", frel_row),
    E_REF_se    = paste0("=E", frel_row),
    E_MON       = paste0("=D", lv_rows[mon_types]),
    E_MON_se    = paste0("=E", lv_rows[mon_types]),
    ER          = paste0("=C", er_rows, "-E", er_rows),
    ER_se       = paste0("=SQRT(D", er_rows, "^2+F", er_rows, "^2)"),
    ER_U        = xl_u(paste0("G", er_rows), paste0("H", er_rows)),
    ER_lower    = paste0("=G", er_rows, "-z_score*H", er_rows),
    ER_upper    = paste0("=G", er_rows, "+z_score*H", er_rows)
  )
  b_er <- xl_block(wb, sh[["err"]], paste0("5. Emission reductions (tCO2e/yr), reference level: ", ref_types[1]), er_df, start_row = er_start,
                   num_fmt = c(E_REF = "#,##0", E_REF_se = "#,##0", E_MON = "#,##0", E_MON_se = "#,##0", ER = "#,##0",
                               ER_se = "#,##0", ER_U = "0.0%", ER_lower = "#,##0", ER_upper = "#,##0"))
  add_form("Emission reductions", "ER = E_REF - E_MON", "Emission reductions of a monitoring period type, against the first reference level.", sh[["err"]],
           paste0(xl_sheet(sh[["err"]]), "!G", er_rows[1], ": ", er_df$ER[1]))
  add_form("Emission reductions", "ER_se = sqrt(E_REF_se^2 + E_MON_se^2)", "Standard error of the emission reductions (sum rule).", sh[["err"]],
           paste0(xl_sheet(sh[["err"]]), "!H", er_rows[1], ": ", er_df$ER_se[1]))

  xl_finish_sheet(wb, sh[["err"]], n_cols = length(tr_cols), first_width = 40)


  ##
  ## 5. Calc MCS and Sims ######################################################
  ##

  ## + 5.1. Simulations ####
  has_sims <- !is.null(.mcs) && !is.null(.mcs$sim_trans) && is.data.frame(.mcs$sim_trans)
  if (has_sims) {
    sim_trans <- .mcs$sim_trans |> dplyr::arrange(.data$sim_no)
    sim_REF <- if (!is.null(.mcs$sim_REF)) .mcs$sim_REF else fct_combine_mcs_P(sim_trans, .period_type = "REF")
    sim_MON <- if (!is.null(.mcs$sim_MON)) .mcs$sim_MON else fct_combine_mcs_P(sim_trans, .period_type = "MON")
    sim_ER  <- if (!is.null(.mcs$sim_ER))  .mcs$sim_ER  else fct_combine_mcs_ER(sim_REF, sim_MON, setup$ad_annual)
    st_by   <- split(sim_trans, sim_trans$trans_id)
  }

  ## Simulations of a column of sim_trans for a land use, initial (_i) or final (_f)
  sims_lu <- function(lu, col) {
    if (!has_sims) return(NULL)
    tr_i <- area$trans_id[area$lu_initial_id == lu]
    tr_f <- area$trans_id[area$lu_final_id   == lu]
    out <- if (length(tr_i) > 0 && paste0(col, "_i") %in% names(sim_trans)) {
      st_by[[tr_i[1]]][[paste0(col, "_i")]]
    } else if (length(tr_f) > 0 && paste0(col, "_f") %in% names(sim_trans)) {
      st_by[[tr_f[1]]][[paste0(col, "_f")]]
    }
    out <- suppressWarnings(as.numeric(out))
    if (length(out) == 0 || all(is.na(out))) NULL else out
  }
  sims_of <- function(df, filter_col, filter_val, col) {
    if (!has_sims) return(NULL)
    out <- df[df[[filter_col]] == filter_val, , drop = FALSE]
    out <- out[order(out$sim_no), , drop = FALSE][[col]]
    if (length(out) == 0) NULL else as.numeric(out)
  }

  ## + 5.2. Rows of Calc MCS, per block ####
  ## Each row: id, description, value/SE references, U formula, simulations
  ad_rows <- dplyr::tibble(
    block       = "AD",
    id          = area$trans_id,
    description = paste0(area$redd_activity, ": ", area$lu_initial_id, " -> ", area$lu_final_id, " (ha)"),
    value       = paste0("=", in_ref("area", "trans_area", seq_len(nrow(area)))),
    se          = paste0("=", in_ref("area", "trans_se",   seq_len(nrow(area)))),
    sims        = purrr::map(area$trans_id, function(x) if (has_sims) as.numeric(st_by[[x]]$AD) else NULL)
  )

  c_rows <- dplyr::tibble(
    block       = "Carbon elements",
    id          = carbon$c_id,
    description = paste0(carbon$c_lu, ": ", carbon$c_element, dplyr::if_else(is.na(carbon$c_unit), "", paste0(" (", carbon$c_unit, ")"))),
    value       = paste0("=", in_ref("carbon", "c_value", carbon$xl_i)),
    se          = paste0("=", in_ref("carbon", "c_se",    carbon$xl_i)),
    sims        = purrr::map2(carbon$c_lu_id, carbon$c_element, function(lu, el) {
      if (el == "CF") {
        if (!has_sims) return(NULL)
        out <- suppressWarnings(as.numeric(st_by[[1]]$CF_i))
        if (all(is.na(out))) NULL else out
      } else {
        sims_lu(lu, el)
      }
    })
  )
  if (!is_v2 && has_cf) {
    c_rows <- dplyr::bind_rows(c_rows, dplyr::tibble(
      block = "Carbon elements", id = "CF", description = "Carbon fraction (setup)",
      value = paste0("=", cf_ref[["value"]]), se = paste0("=", cf_ref[["se"]]),
      sims  = list(if (has_sims) suppressWarnings(as.numeric(st_by[[1]]$CF_i)) else NULL)
    ))
  }

  ef_sheet <- xl_sheet(sh[["ef"]])
  cs_rows_mcs <- dplyr::tibble(
    block       = "EF",
    id          = paste0("C*", lu_tbl$c_lu_id),
    description = paste0("Carbon stock: ", lu_tbl$c_lu, " (tC/ha)"),
    value       = paste0("=", ef_sheet, "!", cs_let[["C"]],    cs_row[lu_tbl$c_lu_id]),
    se          = paste0("=", ef_sheet, "!", cs_let[["C_se"]], cs_row[lu_tbl$c_lu_id]),
    sims        = purrr::map(lu_tbl$c_lu_id, sims_lu, col = "c_stock")
  )
  ef_rows_mcs <- dplyr::tibble(
    block       = "EF",
    id          = paste0("EF*", ef_pairs$ef_id),
    description = paste0("Emission factor: ", ef_pairs$lu_initial_id, " -> ", ef_pairs$lu_final_id, " (tCO2e/ha)"),
    value       = paste0("=", ef_sheet, "!", ef_let[["EF"]],    ef_row[ef_pairs$ef_id]),
    se          = paste0("=", ef_sheet, "!", ef_let[["EF_se"]], ef_row[ef_pairs$ef_id]),
    sims        = purrr::map2(ef_pairs$lu_initial_id, ef_pairs$lu_final_id, function(li, lf) {
      if (!has_sims) return(NULL)
      tr <- area$trans_id[area$lu_initial_id == li & area$lu_final_id == lf][1]
      as.numeric(st_by[[tr]]$EF)
    })
  )

  err_sheet <- xl_sheet(sh[["err"]])
  lv_rows_mcs <- dplyr::bind_rows(
    dplyr::tibble(
      block       = "E_REF, E_MON and ER_MON",
      id          = lv_df$result,
      description = paste0(ifelse(lv_types %in% ref_types, "Reference level ", "Monitoring emissions "), lv_types, " (tCO2e/yr)"),
      value       = paste0("=", err_sheet, "!D", lv_rows),
      se          = paste0("=", err_sheet, "!E", lv_rows),
      sims        = purrr::map(lv_types, function(pt) {
        if (pt %in% ref_types) sims_of(sim_REF, "period_type", pt, "E") else sims_of(sim_MON, "period_type", pt, "E")
      })
    ),
    dplyr::tibble(
      block       = "E_REF, E_MON and ER_MON",
      id          = er_df$result,
      description = paste0("Emission reductions ", mon_types, " (tCO2e/yr)"),
      value       = paste0("=", err_sheet, "!G", er_rows),
      se          = paste0("=", err_sheet, "!H", er_rows),
      sims        = purrr::map(mon_types, function(pt) sims_of(sim_ER, "period_type", pt, "ER_sim"))
    )
  )

  mcs_blocks <- list(
    list(title = "1. AD - Activity data (input)",                             rows = ad_rows,                                   fmt = "#,##0",     input = TRUE),
    list(title = "2. Carbon elements (input)",                                rows = c_rows,                                    fmt = "#,##0.000", input = TRUE),
    list(title = "3. EF - Carbon stocks and emission factors (Approach 1)",   rows = dplyr::bind_rows(cs_rows_mcs, ef_rows_mcs), fmt = "#,##0.000", input = FALSE),
    list(title = "4. E_REF, E_MON and ER_MON (Approach 1)",                   rows = lv_rows_mcs,                               fmt = "#,##0",     input = FALSE)
  )

  ## + 5.3. Sims tab: one column per row of Calc MCS with simulations ####
  all_sims  <- do.call(c, lapply(mcs_blocks, function(b) b$rows$sims))
  all_ids   <- unlist(purrr::map(mcs_blocks, function(b) paste0(b$rows$block, ": ", b$rows$id)))
  sim_col   <- rep(NA_character_, length(all_sims))
  has_s     <- !vapply(all_sims, is.null, logical(1))
  n_iter    <- if (has_sims) max(sim_trans$sim_no) else 0
  write_sims <- has_sims && isTRUE(.sims_sheet) && any(has_s)
  if (write_sims) {
    sim_col[has_s] <- openxlsx2::int2col(seq_len(sum(has_s)))
  }

  ## + 5.4. Calc MCS tab ####
  wb$add_worksheet(sh[["mcs"]])
  xl_title(wb, sh[["mcs"]],
    "Calc MCS - Inputs and Approach 1 results vs. Monte Carlo simulations (Approach 2)",
    paste0("Value, SE and U%: inputs (blocks 1-2) or Approach 1 results (blocks 3-4). Sim columns: statistics of the simulations",
           if (write_sims) paste0(" in the ", sh[["sims"]], " tab (values, not recalculated when inputs change).") else ", as values.")
  )

  mcs_set <- dplyr::tibble(
    Parameter = c("Number of iterations", "Confidence level", "CI lower quantile", "CI upper quantile"),
    Value     = c(if (has_sims) paste0("=", n_iter) else NA_character_, "=conf_level", "=(1-conf_level)/2", "=1-(1-conf_level)/2"),
    Note      = c("Simulations per input.", "From the setup tab.", "Quantile of the simulations for the CI lower bound.", "Quantile of the simulations for the CI upper bound.")
  )
  b_mset <- xl_block(wb, sh[["mcs"]], "Settings", mcs_set, start_row = 5, num_fmt = c(Value = "0.000"))
  q_lo <- paste0("$B$", b_mset$rows[3])
  q_hi <- paste0("$B$", b_mset$rows[4])

  mcs_cols <- c("block", "id", "description", "Value", "SE", "U%", "Sim median", "Sim CI lower", "Sim CI upper", "Sim U%", "Simulations")
  next_row <- b_mset$next_row
  k0 <- 0
  mcs_ex <- NULL
  for (b in mcs_blocks) {
    rws  <- b$rows
    n    <- nrow(rws)
    r1   <- next_row + 2
    rr   <- r1 + seq_len(n) - 1
    kk   <- k0 + seq_len(n)
    sc   <- sim_col[kk]
    rng  <- ifelse(is.na(sc), NA_character_, paste0(xl_sheet(sh[["sims"]]), "!$", sc, "$2:$", sc, "$", n_iter + 1))
    df <- dplyr::tibble(
      block = rws$block, id = rws$id, description = rws$description,
      Value = rws$value, SE = rws$se, `U%` = xl_u(paste0("D", rr), paste0("E", rr))
    )
    if (write_sims) {
      df$`Sim median`   <- ifelse(is.na(rng), NA_character_, paste0("=MEDIAN(", rng, ")"))
      df$`Sim CI lower` <- ifelse(is.na(rng), NA_character_, paste0("=PERCENTILE(", rng, ",", q_lo, ")"))
      df$`Sim CI upper` <- ifelse(is.na(rng), NA_character_, paste0("=PERCENTILE(", rng, ",", q_hi, ")"))
      df$`Sim U%`       <- ifelse(is.na(rng), NA_character_, paste0("=IF(G", rr, "=0,\"\",(I", rr, "-H", rr, ")/2/ABS(G", rr, "))"))
      df$Simulations    <- ifelse(is.na(rng), NA_character_, paste0(sh[["sims"]], "!", sc, ":", sc))
    } else if (has_sims) {
      alpha <- 1 - setup$conf_level
      st <- purrr::map(rws$sims, function(s) {
        if (is.null(s)) return(c(NA_real_, NA_real_, NA_real_))
        c(stats::median(s), stats::quantile(s, alpha / 2, names = FALSE), stats::quantile(s, 1 - alpha / 2, names = FALSE))
      })
      df$`Sim median`   <- purrr::map_dbl(st, 1)
      df$`Sim CI lower` <- purrr::map_dbl(st, 2)
      df$`Sim CI upper` <- purrr::map_dbl(st, 3)
      df$`Sim U%`       <- ifelse(df$`Sim median` == 0, NA_real_, (df$`Sim CI upper` - df$`Sim CI lower`) / 2 / abs(df$`Sim median`))
      df$Simulations    <- "values from R"
    } else {
      df$`Sim median` <- df$`Sim CI lower` <- df$`Sim CI upper` <- df$`Sim U%` <- NA_real_
      df$Simulations  <- NA_character_
    }
    bb <- xl_block(
      wb, sh[["mcs"]], b$title, df, start_row = next_row,
      num_fmt = c(Value = b$fmt, SE = b$fmt, `U%` = "0.0%", `Sim median` = b$fmt, `Sim CI lower` = b$fmt,
                  `Sim CI upper` = b$fmt, `Sim U%` = "0.0%"),
      input_cols = if (b$input) c("Value", "SE") else NULL
    )
    if (k0 == 0 && write_sims && !is.na(rng[1])) mcs_ex <- list(row = rr[1], df = df)
    next_row <- bb$next_row
    k0 <- k0 + n
  }
  xl_finish_sheet(wb, sh[["mcs"]], n_cols = length(mcs_cols), first_width = 24)
  wb$set_col_widths(sheet = sh[["mcs"]], cols = 2:3, widths = c(40, 50))

  add_form("Monte Carlo", "X_sim ~ Normal(mean = value, sd = SE) or Beta(a, b), n_iter draws",
           "Simulation of each input (AD and carbon elements) with its PDF; set to 0 below 0 if trunc_pdf is TRUE. Draws are rounded to 3 decimals.", sh[["mcs"]])
  add_form("Monte Carlo", "Same calculation chain for each simulation",
           "Carbon stocks, EF, emissions, levels and ER are calculated for each simulation (fct_combine_mcs_E/P/ER), then summarised.",
           if (write_sims) sh[["sims"]] else sh[["mcs"]])
  mex <- function(col, letter) {
    if (is.null(mcs_ex)) return(NA_character_)
    paste0(xl_sheet(sh[["mcs"]]), "!", letter, mcs_ex$row, ": ", mcs_ex$df[[col]][1])
  }
  add_form("Monte Carlo", "Sim median = MEDIAN(X_sim)", "Median of the simulations, the Approach 2 estimate.", sh[["mcs"]], mex("Sim median", "G"))
  add_form("Monte Carlo", "Sim CI lower = PERCENTILE(X_sim, (1 - conf_level) / 2)", "Lower bound of the confidence interval of the simulations.", sh[["mcs"]], mex("Sim CI lower", "H"))
  add_form("Monte Carlo", "Sim CI upper = PERCENTILE(X_sim, 1 - (1 - conf_level) / 2)", "Upper bound of the confidence interval of the simulations.", sh[["mcs"]], mex("Sim CI upper", "I"))
  add_form("Monte Carlo", "Sim U% = (Sim CI upper - Sim CI lower) / 2 / |Sim median|", "Uncertainty from the simulations: half-width of the confidence interval relative to the median.", sh[["mcs"]], mex("Sim U%", "J"))

  ## + 5.5. Sims tab ####
  if (write_sims) {
    sims_df <- stats::setNames(as.data.frame(all_sims[has_s]), all_ids[has_s])
    wb$add_worksheet(sh[["sims"]])
    wb$add_data(sheet = sh[["sims"]], x = sims_df, na.strings = NULL)
    wb$add_font(sheet = sh[["sims"]], dims = paste0("A1:", openxlsx2::int2col(ncol(sims_df)), "1"), bold = TRUE)
    wb$set_col_widths(sheet = sh[["sims"]], cols = seq_len(ncol(sims_df)), widths = 18)
    wb$freeze_pane(sheet = sh[["sims"]], first_row = TRUE)
  }


  ##
  ## 6. Calc Formulas ##########################################################
  ##

  form_df <- dplyr::bind_rows(form_rows)
  wb$add_worksheet(sh[["form"]], tab_color = NULL)
  xl_title(wb, sh[["form"]],
    "Calc Formulas - All formulas used in the calculation tabs",
    "Formulas in mocaredd notation: <element> is a value, <element>_se its standard error. Examples show the Excel formula of the first cell using each formula."
  )
  xl_block(wb, sh[["form"]], "Formulas", form_df, start_row = 5)
  wb$set_col_widths(sheet = sh[["form"]], cols = 1:5, widths = c(20, 80, 80, 28, 90))
  ## Move 'Calc Formulas' before 'Sims'
  all_sh <- wb$get_sheet_names()
  if (sh[["sims"]] %in% all_sh) {
    ord <- c(setdiff(seq_along(all_sh), match(sh[["sims"]], all_sh)), match(sh[["sims"]], all_sh))
    wb$set_order(ord)
  }


  ##
  ## 7. Save ###################################################################
  ##

  ## Formulas have no cached results: ask Excel to calculate everything on load
  wb$workbook$calcPr <- '<calcPr calcId="191029" fullCalcOnLoad="1"/>'
  wb$set_active_sheet(sh[["ef"]])
  wb$save(file = .path_out, overwrite = TRUE)

  invisible(.path_out)

}


##
## Helpers #####################################################################
##

## Quoted sheet name for formulas
xl_sheet <- function(sheet) paste0("'", gsub("'", "''", sheet), "'")

## Position of a table in its tab: column letter per column name, Excel row per table row
xl_map_table <- function(wb, sheet, tbl, key) {
  raw <- openxlsx2::wb_to_df(wb, sheet = sheet, col_names = FALSE)
  hdr <- vapply(raw[1, , drop = FALSE], as.character, character(1))
  body <- raw[-1, , drop = FALSE]
  rows <- as.integer(rownames(raw))[-1]
  keep <- rowSums(!is.na(body)) > 0
  body <- body[keep, , drop = FALSE]
  rows <- rows[keep]
  col  <- stats::setNames(names(raw), hdr)
  col  <- col[!is.na(names(col)) & !duplicated(names(col))]

  if (length(rows) != nrow(tbl)) {
    stop("Tab '", sheet, "': ", length(rows), " rows in the file but ", nrow(tbl), " in the checked data.")
  }
  if (key %in% names(col)) {
    x_file <- suppressWarnings(as.numeric(body[[col[[key]]]]))
    x_data <- suppressWarnings(as.numeric(tbl[[key]]))
    same   <- (is.na(x_file) & is.na(x_data)) | abs(x_file - x_data) <= 1e-9 * pmax(1, abs(x_data))
    if (!all(same %in% TRUE)) stop("Tab '", sheet, "': rows do not match the checked data (column ", key, ").")
  }
  list(col = col, row = rows)
}

## Replace variables of an R formula by Excel cells, e.g. "AGB * (1 + RS)" -> "D11*(1+F11)"
xl_translate <- function(form, varmap) {
  out <- stringr::str_replace_all(form, "\\b[A-Za-z][A-Za-z0-9_]*\\b", function(tok) {
    vapply(tok, function(t) {
      if (t == "sqrt") return("SQRT")
      if (!t %in% names(varmap)) stop("Unknown variable '", t, "' in formula: ", form)
      varmap[[t]]
    }, character(1))
  })
  gsub(" ", "", out)
}

## U% formula, blank if the value is 0
xl_u <- function(x, se) paste0("=IF(", x, "=0,\"\",z_score*", se, "/ABS(", x, "))")

## Sheet title and subtitle
xl_title <- function(wb, sheet, title, subtitle) {
  wb$add_data(sheet = sheet, x = title, dims = "A1")
  wb$add_data(sheet = sheet, x = subtitle, dims = "A2")
  wb$add_data(sheet = sheet, x = "Blue: values linked to the input tabs. Black: calculations. Change inputs in the input tabs only.", dims = "A3")
  wb$add_font(sheet = sheet, dims = "A1", bold = TRUE, size = 14)
  wb$add_font(sheet = sheet, dims = "A2:A3", italic = TRUE, color = openxlsx2::wb_color("FF595959"))
}

## Write a block: title at start_row, header at start_row + 1, data below.
## Character cells starting with "=" are written as formulas, other cells as values.
xl_block <- function(wb, sheet, title, df, start_row, num_fmt = NULL, input_cols = NULL) {

  hdr_row <- start_row + 1
  r1      <- start_row + 2
  rows    <- r1 + seq_len(nrow(df)) - 1
  lets    <- stats::setNames(openxlsx2::int2col(seq_along(df)), names(df))

  is_f   <- function(x) is.character(x) & !is.na(x) & startsWith(x, "=")
  static <- df
  for (j in seq_along(df)) if (is.character(df[[j]])) static[[j]][is_f(df[[j]])] <- NA

  wb$add_data(sheet = sheet, x = title, dims = paste0("A", start_row))
  wb$add_data(sheet = sheet, x = static, start_row = hdr_row, start_col = 1, na.strings = NULL)

  for (j in seq_along(df)) {
    x <- df[[j]]
    if (!is.character(x)) next
    ok <- is_f(x)
    if (!any(ok)) next
    r      <- rle(ok)
    ends   <- cumsum(r$lengths)
    starts <- ends - r$lengths + 1
    for (k in which(r$values)) {
      idx <- starts[k]:ends[k]
      wb$add_formula(sheet = sheet, x = substring(x[idx], 2),
                     dims = paste0(lets[[j]], rows[idx[1]], ":", lets[[j]], rows[utils::tail(idx, 1)]))
    }
  }

  ## Styles
  last <- lets[[length(lets)]]
  wb$add_font(sheet = sheet, dims = paste0("A", start_row), bold = TRUE, size = 12)
  wb$add_fill(sheet = sheet, dims = paste0("A", hdr_row, ":", last, hdr_row), color = openxlsx2::wb_color("FFD9E1F2"))
  wb$add_font(sheet = sheet, dims = paste0("A", hdr_row, ":", last, hdr_row), bold = TRUE)
  if (nrow(df) > 0) {
    for (nm in intersect(names(num_fmt), names(df))) {
      wb$add_numfmt(sheet = sheet, dims = paste0(lets[[nm]], r1, ":", lets[[nm]], max(rows)), numfmt = num_fmt[[nm]])
    }
    for (nm in intersect(input_cols, names(df))) {
      wb$add_font(sheet = sheet, dims = paste0(lets[[nm]], r1, ":", lets[[nm]], max(rows)), color = openxlsx2::wb_color("FF0000FF"))
    }
  }

  list(rows = rows, cols = lets, next_row = max(c(rows, hdr_row)) + 3)
}

## Column widths of a calculation tab
xl_finish_sheet <- function(wb, sheet, n_cols, first_width = 24) {
  wb$set_col_widths(sheet = sheet, cols = seq_len(n_cols), widths = c(first_width, rep(14, n_cols - 1)))
}
