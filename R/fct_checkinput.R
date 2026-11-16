#' Check and load the mocaredd input XLSX file
#'
#' @description
#' Loads an XLSX input file (e.g. from a Shiny \code{fileInput} widget), verifies
#' that all required sheets are present, reads the four data tables, and runs six
#' conformity checks:
#' \enumerate{
#'   \item Column names
#'   \item Table sizes
#'   \item Column data types
#'   \item Category variable values
#'   \item Unique IDs
#'   \item Cross-table matching and logical consistency
#' }
#'
#' Each check emits a \code{message()}, so results can be captured outside the function with
#' \code{withCallingHandlers(message = ...)} and displayed in a console-style UI
#' element. A \code{shinyWidgets} progress bar is advanced after each of the
#' \code{n_steps} steps (1 load + 6 checks).
#'
#' @section Server-side usage:
#' \preformatted{
#' log_lines <- character()
#' result <- withCallingHandlers(
#'   tryCatch(
#'     fct_checkinput(.path = path, .pb_session = session, .pb_id = "prog_allchecks"),
#'     error = function(e) {
#'       log_lines <<- c(log_lines, paste0("\u2717 ERROR: ", conditionMessage(e)))
#'       NULL
#'     }
#'   ),
#'   message = function(m) {
#'     log_lines <<- c(log_lines, trimws(conditionMessage(m)))
#'     invokeRestart("muffleMessage")
#'   }
#' )
#' rv$check_log <- log_lines
#' if (!is.null(result)) {
#'   rv$inputs$setup  <- result$data$setup
#'   rv$inputs$time   <- result$data$time
#'   rv$inputs$area   <- result$data$area
#'   rv$inputs$carbon <- result$data$carbon
#'   rv$checks$all_ok <- result$all_ok
#' }
#' }
#'
#' @param .path       Character. Path to the XLSX file (e.g. \code{input$load_xlsx$datapath}).
#' @param .pb_session Shiny session object used to update the progress bar. Pass \code{NULL}
#'                    (default) to skip progress bar updates (e.g. when calling outside Shiny).
#' @param .pb_id      Character. The \code{id} of the \code{shinyWidgets::progressBar()} to
#'                    update. Ignored when \code{.pb_session} is \code{NULL}.
#' @param .pb_max     Value between 0 and 100, default to 100. What should progress bar max be
#'                    after all checks are passed.
#' @param .minislow   NA or numeric. If numeric, add that value in seconds to system sleep to
#'                    make checks more visible in the console.
#'
#' @return A list with three elements:
#' \describe{
#'   \item{\code{all_ok}}{Logical. \code{TRUE} only if every check passed.}
#'   \item{\code{template_version}}{Integer, 1 or 2, the detected template version.
#'         Downstream calculation functions take this whole list as \code{.checked_data}
#'         and read \code{.checked_data$template_version} to know where to find the carbon
#'         unit and carbon fraction.}
#'   \item{\code{data}}{Named list with elements \code{setup}, \code{time}, \code{area},
#'         \code{carbon} (the four loaded tables as tibbles).}
#' }
#' The function throws an error (via \code{stop()}) only for unrecoverable failures
#' (unreadable file, missing sheets). All check outcomes are reported via \code{message()}.
#'
#' @importFrom rlang .data
#'
#' @examples
#' path <- system.file("extdata/example1-4pools.xlsx", package = "mocaredd")
#'
#' log_lines <- character()
#' result <- withCallingHandlers(
#'   fct_checkinput(.path = path),
#'   message = function(m) {
#'     log_lines <<- c(log_lines, trimws(conditionMessage(m)))
#'     invokeRestart("muffleMessage")
#'   }
#' )
#' cat(log_lines, sep = "\n")
#'
#' @export
fct_checkinput <- function(.path, .pb_session = NULL, .pb_id = NULL, .pb_max = 100, .minislow = NA) {


  ## !!! FOR TESTING ONLY
  # .path = system.file("extdata/example1-4pools.xlsx", package = "mocaredd")
  # .path = system.file("extdata/mocaredd-templatev2-4pools.xlsx", package = "mocaredd")
  # .path = system.file("extdata/mocaredd-templatev2-simple.xlsx", package = "mocaredd")
  # .path = system.file("extdata/test-simple-noCF.xlsx", package = "mocaredd")
  # .pb_id = NULL ; .pb_session = NULL ; .pb_max = 80 ; .minislow = NA
  ## !!!

  ##
  ## Constants #################################################################
  ##

  checklist <- list(
    tabs   = c("user_inputs", "time_periods", "AD_lu_transitions", "c_stocks"),
    tabsv2 = c("setup", "time", "area", "carbon"),
    cats   = list(
      c_unit = c("DM", "C", "ratio"),
      c_pool = c("AGB", "BGB", "DW", "LI", "SOC", "ALL"),
      r_acti = c("DF", "DG", "EN", "EN_AF", "EN_RE"),
      p_type = c("REF", "REF[0-9]", "MON", "MON[0-9]"),
      pdf    = c("normal", "beta"),
      c_var  = c("CF", "RS", "DG_ratio")
    ),
    cols = list(
      setup  = c("n_iter", "c_unit", "dg_ext", "dg_pool", "ad_annual",
                 "conf_level", "digits", "ran_seed", "trunc_pdf"),
      setupv2  = c("n_iter", "dg_ext", "dg_pool", "ad_annual",
                 "conf_level", "digits", "ran_seed", "trunc_pdf"),
      time   = c("period_no", "year_start", "year_end", "period_type"),
      area   = c("trans_no", "trans_id", "trans_period", "lu_initial_id", "lu_final_id",
                 "trans_area", "trans_se", "trans_pdf", "trans_pdf_a",
                 "trans_pdf_b", "trans_pdf_c", "redd_activity"),
      carbon = c("c_no", "c_id", "c_period", "c_element", "c_lu_id", "c_value",
                 "c_se", "c_pdf", "c_pdf_a", "c_pdf_b", "c_pdf_c"),
      areav2 = c("trans_no", "trans_period", "lu_initial", "lu_final",
                 "trans_area", "trans_se", "trans_pdf", "trans_pdf_a",
                 "trans_pdf_b", "trans_pdf_c", "redd_activity"),
      carbonv2 = c("c_no", "c_period", "c_element", "c_unit", "c_lu", "c_value",
                   "c_se", "c_pdf", "c_pdf_a", "c_pdf_b", "c_pdf_c")
    )
  )


  ##
  ## Helpers ###################################################################
  ##

  ## Advance the progress bar by one step
  show_progress <- function(.title, .value) {
    if (!is.null(.pb_session) && !is.null(.pb_id)) {
      shinyWidgets::updateProgressBar(
        session = .pb_session,
        id      = .pb_id,
        value   = round(.value),
        title   = .title,
        status  = "primary"
      )
    } else {
      message(paste0(.title, " - progress: ", round(.value), "%."))
    }
  }

  ## Emit a pass or fail message; mark all_ok FALSE on failure
   show_msg <- function(check, pass_label, fail_label, problems = NULL, stop_on_fail = TRUE) {
    if (check) {
      message("\u2713 ", pass_label)
    } else {
      detail <- if (length(problems) > 0L) {
        paste0(" [", paste(problems, collapse = "; "), "]")
      } else {
        ""
      }
      message("\u2717 ", fail_label, detail)
      if (stop_on_fail) {
        if (!is.null(.pb_session) && !is.null(.pb_id)) {
          shinyWidgets::updateProgressBar(
            session = .pb_session,
            id      = .pb_id,
            value   = 100,
            title   = "Checks halted \u2014 fix the issue above and re-run.",
            status  = "danger"
            )
          }
        message("-- Subsequent checks skipped: fix the issue above first.")
        return(list(all_ok = FALSE, data = NULL))
        #stop(fail_label, detail, call. = FALSE)
        }
    }
  }

  ## + Problem reporting ------
  ## Each table check returns NULL when it passes, or a short text locating the
  ## first problem, e.g. "first problem: col: c_lu, row: 5, value: empty".
  ## Rows are Excel row numbers: row 1 is the header, data row i is row i + 1.

  table_labels <- c(
    setup  = "setup/user_inputs",
    time   = "time/time_periods",
    area   = "area/AD_lu_transitions",
    carbon = "carbon/c_stocks"
  )

  ## Return "<table label> (<first problem>)" for each failed table. Problems
  ## are passed as named arguments (setup = , time = , area = , carbon = ) so
  ## each one is matched to its table label by name, never by position.
  failed_tables <- function(...) {
    probs <- list(...)
    stopifnot(
      !is.null(names(probs)),
      all(names(probs) %in% names(table_labels))
    )
    probs <- Filter(Negate(is.null), probs)
    vapply(
      names(probs),
      function(tbl) paste0(table_labels[[tbl]], " (", probs[[tbl]], ")"),
      character(1),
      USE.NAMES = FALSE
    )
  }

  ## Prefix a problem text with its table label, for cross-table checks
  in_table <- function(.tbl, .txt) {
    if (is.null(.txt)) NULL else paste0(table_labels[[.tbl]], ", ", .txt)
  }

  ## Return the first non-NULL argument
  first_non_null <- function(...) {
    for (p in list(...)) if (!is.null(p)) return(p)
    NULL
  }

  ## Format a cell value for messages
  fmt_value <- function(x) {
    if (length(x) == 0L || is.na(x)) "empty" else paste0("'", x, "'")
  }

  ## Locate the first problem cell. '.bad' is a named list with one logical
  ## vector per column (TRUE = problem in that row). The lowest row wins, ties
  ## go to the first column listed. '.notes' optionally adds a text per column.
  first_cell <- function(.df, .bad, .notes = NULL) {
    rows <- vapply(.bad, function(b) {
      i <- which(b)[1]
      if (is.na(i)) NA_integer_ else as.integer(i)
    }, integer(1))
    if (length(rows) == 0L || all(is.na(rows))) return(NULL)
    col  <- names(rows)[which.min(rows)]
    i    <- rows[[col]]
    val  <- if (col %in% names(.df)) paste0(", value: ", fmt_value(.df[[col]][i])) else ""
    note <- if (!is.null(.notes) && col %in% names(.notes)) paste0(", ", .notes[[col]]) else ""
    paste0("first problem: col: ", col, ", row: ", i + 1L, val, note)
  }

  ## Make ID from labels
  label2id <- function(label) {
    label |>
      stringr::str_to_lower() |>
      stringr::str_replace_all(" ", "_") |>
      stringr::str_replace("non.forest", "NF") |>
      stringr::str_remove("_forest") |>
      stringr::str_remove("forest")
  }

  ## system sleep to make checkks more visible in console
  minislow <- function(.num) {

    if (is.null(.num) || length(.num) != 1) return(invisible(NULL))

    num <- suppressWarnings(as.numeric(.num))

    if (!is.na(num) && is.finite(num) && num >= 0) {
      Sys.sleep(num)
    }

    invisible(NULL)

  }

  ##
  ## State #####################################################################
  ##

  ## 1 load step + 6 checks - update this if steps are added or removed
  n_steps <- 7L
  pb_factor <- .pb_max / n_steps

  ## Initiate
  step   <- 0L

  ## If check fails return out_fail
  out_fail <- list(all_ok = FALSE, data = NULL)

  ##
  ## Step 1: Load data #########################################################
  ##

  show_progress("Loading data...", step * pb_factor)

  sheets_found <- tryCatch(
    readxl::excel_sheets(.path),
    error = function(e) NULL
  )
  if (is.null(sheets_found)) {
    message("-- Cannot read the file. Is it a valid .xlsx file?")
    return(out_fail)
  }

  missing_tabsv1 <- setdiff(checklist$tabs, sheets_found)
  missing_tabsv2 <- setdiff(checklist$tabsv2, sheets_found)

  if (length(missing_tabsv1) == 0L) {
    check_version <- 1

    setup  <- readxl::read_xlsx(.path, sheet = "user_inputs",       na = "NA")
    ## Older v1 templates name 'digits' 'round_digits'
    if (!"digits" %in% names(setup) && "round_digits" %in% names(setup)) {
      setup <- dplyr::rename(setup, digits = "round_digits")
    }
    time   <- readxl::read_xlsx(.path, sheet = "time_periods",      na = "NA")
    area   <- readxl::read_xlsx(.path, sheet = "AD_lu_transitions", na = "NA")
    carbon <- readxl::read_xlsx(.path, sheet = "c_stocks",          na = "NA")

    show_msg(check = T, pass_label = paste0("Tables loaded successfully from template v", check_version))

  } else if (length(missing_tabsv2) == 0L) {
    check_version <- 2

    setup  <- readxl::read_xlsx(.path, sheet = "setup" , na = "NA")
    time   <- readxl::read_xlsx(.path, sheet = "time"  , na = "NA")
    area   <- readxl::read_xlsx(.path, sheet = "area"  , na = "NA")
    carbon <- readxl::read_xlsx(.path, sheet = "carbon", na = "NA")

    show_msg(check = T, pass_label = paste0("Tables loaded successfully from template v", check_version))

  } else {
    show_msg(check = F, fail_label = paste0("Missing required sheet(s): ", paste(missing_tabsv2, collapse = ", ")))
    return(out_fail)
  }

  minislow(.minislow)

  step <- step + 1L

  ##
  ## Step 2: Column names #####################################################
  ##

  show_progress("Checking column names...", step * pb_factor)

  ## First required column not found in the header row
  missing_col <- function(.df, .cols) {
    m <- setdiff(.cols, names(.df))
    if (length(m) == 0L) return(NULL)
    paste0(
      "first problem: col: ", m[1], ", row: 1, column not found",
      if (length(m) > 1L) paste0(" (", length(m) - 1L, " more missing)") else ""
    )
  }

  cols_req <- if (check_version == 1) {
    checklist$cols[c("setup", "time", "area", "carbon")]
  } else {
    checklist$cols[c("setupv2", "time", "areav2", "carbonv2")]
  }

  prob_cols <- failed_tables(
    setup  = missing_col(setup,  cols_req[[1]]),
    time   = missing_col(time,   cols_req[[2]]),
    area   = missing_col(area,   cols_req[[3]]),
    carbon = missing_col(carbon, cols_req[[4]])
  )

  check_cols_all <- length(prob_cols) == 0L

  show_msg(
    check      = check_cols_all,
    pass_label = "Column names: all required columns present",
    fail_label = "Column names: missing columns in",
    problems   = prob_cols
  )

  if (!check_cols_all) return(out_fail)

  minislow(.minislow)

  step <- step + 1L

  ##
  ## Step 3: Table sizes ######################################################
  ##

  show_progress("Checking table dimensions...", step * pb_factor)

  ## Row count outside [.min, .max]
  size_problem <- function(.df, .min, .max = Inf) {
    n <- nrow(.df)
    if (n < .min) {
      return(paste0("first problem: ", n, " data row(s) found, at least ", .min, " expected"))
    }
    if (n > .max) {
      return(paste0("first problem: row: ", .max + 2L, ", only ", .max, " data row(s) expected"))
    }
    NULL
  }

  ## Period type present at least once (at least one REF + one MON period)
  ptype_problem <- function(.ptype) {
    if (any(stringr::str_detect(time$period_type, .ptype), na.rm = TRUE)) return(NULL)
    paste0("first problem: col: period_type, no ", .ptype, " period found")
  }

  prob_size <- failed_tables(
    setup  = size_problem(setup, 1L, 1L),  # exactly one settings row
    time   = first_non_null(size_problem(time, 2L), ptype_problem("REF"), ptype_problem("MON")),
    area   = size_problem(area, 2L),       # at least two land use transitions
    carbon = size_problem(carbon, 2L)      # at least one initial + one final C stock
  )

  check_size_all <- length(prob_size) == 0L

  show_msg(
    check      = check_size_all,
    pass_label = "Table sizes: all tables have sufficient rows",
    fail_label = "Table sizes: unexpected row count in",
    problems   = prob_size
  )

  if (!check_size_all) return(out_fail)

  minislow(.minislow)

  step <- step + 1L

  ##
  ## Step 4: Data types #######################################################
  ##

  show_progress("Checking column data types...", step * pb_factor)

  ## Expected types: num, chr, lgl, or num_lgl / chr_lgl (also accept an
  ## empty column, which readxl reads as logical)
  type_notes <- c(
    num     = "expected a number",
    chr     = "expected text",
    lgl     = "expected TRUE or FALSE",
    num_lgl = "expected a number or empty",
    chr_lgl = "expected text or empty"
  )

  type_ok <- function(x, type) {
    switch(
      type,
      num     = is.numeric(x),
      chr     = is.character(x),
      lgl     = is.logical(x),
      num_lgl = is.numeric(x) || is.logical(x),
      chr_lgl = is.character(x) || is.logical(x)
    )
  }

  ## Rows that make a column fail its type: text in a numeric column, non
  ## TRUE/FALSE values in a logical column, else empty cells (a column with no
  ## valid value), else the first row.
  type_bad_rows <- function(x, type) {
    bad <- switch(
      sub("_lgl$", "", type),
      num = !is.na(x) & is.na(suppressWarnings(as.numeric(x))),
      lgl = !is.na(x) & !toupper(as.character(x)) %in% c("TRUE", "FALSE"),
      chr = rep(FALSE, length(x))
    )
    if (!any(bad)) bad <- is.na(x)
    if (!any(bad)) bad <- seq_along(x) == 1L
    bad
  }

  type_problem <- function(.df, .spec) {
    is_ok <- vapply(names(.spec), function(col) type_ok(.df[[col]], .spec[[col]]), logical(1))
    .spec <- .spec[!is_ok]
    if (length(.spec) == 0L) return(NULL)
    bad   <- purrr::imap(.spec, function(type, col) type_bad_rows(.df[[col]], type))
    notes <- stats::setNames(unname(type_notes[.spec]), names(.spec))
    first_non_null(
      first_cell(.df, bad, .notes = notes),
      paste0("first problem: col: ", names(.spec)[1], ", ", notes[[1]])
    )
  }

  v1 <- check_version == 1

  spec_setup <- c(
    n_iter = "num", if (v1) c(c_unit = "chr"), dg_pool = "chr_lgl", ad_annual = "lgl",
    conf_level = "num", digits = "num", ran_seed = "num_lgl", trunc_pdf = "lgl"
  )

  spec_time <- c(period_no = "chr", year_start = "num", year_end = "num", period_type = "chr")

  spec_area <- c(
    trans_no = "num", if (v1) c(trans_id = "chr"), trans_period = "chr",
    if (v1) c(lu_initial_id = "chr", lu_final_id = "chr") else c(lu_initial = "chr", lu_final = "chr"),
    trans_area = "num", trans_se = "num", trans_pdf = "chr",
    trans_pdf_a = "num_lgl", trans_pdf_b = "num_lgl", trans_pdf_c = "num_lgl",
    redd_activity = "chr"
  )

  spec_carbon <- c(
    c_no = "num", if (v1) c(c_id = "chr"), c_period = "chr", c_element = "chr",
    if (v1) c(c_lu_id = "chr") else c(c_unit = "chr", c_lu = "chr"),
    c_value = "num", c_se = "num", c_pdf = "chr",
    c_pdf_a = "num_lgl", c_pdf_b = "num_lgl", c_pdf_c = "num_lgl"
  )

  prob_dt <- failed_tables(
    setup  = type_problem(setup,  spec_setup),
    time   = type_problem(time,   spec_time),
    area   = type_problem(area,   spec_area),
    carbon = type_problem(carbon, spec_carbon)
  )

  check_dt_all <- length(prob_dt) == 0L

  show_msg(
    check      = check_dt_all,
    pass_label = "Data types: all columns have correct data types",
    fail_label = "Data types: unexpected column types in",
    problems   = prob_dt
  )

  if (!check_dt_all) return(out_fail)

  minislow(.minislow)

  step <- step + 1L


  ##
  ## Step 5: Category variables ###############################################
  ##

  show_progress("Checking category values...", step * pb_factor)

  cats    <- checklist$cats
  allowed <- function(x) paste0("allowed: ", paste(x, collapse = ", "))

  ## dg_pool can be a comma-separated string e.g. "AGB, BGB"
  bad_dg_pool <- !purrr::map_lgl(
    stringr::str_split(setup$dg_pool, pattern = ","),
    function(x) all(stringr::str_trim(x) %in% cats$c_pool)
  )

  ptype_pattern <- paste0(cats$p_type, collapse = "|")

  ## v2: c_unit must be a known unit, and DM or C for carbon pools
  bad_c_unit <- if (!v1) {
    (!is.na(carbon$c_unit) & !carbon$c_unit %in% cats$c_unit) |
      (carbon$c_element %in% cats$c_pool & !carbon$c_unit %in% c("DM", "C"))
  }

  prob_cats <- failed_tables(
    setup = first_cell(
      setup,
      c(if (v1) list(c_unit = !setup$c_unit %in% cats$c_unit), list(dg_pool = bad_dg_pool)),
      .notes = c(c_unit = allowed(cats$c_unit), dg_pool = allowed(cats$c_pool))
    ),
    time = first_cell(
      time,
      list(period_type = !is.na(time$period_type) & !stringr::str_detect(time$period_type, ptype_pattern)),
      .notes = c(period_type = "allowed: REF, MON, or with a number e.g. REF1, MON2")
    ),
    area = first_cell(
      area,
      list(
        trans_pdf     = !stringr::str_to_lower(area$trans_pdf) %in% cats$pdf,
        redd_activity = !area$redd_activity %in% cats$r_acti
      ),
      .notes = c(trans_pdf = allowed(cats$pdf), redd_activity = allowed(cats$r_acti))
    ),
    carbon = first_cell(
      carbon,
      c(
        list(c_element = !carbon$c_element %in% c(cats$c_pool, cats$c_var)),
        if (!v1) list(c_unit = bad_c_unit),
        list(c_pdf = !stringr::str_to_lower(carbon$c_pdf) %in% cats$pdf)
      ),
      .notes = c(
        c_element = allowed(c(cats$c_pool, cats$c_var)),
        c_unit    = paste0(allowed(cats$c_unit), " (DM or C for carbon pools)"),
        c_pdf     = allowed(cats$pdf)
      )
    )
  )

  check_cats_all <- length(prob_cats) == 0L

  show_msg(
    check      = check_cats_all,
    pass_label = "Category values: all categories are valid",
    fail_label = "Category values: invalid categories in",
    problems   = prob_cats
  )

  if (!check_cats_all) return(out_fail)

  minislow(.minislow)

  step <- step + 1L


  ##
  ## Step 6: Unique IDs #######################################################
  ##

  show_progress("Checking unique IDs...", step * pb_factor)

  ## Create unique IDs for 'area' and 'carbon'
  if (check_version == 2) {
    area <- area |>
      dplyr::mutate(
        lu_initial_id = label2id(.data$lu_initial),
        lu_final_id = label2id(.data$lu_final),
        #trans_id = paste(trans_period, lu_initial_id, "2", lu_final_id, sep = "*")
        trans_id = paste(.data$trans_period, .data$lu_initial_id, .data$lu_final_id, sep = "*")
      )

    unique_c_period <- identical(unique(carbon$c_period), "ALL")

    carbon <- carbon |>
      dplyr::mutate(
        c_lu_id = label2id(.data$c_lu),
        c_plu_id = if (unique_c_period) .data$c_lu_id else paste(.data$c_period, .data$c_lu_id, sep = "*"),
        c_id = if (unique_c_period) paste(.data$c_lu_id, .data$c_element, sep = "*") else paste(.data$c_period, .data$c_lu_id, .data$c_element, sep = "*")
      )
  } else {
    ## v1: IDs are in the template, only the period x LU ID is added (RS / AGB match)
    carbon <- carbon |>
      dplyr::mutate(c_plu_id = paste(.data$c_period, .data$c_lu_id, sep = "*"))
  }


  ## Get dg_ext to IDs
  setup$dg_ext <- label2id(setup$dg_ext) |> stringr::str_remove_all("\"")

  ## First missing or duplicated ID. '.parts' are the template columns the ID
  ## is made of (one column in v1, several in v2 where IDs are built above).
  id_problem <- function(.df, .id, .parts) {
    miss <- Reduce(`|`, lapply(.parts, function(p) is.na(.df[[p]])))
    dup  <- duplicated(.id)
    i    <- which(miss | dup)[1]
    if (is.na(i)) return(NULL)
    if (miss[i]) {
      col <- .parts[is.na(unlist(.df[i, .parts]))][1]
      paste0("first problem: col: ", col, ", row: ", i + 1L, ", value: empty")
    } else {
      paste0(
        "first problem: col: ", paste(.parts, collapse = " + "), ", row: ", i + 1L,
        ", duplicate of row ", match(.id[i], .id) + 1L
      )
    }
  }

  id_parts_area   <- if (v1) "trans_id" else c("trans_period", "lu_initial", "lu_final")
  id_parts_carbon <- if (v1) "c_id" else if (unique_c_period) c("c_lu", "c_element") else c("c_period", "c_lu", "c_element")

  prob_ids <- failed_tables(
    time   = id_problem(time,   time$period_no, "period_no"),
    area   = id_problem(area,   area$trans_id,  id_parts_area),
    carbon = id_problem(carbon, carbon$c_id,    id_parts_carbon)
  )

  check_ids_all <- length(prob_ids) == 0L

  show_msg(
    check      = check_ids_all,
    pass_label = "Unique IDs: no duplicates or missing IDs found",
    fail_label = "Unique IDs: duplicate or missing IDs found in",
    problems   = prob_ids
  )

  if (!check_ids_all) return(out_fail)

  minislow(.minislow)

  step <- step + 1L


  ##
  ## Step 7: Cross- and intra-table matching and logical consistency ##########
  ##

  show_progress("Checking cross-table and intra_table consistency...", step * pb_factor)

  ## + Match periods ------
  ## Periods match between AD and time tables
  check_match_period_area <- setequal(area$trans_period, time$period_no)

  ## Periods match between CS and time tables (skip when all rows use "ALL")
  carbon_with_period <- carbon |> dplyr::filter(.data$c_period != "ALL")
  check_match_period_carbon <- if (nrow(carbon_with_period) > 0L) {
    setequal(carbon$c_period, time$period_no)
  } else {
    TRUE
  }

  ## + Match land uses ------
  ## All land uses in 'area' must be present in 'carbon'
  # lu_in_area <- sort(unique(c(area$lu_initial, area$lu_final)))
  # lu_in_carbon <- sort(unique(carbon$c_lu))
  # check_match_lu <- all(lu_in_area %in% lu_in_carbon)

  lu_in_area <- sort(unique(c(area$lu_initial_id, area$lu_final_id)))
  lu_in_carbon <- sort(unique(carbon$c_lu_id))
  check_match_lu_id <- all(lu_in_area %in% lu_in_carbon)

  ## + Carbon fraction x dry matter ------
  ## Carbon fraction must be provided when unit is dry matter
  check_match_dm <- if (check_version == 1) {
    if (setup$c_unit == "DM") {
      all(setup$c_fraction > 0, setup$c_fraction_se >= 0)
    } else if (setup$c_unit == "C") {
      TRUE
    } else {
      FALSE
    }
  } else if (check_version == 2) {
    if ("DM" %in% unique(carbon$c_unit)) {
      all("CF" %in% carbon$c_element, carbon$c_value[carbon$c_element == "CF"] > 0, carbon$c_se[carbon$c_element == "CF"] >= 0)
    } else if ("C" %in% unique(carbon$c_unit)) {
      TRUE
    } else {
      FALSE
    }
  } else {
    FALSE
  }

  ## + Match RS with AGB ------
  ## if RS is use, there should be AGB for the same LU
  check_match_rsagb1 <- if ("RS" %in% unique(carbon$c_element)) {
    lu_rs <- carbon |> dplyr::filter(.data$c_element == "RS") |> dplyr::pull("c_plu_id") |> sort()
    lu_agb <- carbon |> dplyr::filter(.data$c_element == "AGB") |> dplyr::pull("c_plu_id") |> sort()
    length(setdiff(lu_rs, lu_agb)) == 0
  } else {
    TRUE
  }

  # ## C unit for AGB and RS should be same -- NOT USED now RS dimensionless ratio
  # check_match_rsagb2 <- if ("RS" %in% unique(carbon$c_element) & check_match_rsagb1) {
  #   lu_rsagb <- carbon |>
  #     dplyr::filter(c_element %in% c("AGB", "RS")) |>
  #     tidyr::pivot_wider(id_cols = c(c_plu_id), values_from = c_unit, names_from = c_element)
  #   all(lu_rsagb$AGB == lu_rsagb$RS)
  # } else if (!check_match_rsagb1) {
  #   FALSE
  # } else {
  #   TRUE
  # }

  # check_match_rsagb <- all(check_match_rsagb1, check_match_rsagb2)
  check_match_rsagb <- all(check_match_rsagb1)

  ## + Match degraded and intact LU ------
  ## If DG_ratio is used, dg_ext must strip degraded LU names back to intact LUs
  dg_lu_bad <- character(0)
  if (!is.na(setup$dg_ext) && "DG_ratio" %in% unique(carbon$c_element)) {
    dg_lu <- carbon |>
      dplyr::filter(.data$c_element == "DG_ratio") |>
      dplyr::pull("c_lu_id") |>
      unique()
    dg_lu_intact <- dg_lu |>
      stringr::str_remove(pattern = setup$dg_ext) |>
      stringr::str_replace("[^A-Za-z]$", "") ## Remove trailing special character if any

    ## The intact LU must exist and not be degraded itself. If dg_ext is not found in
    ## a degraded LU ID, the 'intact' LU is the degraded LU and the C stock calculation
    ## would never end.
    dg_ok     <- dg_lu_intact %in% unique(carbon$c_lu_id) & !dg_lu_intact %in% dg_lu
    dg_lu_bad <- dg_lu[!dg_ok]
    check_match_dg <- all(dg_ok)
  } else if (is.na(setup$dg_ext) && "DG_ratio" %in% unique(carbon$c_element)) {
    check_match_dg <- FALSE   # DG_ratio used but no extension provided
  } else {
    check_match_dg <- TRUE
  }




  ## + Locate first problems ------
  ## Only computed for failed checks
  lu_cols <- if (v1) c("lu_initial_id", "lu_final_id") else c("lu_initial", "lu_final")
  c_lu_col <- if (v1) "c_lu_id" else "c_lu"

  loc_period_area <- if (!check_match_period_area) first_non_null(
    in_table("area", first_cell(area, list(trans_period = !area$trans_period %in% time$period_no))),
    in_table("time", first_cell(
      time, list(period_no = !time$period_no %in% area$trans_period),
      .notes = c(period_no = "no transitions for this period"))
    )
  )

  loc_period_carbon <- if (!check_match_period_carbon) first_non_null(
    in_table("carbon", first_cell(carbon, list(c_period = !carbon$c_period %in% time$period_no))),
    in_table("time", first_cell(
      time, list(period_no = !time$period_no %in% carbon$c_period),
      .notes = c(period_no = "no carbon stocks for this period"))
    )
  )

  loc_lu_id <- if (!check_match_lu_id) in_table("area", first_cell(
    area,
    stats::setNames(list(!area$lu_initial_id %in% lu_in_carbon, !area$lu_final_id %in% lu_in_carbon), lu_cols)
  ))

  loc_dm <- if (!check_match_dm) {
    if (v1) {
      in_table("setup", first_cell(setup, list(
        c_unit        = !setup$c_unit %in% c("DM", "C"),
        c_fraction    = !isTRUE(setup$c_fraction > 0),
        c_fraction_se = !isTRUE(setup$c_fraction_se >= 0)
      )))
    } else if (!any(carbon$c_unit %in% c("DM", "C"))) {
      in_table("carbon", "first problem: col: c_unit, no row with unit DM or C")
    } else if (!"CF" %in% carbon$c_element) {
      in_table("carbon", "first problem: col: c_element, no CF row found")
    } else {
      is_cf <- carbon$c_element %in% "CF"
      in_table("carbon", first_cell(carbon, list(
        c_value = is_cf & !dplyr::coalesce(carbon$c_value > 0, FALSE),
        c_se    = is_cf & !dplyr::coalesce(carbon$c_se >= 0, FALSE)
      )))
    }
  }

  loc_rsagb <- if (!check_match_rsagb) in_table("carbon", first_cell(
    carbon,
    stats::setNames(list(
      carbon$c_element %in% "RS" & !carbon$c_plu_id %in% carbon$c_plu_id[carbon$c_element %in% "AGB"]
    ), c_lu_col),
    .notes = stats::setNames("no AGB for this land use", c_lu_col)
  ))

  loc_dg <- if (!check_match_dg) {
    if (is.na(setup$dg_ext)) {
      in_table("setup", "first problem: col: dg_ext, row: 2, value: empty, needed when DG_ratio is used")
    } else {
      in_table("carbon", first_cell(
        carbon,
        stats::setNames(list(carbon$c_element %in% "DG_ratio" & carbon$c_lu_id %in% dg_lu_bad), c_lu_col)
      ))
    }
  }

  ## + Recap ------
  with_loc <- function(.label, .loc) if (is.null(.loc)) .label else paste0(.label, " (", .loc, ")")

  match_flags  <- c(check_match_period_area, check_match_period_carbon, check_match_lu_id,
                    check_match_dm, check_match_rsagb, check_match_dg)
  match_labels <- c(
    with_loc("trans_period does not match period_no", loc_period_area),
    with_loc("c_period does not match period_no", loc_period_carbon),
    with_loc("land use IDs present in AD/area but missing from c_stocks/carbon", loc_lu_id),
    with_loc("c_unit is 'DM' but CF is missing in usr(v1)/carbon(v2)", loc_dm),
    with_loc("RS defined for land uses with no AGB values or mismatch C unit", loc_rsagb),
    with_loc(
      paste0(
        "DG_ratio used but dg_ext does not resolve to existing LU IDs",
        if (length(dg_lu_bad) > 0) paste0(" (dg_ext: '", setup$dg_ext, "', LUs: ", paste(dg_lu_bad, collapse = ", "), ")")
      ),
      loc_dg
    )
  )

  check_match_all <- all(match_flags)

  show_msg(
    check      = check_match_all ,
    pass_label = "Cross-table consistency: all references match",
    fail_label = "Cross-table consistency: mismatches found",
    problems   = match_labels[!match_flags]
  )

  if (!check_match_all) return(out_fail)

  minislow(.minislow)


  ##
  ## Final progress bar update #################################################
  ##

  all_ok <- all(
    check_cols_all, check_size_all, check_dt_all, check_cats_all, check_ids_all,
    check_match_all
    )

  if (!is.null(.pb_session) && !is.null(.pb_id)) {
    shinyWidgets::updateProgressBar(
      session = .pb_session,
      id      = .pb_id,
      value   = 100,
      title   = if (all_ok) "All checks passed!" else "Some checks failed - see details.",
      status  = if (all_ok) "success" else "danger"
    )
  }

  if (all_ok) {
    message("-- All checks passed.")
  } else {
    message("-- One or more checks failed. Fix the issues above before running simulations.")
  }

  ## ADD calculated elements to data
  time$nb_years  <- time$year_end - time$year_start + 1
  #setup$ci_alpha <- 1 - setup$conf_level
  setup$z_score        <- stats::qnorm(1 - (1 - setup$conf_level) / 2)
  setup$conf_level_txt <- paste0(setup$conf_level * 100, "%")


  ##
  ## Return ####################################################################
  ##

  list(
    all_ok           = all_ok,
    template_version = check_version,
    data             = list(setup = setup, time = time, area = area, carbon = carbon)
  )

}
