#' Write table headers
#'
#' @inheritParams write_stub_and_table_body
#'
#' @return Nothing. It will update the wb object with table headers.
#' @export
write_table_header <- function(gt_table, ordered_gt_data, wb, sheet_name, row_to_start) {
  current_row <- row_to_start
  total_cols <- ncol(ordered_gt_data)

  if (!is.null(gt_table[["_heading"]][["title"]])) {
    wb <- openxlsx2::wb_add_data(wb, sheet_name, gt_table[["_heading"]][["title"]], start_col = 1, start_row = current_row, col_names = FALSE)
    wb <- openxlsx2::wb_merge_cells(wb, sheet_name, dims = openxlsx2::wb_dims(cols = 1:total_cols, rows = current_row))

    # FIXME: Add style setting
    # openxlsx::addStyle(wb, sheet_name, rows = current_row, cols = 1:total_cols, style = title_style)
    current_row <- current_row + 1
  }

  ## write subtitle
  if (!is.null(gt_table[["_heading"]][["subtitle"]])) {
    wb <- openxlsx2::wb_add_data(wb, sheet_name, gt_table[["_heading"]][["subtitle"]], start_col = 1, start_row = current_row, col_names = FALSE)
    wb <- openxlsx2::wb_merge_cells(wb, sheet_name, dims = openxlsx2::wb_dims(cols = 1:total_cols, rows = current_row))
    # FIXME: Add style setting
    # openxlsx::addStyle(wb, sheet_name, rows = current_row, cols = 1:total_cols, style = title_style)
  }

  wb
}
