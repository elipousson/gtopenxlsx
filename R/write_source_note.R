#' Write table headers
#'
#' @inheritParams write_stub_and_table_body
#'
#' @return Nothing. It will update the wb object with source note.
#' @export
write_source_note <- function(gt_table, ordered_gt_data, wb, sheet_name, row_to_start) {
  current_row <- row_to_start
  total_cols <- ncol(ordered_gt_data)

  for (i in 1:length(gt_table$`_source_notes`)) {
    wb <- openxlsx2::wb_add_data(
      wb = wb,
      sheet_name = sheet_name,
      x = as.character(gt_table[["_source_notes"]][[i]]),
      start_row = current_row,
      col_names = FALSE
    )

    wb <- openxlsx2::wb_merge_cells(
      wb,
      sheet = sheet_name,
      dims = openxlsx2::wb_dims(cols = seq(total_cols), rows = current_row)
    )

    wb <- openxlsx2::wb_add_cell_style(
      wb,
      sheet = sheet_name,
      dims = openxlsx2::wb_dims(
        rows = current_row,
        cols = seq(total_cols)
        ),
      style = sourcenote_style
    )

    current_row <- current_row + 1
  }

  wb
}
