#' Write stub and table body
#'
#' @inheritParams gt_to_xlsx
#' @param ordered_gt_data ordered data from `create_ordered_data`
#' @param row_to_start Number of the row where to start writting the stub and body.
#'
#' @return It will update the wb object with stub and body table. In addition, it will return a list
#' with the first row and last row with information for each group.
#' @export
#'
#' @examples
#' wb <- openxlsx2::wb_workbook()
#' wb <- openxlsx2::wb_add_worksheet(wb, "gtcars_example")
#'
#' example_0 <- gt::gtcars |>
#'   gt::gt()
#'
#' gtcars_example <- example_0 |>
#'   create_ordered_data()
#'
#' example_0 |>
#'   write_stub_and_table_body(ordered_gt_data = gtcars_example, wb, "gtcars_example", 1)
write_stub_and_table_body <- function(gt_table, ordered_gt_data, wb, sheet_name, row_to_start) {
  label_groups_row <- list()
  cell_data_row <- list()
  if (rlang::is_empty(gt_table[["_row_groups"]])) {
    wb <- openxlsx2::wb_add_data(wb, sheet_name, ordered_gt_data,
      start_row = row_to_start, col_names = FALSE#,
      # FIXME: Implement borders
      # borders = "surrounding",
      # borderColour = "#D3D3D3"
    )

    cell_data_row[["all"]][["start"]] <- row_to_start
    cell_data_row[["all"]][["end"]] <- row_to_start + nrow(ordered_gt_data) - 1
  } else {
    # stub
    ## row group label
    row_start_stub <- row_to_start

    for (i in unique(gt_table[["_row_groups"]])) {
      wb <- openxlsx2::wb_add_data(wb, sheet_name, i, start_row = row_start_stub)
      wb <- openxlsx2::wb_merge_cells(wb, sheet_name, dims = openxlsx2::wb_dims(rows = row_start_stub, cols = seq(ncol(ordered_gt_data))))

      # FIXME: Replace addStyle
      # openxlsx::addStyle(
      #   wb = wb, sheet = sheet_name, rows = row_start_stub,
      #   cols = 1:ncol(ordered_gt_data), style = column_labels_border
      # )

      wb <- openxlsx2::wb_add_cell_style(
          wb = wb, sheet = sheet_name,
          dims = openxlsx2::wb_dims(
            rows = row_start_stub,
            cols = seq(ncol(ordered_gt_data))
          ),
          style = column_labels_border
      )

      label_groups_row[[i]] <- row_start_stub

      row_start_stub <- row_start_stub + 1

      data_to_write <- ordered_gt_data[gt_table$`_stub_df`$group_id == i, ]

      wb <- openxlsx2::wb_add_data(
        wb = wb,
        sheet = sheet_name,
        x = data_to_write,
        start_row = row_start_stub,
        col_names = FALSE#,
        # FIXME: Figure out how to implement borders
        # borders = "surrounding",
        # borderColour = "#D3D3D3"
      )

      cell_data_row[[i]][["start"]] <- row_start_stub
      cell_data_row[[i]][["end"]] <- row_start_stub + nrow(data_to_write) - 1

      row_start_stub <- row_start_stub + nrow(data_to_write)
    }
  }

  wb
  # return(cell_data_row)
}
