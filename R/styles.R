
title_style <- openxlsx::createStyle(halign = "center", border = "TopLeftRight", borderColour = "#D3D3D3")

# column_labels_border <- openxlsx::createStyle(
#   border = "TopLeftRight",
#   borderColour = "#D3D3D3"
# )

column_labels_border <- openxlsx2::create_border(
  top = "medium",
  top_color = openxlsx2::wb_color("#D3D3D3"),
  left = "medium",
  left_color = openxlsx2::wb_color("#D3D3D3"),
  right = "medium",
  right_color = openxlsx2::wb_color("#D3D3D3")
)

percent_style <- openxlsx::createStyle(numFmt = "0%")

sourcenote_style <- openxlsx::createStyle(fontSize = 10)
