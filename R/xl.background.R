#' apply shadowed background
#' @param xlsx sheet object
#' @param  background_style: xlsx style object
#' @keywords styler
#' @family internal
xl.background <- function(sheet, background_style){
  #' create a background color fill
  rows <- xslx::createRow(sheet,rowIndex=1:100)
  cells <- xslx::createCell(rows, colIndex=1:100)

  wrap <- function(cell){
    xslx::setCellStyle(cell, background_style)
  }

  mydummy <- lapply(cells, wrap)
}
