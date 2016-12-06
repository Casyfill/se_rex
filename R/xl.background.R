#' apply shadowed background
#' @param xlsx sheet object
#' @param  background_style: xlsx style object
#' @keywords styler
#' @family internal
xl.background <- function(sheet, background_style){
  #' create a background color fill
  rows <- xlsx::createRow(sheet,rowIndex=1:1000)
  cells <- xlsx::createCell(rows, colIndex=1:200)

  wrap <- function(cell){
    xlsx::setCellStyle(cell, background_style)
  }

  mydummy <- lapply(cells, wrap)
}
