#' sets table rows to colored formats,
#' @param startRow: int, row for the upper left corner of the datasheet
#' @param startCol: int, column for the upper left corner of the datasheet
#' @param shape: vector of 2, dataframe lengiht and width
#' @param sheet: sheet object to write on
#' @return nothing
#' @export
#' @include setStyles.R
setTableRows <- function(startRow, startCol, shape, sheet){

  startRow <- startRow + 1 # account for header
  height <- shape[[1]]
  width <- shape[[2]]


  # apply odd rows format
  rows <-xlsx::getRows(sheet,rowIndex= startRow:(startRow+height-1))
  cells <-xlsx::getCells(rows, colIndex=startCol:(startCol + width - 1))
  mydummy <- lapply(cells, function(cell){mydummy <- xlsx::setCellStyle(cell, styles[['table']][['odd_row']])})

  startRow <- 5
  height <- 12
  for(i in seq.int( (startRow+1), (startRow+floor(height/2)*2) , by=2)){
    rows <-xlsx::getRows(sheet,rowIndex=i)
    cells <-xlsx::getCells(rows, colIndex=startCol:(startCol + width - 1))
    mydummy <- lapply(cells, function(cell){mydummy <- xlsx::setCellStyle(cell, styles[['table']][['even_row']])})
  }

}
