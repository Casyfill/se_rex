#' sets table rows to colored formats,
#' @param startRow: int, row for the upper left corner of the datasheet
#' @param startCol: int, column for the upper left corner of the datasheet
#' @param shape: vector of 2, dataframe lengiht and width
#' @param sheet: sheet object to write on
#' @return nothing
#' @export
#' @include setStyles.R
setTableRows <- function(startRow, startCol, shape, sheet){

  rowStart <- start_pos[[1]]
  colStart <- start_pos[[2]]
  height <- shape[[1]]
  wight <- shape[[2]]


  # apply odd rows format
  rows <-xlsx::getRows(sheet,rowIndex= rowStart:(rowStart+height-1))
  cells <-xlsx::getCells(rows, colIndex=colStart:(colStart + width - 1))
  mydummy <- lapply(cells, function(cell){mydummy <- xlsx::setCellStyle(cell, styles[['table']][['odd_row']])})

  rowStart <- 5
  height <- 12
  for(i in seq.int( (rowStart+1), (rowStart+floor(height/2)*2) , by=2)){
    rows <-xlsx::getRows(sheet,rowIndex=i)
    cells <-xlsx::getCells(rows, colIndex=colStart:(colStart + width - 1))
    mydummy <- lapply(cells, function(cell){mydummy <- xlsx::setCellStyle(cell, styles[['table']][['even_row']])})
  }

}
