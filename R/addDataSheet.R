#' writes a dataframe to the chosen sheet,
#' format it as a "datasheet"
#' @param x: dataframe to store
#' @param name: title of the datasheet to write
#' @param sheet: sheet object to write on
#' @param startRow: int, row for the upper left corner of the datasheet
#' @param startCol: int, column for the upper left corner of the datasheet
#' @param topspace: int, space between the top edge of the sheet and the header of the table
#' @return nothing
#' @export
#' @include xl.background.R xl.sheetborder.R setTableRows.r
addDataSheet <- function(df, name='DataSheet', sheet, startRow=12, startCol=2, topspace=3){


  dfshape <- c(nrow(df)+2+topspace, ncol(df)+1) # infer size of the sheet

  # add dataframe
  addDataFrame(x=df, sheet=sheet, startRow=(startRow+1+topspace),
               startCol=(startCol+1), showNA=T, row.names = F,
               characterNA="NA", # TODO: colStyle = NULL,  colSTYLE
               colnamesStyle=styles[['table']][['header']])

  setTableRows(startRow=12, startCol=2, dim(df), sheet)
  # add sheet borders
  xl.sheetborder(sheet=sheet, width=dfshape[2], height=dfshape[1], rowStart=startRow, colStart=startCol)

  wrap_blank <- function(cell){
    mydummy <- setCellStyle(cell, styles[['table']][['odd_row']])
  }

  rows <-getRows(sheet, rowIndex=(startRow+1):(startRow+topspace) )
  cells <-getCells(rows, colIndex=(startCol+1):(startCol+dfshape[2]-1))
  mydummy <- lapply(cells, wrap_blank)

  cell <-getCells(getRows(sheet,rowIndex=(startRow+1)), colIndex=(startCol+1))
  mydummy <-setCellValue(cell[[1]],name)

}
