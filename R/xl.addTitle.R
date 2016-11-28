#' add title to the sheet
#' @param xlsx sheet object
#' @param  rowIndex: vector of integers, difining rows to work on
#' @param  colIndex: vector of integers, difining columns to work on
#' @param  title: string, title to write
#' @param  titleStyle: xlsx style format
#' @keywords styler
#' @family internal
xl.addTitle<-function(sheet, rowIndex, colIndex, title, titleStyle){
  rows <-xslx::getRows(sheet,rowIndex=rowIndex)
  sheetTitle <-xslx::getCells(rows, colIndex=colIndex)
  xlsx::setCellValue(sheetTitle[[1]], title)
  xlsx::setCellStyle(sheetTitle[[1]], titleStyle)
}
