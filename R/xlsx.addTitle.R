#' add title to the sheet
#' @param xlsx sheet object
#' @param  rowIndex: vector of integers, difining rows to work on
#' @param  colIndex: vector of integers, difining columns to work on
#' @param  title: string, title to write
#' @param  titleStyle: xlsx style format
#' @keywords styler
#' @family internal
xlsx.addTitle<-function(sheet, rowIndex, colIndex, title, titleStyle){
  rows <-getRows(sheet,rowIndex=rowIndex)
  sheetTitle <-getCells(rows, colIndex=colIndex)
  setCellValue(sheetTitle[[1]], title)
  setCellStyle(sheetTitle[[1]], titleStyle)
}
