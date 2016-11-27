

xlsx.background <- function(sheet, background_style){
  #' create a background color fill
  rows <- createRow(sheet,rowIndex=1:100)
  cells <- createCell(rows, colIndex=1:100)
  
  wrap <- function(cell){
    setCellStyle(cell, background_style)
  }
  
  mydummy <- lapply(cells, wrap)
}

xlsx.addTitle<-function(sheet, rowIndex, colIndex, title, titleStyle){
  rows <-getRows(sheet,rowIndex=rowIndex)
  sheetTitle <-getCells(rows, colIndex=colIndex)
  setCellValue(sheetTitle[[1]], title)
  setCellStyle(sheetTitle[[1]], titleStyle)
}


xlsx.sheetborder <- function(sheet, width, height, rowStart, colStart){
  #' draws a sheet border surrounding the dataframe
  rows <-getRows(sheet,rowIndex=rowStart:rowStart+height)
  cells <-getCells(rows, colIndex=colStart)
  setCellStyle(cells, styles[['border']])
}