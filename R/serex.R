#!/usr/bin/Rscript --vanilla
# SE_REX is a R to excel writer for Streeteasy,
# A tiny wrapper that helps us to spend time on research, not formatting
# Written by Philipp Kats, 2016_11_24

# TINY ROADMAP

# [x] Background style
# [x] Title Style
# [X] header style
# [x] border style
# [x] background
# [x] title
# [x] Logo 
# [x] Sheet drawer
# [x] Store dataframe
# [x] Margins
# [x] and header
# [x] Shadow
# [ ] Wrap into the package
# [ ] even and odd rows
# [ ] format style
# [ ] Infer Format
# [ ] Reference
# [ ] Matrix/Stack of Sheets
# [ ] Formulas?

library(xlsx)

setwd("./R")
source(file.path('misc','setStyles.R'))
source(file.path('misc','set_general.R'))

createSE <- function(sheetlist=list(SUMMARY='Neighborhood inventory')){
  #' creates new workbook, registers styles
  #' and adds given sheets
  #' NOTE: maybe should return wb only, then get sheets through getSheets()
  #' @param sheets: named list, list(SheetName='Sheet title')
  #' @return named list of WB (workbook) and Sheets (named list of sheets)
  #' @export

  wb<<-createWorkbook(type="xlsx")
  styles <<- setStyles(wb)
  
  sheets = list()
  for(sheet.name in names(sheetlist)){

    sheet <- createSheet(wb, sheetName = sheet.name)
    xlsx.background(sheet, styles[['background']])
    addPicture('../resources/SE_logo.png', sheet, startRow=2, startColumn=2 )
    xlsx.addTitle(sheet, 9, 2, sheetlist[[sheet.name]], styles[['title']] )
    sheets[[sheet.name]]= sheet
  }
  
  return(sheets)
}


addDataSheet <- function(df, name='DataSheet', sheet, startRow=12, startCol=2, topspace=3){
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
  
  
  dfshape <- c(nrow(df)+2+topspace, ncol(df)+1) # infer size of the sheet
  
  # add dataframe
  addDataFrame(x=df, sheet=sheet, startRow=(startRow+1+topspace), 
               startCol=(startCol+1), showNA=T, row.names = F,
               characterNA="NA", # TODO: colStyle = NULL,  colSTYLE
               colnamesStyle=styles[['table']][['header']])
  
  # add sheet borders
  xlsx.sheetborder(sheet=sheet, width=dfshape[2], height=dfshape[1], rowStart=startRow, colStart=startCol)
  
  wrap_blank <- function(cell){
    mydummy <- setCellStyle(cell, styles[['table']][['odd_row']])
  }
  
  rows <-getRows(sheet, rowIndex=(startRow+1):(startRow+topspace) )
  cells <-getCells(rows, colIndex=(startCol+1):(startCol+dfshape[2]-1))
  mydummy <- lapply(cells, wrap_blank)
  
  cell <-getCells(getRows(sheet,rowIndex=(startRow+1)), colIndex=(startCol+1))
  mydummy <-setCellValue(cell[[1]],name)
  
}



saveSE <-function(wb, path){
  #' store SE workbook as an excel file 
  #' @param wb: workbook object to save
  #' @param path: string, path to the file
  #' @export
  saveWorkbook(wb, path)
}



if(interactive()){
  library(datasets) # get example
  sheetlist <- list(SUMMARY='Neighborhood inventory')
  sheets <- createSE(sheetlist=sheetlist)
  addDataSheet(cars, name='Cars Sample Dataset', sheets[['SUMMARY']], startRow=12, startCol=2, topspace=3)
  saveSE(wb, 'test.xlsx')
}

  
