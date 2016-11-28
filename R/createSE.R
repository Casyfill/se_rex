#' Main workbook creator
#' NOTE: maybe should return wb only, then get sheets through getSheets()'
#' @description  creates new workbook, registers styles
#' and adds given sheets
#' @param sheetList: named list, list(SheetName='Sheet title')
#' @return named list of WB (workbook) and Sheets (named list of sheets)
#' @return sheets: named list of registered sheets
#' @include setStyles.R xl.background.R xl.addTitle.R
#' @export
createSE <- function(sheetlist=list(SUMMARY='Neighborhood inventory')){

  wb<<-xslx::createWorkbook(type="xlsx")
  styles <<- setStyles(wb)

  logo_path <-  system.file(system.file("resources/SE_logo.png", package="serex"), package="serex")
  sheets = list()
  for(sheet.name in names(sheetlist)){

    sheet <- xslx::createSheet(wb, sheetName = sheet.name)
    xl.background(sheet, styles[['background']])

    xslx::addPicture(logo_path, sheet, startRow=2, startColumn=2 )
    xl.addTitle(sheet, 9, 2, sheetlist[[sheet.name]], styles[['title']] )
    sheets[[sheet.name]]= sheet
  }

  return(sheets)
}
