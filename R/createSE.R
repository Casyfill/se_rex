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

  wb<<-xlsx::createWorkbook(type="xlsx")
  styles <<- setStyles(wb)

  logo_path <-  system.file(file.path('resources','SE_logo.png'), package="serex", mustWork = TRUE)
  # print(logo_path)
  sheets = list()
  for(sheet.name in names(sheetlist)){

    sheet <- xlsx::createSheet(wb, sheetName = sheet.name)
    xl.background(sheet, styles[['background']])

    xlsx::addPicture(logo_path, sheet, startRow=2, startColumn=2 )
    xl.addTitle(sheet, 9, 2, sheetlist[[sheet.name]], styles[['title']] )
    sheets[[sheet.name]]= sheet
  }

  return(sheets)
}
