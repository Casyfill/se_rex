#' @description  creates new workbook, registers styles
#' and adds given sheets NOTE: maybe should return wb only, then get sheets through getSheets()'
#' @param sheetList: named list, list(SheetName='Sheet title')
#' @return named list of WB (workbook) and Sheets (named list of sheets)
#' @return sheets: named list of registered sheets
#' @include setStyles.R xlsx.background.R xlsx.addTitle.R
#' @export
createSE <- function(sheetlist=list(SUMMARY='Neighborhood inventory')){

  wb<<-createWorkbook(type="xlsx")
  styles <<- setStyles(wb)

  logo_path <-  system.file(file.path('data', 'SE_logo.png'), package="serex")
  sheets = list()
  for(sheet.name in names(sheetlist)){

    sheet <- createSheet(wb, sheetName = sheet.name)
    xlsx.background(sheet, styles[['background']])

    addPicture(logo_path, sheet, startRow=2, startColumn=2 )
    xlsx.addTitle(sheet, 9, 2, sheetlist[[sheet.name]], styles[['title']] )
    sheets[[sheet.name]]= sheet
  }

  return(sheets)
}
