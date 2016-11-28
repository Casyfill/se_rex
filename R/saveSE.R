#' store SE workbook as an excel file
#' @param wb: workbook object to save
#' @param path: string, path to the file
#' @export
saveSE <-function(wb, path){
  saveWorkbook(wb, path)
}
