#' retirn a set of 4 boder edge styles
#' and 4 border cornet styles
#' + shadow style
#' @param workbook object
#' @keywords style
#' @family internal
#' @return named list of styles
#' @include borderStyles.R
setStyles <- function(wb){
  #' register styles and return named list
  #' of them

  font_size <- 18
  font_name <- "Source Sans Pro"

  Background <-  CellStyle(wb) + Fill(foregroundColor='#f4f4f4');



  Header_Style <- CellStyle(wb) + Font(wb, heightInPoints=14,
                                       name=font_name) +
                  Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT") +
                  Border(color="black", position="BOTTOM", pen="BORDER_THIN") +
                  Fill(foregroundColor='white');



  odd_row <-  CellStyle(wb) + Font(wb, heightInPoints=14, name = font_name) +
              Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT") +
              Fill(foregroundColor='white');



  even_row <- CellStyle(wb) + Font(wb,  heightInPoints=14, name = font_name) +
              Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT") +
              Fill(foregroundColor='#f4f4f4');

  Title_Style <- CellStyle(wb) + Font(wb,  heightInPoints=18,
                                      name = font_name) +
                               Fill(foregroundColor='#f4f4f4');

  borders <- borderStyles(wb)



  return(
        list(title=Title_Style,
              background = Background,
              borders = borders,
              table = list(
                          header=Header_Style,
                          even_row=even_row,
                          odd_row=odd_row
                          )

              )
         )
}

