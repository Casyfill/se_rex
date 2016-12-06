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

  Background <-  xlsx::CellStyle(wb) + xlsx::Fill(foregroundColor='#f4f4f4');


  Header_Style <- xlsx::CellStyle(wb) + xlsx::Font(wb, heightInPoints=14,
                                       name=font_name) +
    xlsx::Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT") +
    xlsx::Border(color="black", position="BOTTOM", pen="BORDER_THIN") +
    xlsx::Fill(foregroundColor='white');

  odd_row <-  xlsx::CellStyle(wb) + xlsx::Font(wb, heightInPoints=14, name = font_name) +
    xlsx::Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT") +
    xlsx::Fill(foregroundColor='white');

  even_row <- xlsx::CellStyle(wb) + xlsx::Font(wb,  heightInPoints=14, name = font_name) +
    xlsx::Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT") +
    xlsx::Fill(foregroundColor='#f4f4f4');


  Title_Style <- xlsx::CellStyle(wb) + xlsx::Font(wb,  heightInPoints=18,
                                      name = font_name) +
    xlsx::Fill(foregroundColor='#f4f4f4');

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

