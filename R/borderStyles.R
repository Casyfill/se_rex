#' retirn a set of 4 boder edge styles
#' and 4 border cornet styles
#' + shadow style
#' @param workbook object
#' @keywords style
#' @family internal
#'
borderStyles <- function(wb){

  top <- xlsx::CellStyle(wb) + xlsx::Border(color="black", position="TOP", pen="BORDER_THICK") +
    xlsx::Fill(foregroundColor='white');
  bottom <- xlsx::CellStyle(wb) + xlsx::Border(color="black", position="BOTTOM", pen="BORDER_THICK") +
    xlsx::Fill(foregroundColor='white');
  left <- xlsx::CellStyle(wb) + xlsx::Border(color="black", position="LEFT", pen="BORDER_THICK") +
    xlsx::Fill(foregroundColor='white');
  right <- xlsx::CellStyle(wb) + xlsx::Border(color="black", position="RIGHT", pen="BORDER_THICK") +
    xlsx::Fill(foregroundColor='white');

  ul_corner <-  xlsx::CellStyle(wb) + xlsx::Border(color="black", position=c("TOP", 'LEFT'), pen="BORDER_THICK") +
    xlsx::Fill(foregroundColor='white');
  ur_corner <-  xlsx::CellStyle(wb) + xlsx::Border(color="black", position=c("TOP", 'RIGHT'), pen="BORDER_THICK") +
    xlsx::Fill(foregroundColor='white');
  bl_corner <-  xlsx::CellStyle(wb) + xlsx::Border(color="black", position=c("BOTTOM", 'LEFT'), pen="BORDER_THICK") +
    xlsx::Fill(foregroundColor='white');
  br_corner <-  xlsx::CellStyle(wb) + xlsx::Border(color="black", position=c("BOTTOM", 'RIGHT'), pen="BORDER_THICK") +
    xlsx::Fill(foregroundColor='white');

  shadow <-   xlsx::CellStyle(wb) + xlsx::Fill(foregroundColor='#969696');


  return(
    list(top=top,
         bottom=bottom,
         left=left,
         right=right,
         ul_corner = ul_corner,
         ur_corner = ur_corner,
         bl_corner = bl_corner,
         br_corner = br_corner,
         shadow = shadow
    )
  )
}
