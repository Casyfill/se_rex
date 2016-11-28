#' retirn a set of 4 boder edge styles
#' and 4 border cornet styles
#' + shadow style
#' @param workbook object
#' @keywords style
#' @family internal
#'
borderStyles <- function(wb){

  top <- xslx::CellStyle(wb) + xslx::Border(color="black", position="TOP", pen="BORDER_THICK") +
    xslx::Fill(foregroundColor='white');
  bottom <- xslx::CellStyle(wb) + xslx::Border(color="black", position="BOTTOM", pen="BORDER_THICK") +
    xslx::Fill(foregroundColor='white');
  left <- xslx::CellStyle(wb) + xslx::Border(color="black", position="LEFT", pen="BORDER_THICK") +
    xslx::Fill(foregroundColor='white');
  right <- xslx::CellStyle(wb) + xslx::Border(color="black", position="RIGHT", pen="BORDER_THICK") +
    xslx::Fill(foregroundColor='white');

  ul_corner <-  xslx::CellStyle(wb) + xslx::Border(color="black", position=c("TOP", 'LEFT'), pen="BORDER_THICK") +
    xslx::Fill(foregroundColor='white');
  ur_corner <-  xslx::CellStyle(wb) + xslx::Border(color="black", position=c("TOP", 'RIGHT'), pen="BORDER_THICK") +
    xslx::Fill(foregroundColor='white');
  bl_corner <- xslx::CellStyle(wb) + xslx::Border(color="black", position=c("BOTTOM", 'LEFT'), pen="BORDER_THICK") +
    xslx::Fill(foregroundColor='white');
  br_corner <- CellStyle(wb) + xslx::Border(color="black", position=c("BOTTOM", 'RIGHT'), pen="BORDER_THICK")  +
    xslx::Fill(foregroundColor='white');

  shadow <-   xslx::CellStyle(wb) + xslx::Fill(foregroundColor='#969696');


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
