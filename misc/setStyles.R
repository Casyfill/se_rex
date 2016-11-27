

borderStyles <- function(wb){
  #' retirn a set of 4 boder edge styles
  #' and 4 border cornet styles
  #' + shadow style
  
  top <- CellStyle(wb) + Border(color="black", position="TOP", pen="BORDER_THICK");
  bottom <- CellStyle(wb) + Border(color="black", position="BOTTOM", pen="BORDER_THICK");
  left <- CellStyle(wb) + Border(color="black", position="LEFT", pen="BORDER_THICK");
  right <- CellStyle(wb) + Border(color="black", position="RIGHT", pen="BORDER_THICK");
  
  ul_corner <-  CellStyle(wb) + Border(color="black", position=c("TOP", 'LEFT'), pen="BORDER_THICK");
  ur_corner <-  CellStyle(wb) + Border(color="black", position=c("TOP", 'RIGHT'), pen="BORDER_THICK");
  bl_corner <- CellStyle(wb) + Border(color="black", position=c("BOTTOM", 'LEFT'), pen="BORDER_THICK");
  br_corner <- CellStyle(wb) + Border(color="black", position=c("BOTTOM", 'RIGHT'), pen="BORDER_THICK");
  
  shadow <-   CellStyle(wb) + Fill(foregroundColor='#969696');
  
  
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


setStyles <- function(wb){
  #' register styles and return named list
  #' of them
  
  font_size <- 18
  font_name <- "Source Sans Pro"
  
  Background <-  CellStyle(wb) + Fill(foregroundColor='#f4f4f4');
  
  
  
  Header_Style <- CellStyle(wb) + Font(wb, heightInPoints=14, 
                                       name=font_name) +
                  Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT") +
                  Border(color="black", position="BOTTOM", pen="BORDER_THICK");
  
  
  
  odd_row <-  CellStyle(wb) + Font(wb, heightInPoints=14, name = font_name) +
                            Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT");
  
  
  
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
              borders <- borders,
              table = list(
                          header=Header_Style,
                          even_row=even_row,
                          odd_row=odd_row
                          )
              )
         )
}

