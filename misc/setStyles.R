
setStyles <- function(wb){
  #' register styles and return named list
  #' of them
  
  font_size <- 18
  font_name <- "Source Sans Pro"
  
  Background <-  CellStyle(wb) + Fill(backgroundColor='#f4f4f4');
  
  
  Header_Style <- CellStyle(wb) + Font(wb, heightInPoints=14, 
                                       name = font_name)
                  + Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT") 
                  + Border(color="black", position="BOTTOM", pen="BORDER_THICK");
  
  print('xoxo2');
  
  odd_row <-  CellStyle(wb) + Font(wb, heightInPoints=14, name = font_name)
                            + Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT");
  
  print('xoxo3');
  
  even_row <- CellStyle(wb) + Font(wb,  heightInPoints=14, name = font_name)
              + Alignment(wrapText=TRUE, horizontal="ALIGN_RIGHT")
              + Fill(backgroundColor='#f4f4f4');
  
  Title_Style <- CellStyle(wb) + Font(wb,  heightInPoints=18,
                                      name = font_name) 
                               + Fill(backgroundColor='#f4f4f4');
  
  print('xoxo5');
  
  
  return(
        list(title=Title_Style,
              background = Background,
              table = list(
                          header=Header_Style,
                          even_row=even_row,
                          odd_row=odd_row
                          )
              )
         )
}