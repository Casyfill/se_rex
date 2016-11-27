

xlsx.background <- function(sheet, background_style){
  #' create a background color fill
  rows <- createRow(sheet,rowIndex=1:100)
  cells <- createCell(rows, colIndex=1:100)
  
  wrap <- function(cell){
    setCellStyle(cell, background_style)
  }
  
  mydummy <- lapply(cells, wrap)
}

xlsx.addTitle<-function(sheet, rowIndex, colIndex, title, titleStyle){
  rows <-getRows(sheet,rowIndex=rowIndex)
  sheetTitle <-getCells(rows, colIndex=colIndex)
  setCellValue(sheetTitle[[1]], title)
  setCellStyle(sheetTitle[[1]], titleStyle)
}



xlsx.sheetborder <- function(sheet, width, height, rowStart, colStart){
  #' draws a sheet border surrounding the dataframe
  # 
  
  {
    wrap_left <- function(cell){
      mydummy <- setCellStyle(cell, styles[['borders']][['left']])
    }
    
    wrap_right <- function(cell){
      mydummy <- setCellStyle(cell, styles[['borders']][['right']])
    }
    
    
    wrap_top <- function(cell){
      mydummy <- setCellStyle(cell, styles[['borders']][['top']])
    }
    
    wrap_bottom <- function(cell){
      mydummy <- setCellStyle(cell, styles[['borders']][['bottom']])
    }
    
  }
  
  { # left edge
    rows <-getRows(sheet,rowIndex=(rowStart+1):(rowStart+height-1))
    cells <-getCells(rows, colIndex=colStart)
    mydummy <- lapply(cells, wrap_left)
  }
  
  { # right edge
    rows <-getRows(sheet,rowIndex=(rowStart+1):(rowStart+height-1))
    cells <-getCells(rows, colIndex=colStart+width)
    mydummy <- lapply(cells, wrap_right)
  }
  
  { # top edge
    rows <-getRows(sheet,rowIndex=rowStart)
    cells <-getCells(rows, colIndex=(colStart+1):(colStart+width-1))
    mydummy <- lapply(cells, wrap_top)
  }
  
  { #  bottom
    rows <-getRows(sheet,rowIndex=(rowStart+height))
    cells <-getCells(rows, colIndex=(colStart+1):(colStart+width-1))
    mydummy <- lapply(cells, wrap_bottom)
  }
  
  # ul corner
  {
    cell <-getCells(getRows(sheet,rowIndex=rowStart), colIndex=colStart)
    mydummy <-setCellStyle(cell[[1]],styles[['borders']][['ul_corner']])
    
    cell <-getCells(getRows(sheet,rowIndex=rowStart), colIndex=(colStart+width))
    mydummy <-setCellStyle(cell[[1]],styles[['borders']][['ur_corner']])
    
    cell <-getCells(getRows(sheet,rowIndex=(rowStart+height)), colIndex=(colStart))
    mydummy <-setCellStyle(cell[[1]],styles[['borders']][['bl_corner']])
    
    cell <-getCells(getRows(sheet,rowIndex=(rowStart+height)), colIndex=(colStart+width))
    mydummy <-setCellStyle(cell[[1]],styles[['borders']][['br_corner']])
  }
  
  { #shadow
    wrap_shadow <- function(cell){
      mydummy <- setCellStyle(cell, styles[['borders']][['shadow']])
    }
    
    rows <-getRows(sheet, rowIndex=rowStart+height+1 )
    cells <-getCells(rows, colIndex=colStart:(colStart+width))
    mydummy <- lapply(cells, wrap_shadow)
    
    rows <-getRows(sheet, rowIndex=(rowStart+1):(rowStart+height+1) )
    cells <-getCells(rows, colIndex=(colStart+width+1))
    mydummy <- lapply(cells, wrap_shadow)
  }
 
}

