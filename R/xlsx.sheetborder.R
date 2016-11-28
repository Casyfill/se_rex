#' draws a sheet border surrounding the dataframe
#' @param xlsx sheet object
#' @param  wight: integer, wight of the sheet, edge to edge
#' @param  height: integer, height of the sheet, edge to edge
#' @param  rowStart: integer, row for the upper left corner
#' @param  colStart: integer, col for the upper left corner
#' @keywords styler
#' @family internal
xlsx.sheetborder <- function(sheet, width, height, rowStart, colStart){


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

