# SE_REX is a R to excel writer for Streeteasy,
# A tiny wrapper that helps us to spend time on research, not formatting
# Written by Philipp Kats, 2016_11_24
library(xlsx)


# TINY ROADMAP

# [x] Background style
# [x] Title Style
# [X]  header style
# [ ] border style
#
# Sheet drawer
# even and odd rows
# format style
# Infer Format
# Shadow
# Reference
# Wrap into the package
# Matrix of Sheets

setwd('/Users/casy/Dropbox/CUSP/project/2_EXPERIMENTS/se_rex')
source('misc/setStyles.R')


wb<-createWorkbook(type="xlsx")
styles <- setStyles(wb)

sheet <- createSheet(wb, sheetName = "US State Facts")


#++++++++++++++++++++++++
# Helper function to add titles
#++++++++++++++++++++++++
# - sheet : sheet object to contain the title
# - rowIndex : numeric value indicating the row to 
#contain the title
# - title : the text to use as title
# - titleStyle : style object to use for title
xlsx.addTitle<-function(sheet, rowIndex, title, titleStyle){
  rows <-createRow(sheet,rowIndex=rowIndex)
  sheetTitle <-createCell(rows, colIndex=1)
  setCellValue(sheetTitle[[1,1]], title)
  setCellStyle(sheetTitle[[1,1]], titleStyle)
}



# Add title
xlsx.addTitle(sheet, rowIndex=1, title="US State Facts",
              titleStyle = TITLE_STYLE)
# Add sub title
xlsx.addTitle(sheet, rowIndex=2, 
              title="Data sets related to the 50 states of USA.",
              titleStyle = SUB_TITLE_STYLE)

head(state.x77)

# Add a table
addDataFrame(state.x77, sheet, startRow=3, startColumn=1, 
             colnamesStyle = TABLE_COLNAMES_STYLE,
             rownamesStyle = TABLE_ROWNAMES_STYLE)
# Change column width
setColumnWidth(sheet, colIndex=c(1:ncol(state.x77)), colWidth=11)


# Save the workbook to a file
saveWorkbook(wb, "r-xlsx-report-example.xlsx")

