# SE_REX is a R to excel writer for Streeteasy,
# A tiny wrapper that helps us to spend time on research, not formatting
# Written by Philipp Kats, 2016_11_24
library(xlsx)


# TINY ROADMAP

# [x] Background style
# [x] Title Style
# [X] header style
# [x] border style
# [x] background
# [x] title
# [x] Logo 
# [ ] Sheet drawer
#   [ ] 
# even and odd rows
# format style
# Infer Format
# Shadow
# Reference
# Wrap into the package
# Matrix of Sheets

# setwd('/Users/casy/Dropbox/CUSP/project/2_EXPERIMENTS/se_rex')
setwd('/Users/philippk/Dropbox (Zillow Group)/my_projects/se_rex')
source('misc/setStyles.R')
source('misc/set_general.R')


wb<-createWorkbook(type="xlsx")
styles <- setStyles(wb)
names(styles)
sheet <- createSheet(wb, sheetName = "Borough")

{
  xlsx.background(sheet, styles[['background']])
  addPicture('resources/SE_logo.png', sheet, startRow=2, startColumn=2 )
  xlsx.addTitle(sheet, 9, 2, 'Current Inventory:', styles[['title']])
  saveWorkbook(wb, "r-xlsx-test.xlsx")
}

# xlsx.addTitle(sheet, 9, 2, 'Hello World!', styles[['Title_Style']])

# Add title
# xlsx.addTitle(sheet, rowIndex=1, title="US State Facts",
              # titleStyle = TITLE_STYLE)
# Add sub title
# xlsx.addTitle(sheet, rowIndex=2, 
              # title="Data sets related to the 50 states of USA.",
              # titleStyle = SUB_TITLE_STYLE)

# head(state.x77)
s
# Add a table
# addDataFrame(state.x77, sheet, startRow=3, startColumn=1, 
#              colnamesStyle = TABLE_COLNAMES_STYLE,
#              rownamesStyle = TABLE_ROWNAMES_STYLE)
# # Change column width
# setColumnWidth(sheet, colIndex=c(1:ncol(state.x77)), colWidth=11)
# 

# Save the workbook to a file
rows <-getRows(sheet,rowIndex=9 )
getCells(rows, colIndex=B)
