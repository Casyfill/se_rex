# SE_REX is a R to excel writer for Streeteasy,
# A tiny wrapper that helps us to spend time on research, not formatting
# Written by Philipp Kats, 2016_11_24

# TINY ROADMAP

# [x] Background style
# [x] Title Style
# [X] header style
# [x] border style
# [x] background
# [x] title
# [x] Logo 
# [x] Sheet drawer
# [x] Store dataframe
# [x] Margins
# [x] and header
# [ ] even and odd rows
# [ ] format style
# Infer Format
# [x] Shadow
# Reference
# Wrap into the package
# Matrix of Sheets

library(xlsx)
library(datasets)

# setwd('/Users/casy/Dropbox/CUSP/project/2_EXPERIMENTS/se_rex')
setwd('/Users/philippk/Dropbox (Zillow Group)/my_projects/se_rex')
source('misc/setStyles.R')
source('misc/set_general.R')


wb<-createWorkbook(type="xlsx")
styles <- setStyles(wb)
summary <- createSheet(wb, sheetName = "SUMMARY")

  {
    xlsx.background(summary, styles[['background']])
    addPicture('resources/SE_logo.png', summary, startRow=2, startColumn=2 )
    xlsx.addTitle(summary, 9, 2, 'Current Inventory:', styles[['title']])
    
        
    # xlsx.sheetborder(sheet, width=5, height=5, rowStart=12, colStart=2 )
    addDataSheet(cars, summary, startRow=12, startCol=2, name='Cars')
    saveWorkbook(wb, "r-xlsx-test.xlsx")
  }

