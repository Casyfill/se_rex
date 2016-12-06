# se_rex
se_excel r writer

# INSTALLATION

- `library(devtools); install_git("casyfill/se_rex");`
```
library(devtools); 
install_git("https://github.com/Casyfill/se_rex.git");
library(serex)
```
# USAGE

```
z <- data.frame(iris)

sheets <- createSE() # create WB object and styles. sheets is the named list of sheets, obviously
addDataSheet(z, name="The famous Iris",sheets[['SUMMARY']]) 
saveSE(wb, "text.xlsx") # store xlsx 

```

# THE ROADMAP

- [x] Background style
- [x] Title Style
- [X] header style
- [x] border style
- [x] background
- [x] title
- [x] Logo
- [x] Sheet drawer
- [x] Store dataframe
- [x] Margins
- [x] and header
- [x] Shadow
- [x] Wrap into the package
- [x] even and odd rows
- [ ] Reference
- [ ] stackSheet
- [ ] Improve interface
- [ ] format style
- [ ] Infer Format
- [ ] Matrix/Stack of Sheets
- [ ] Formulas?
