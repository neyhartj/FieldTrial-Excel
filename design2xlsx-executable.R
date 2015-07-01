##Executing design2xlsx

setwd("C:/Users/Jeff/Documents/University of Minnesota/Barley Lab/Projects/Repositories/plotdesign-excel")

source("design2xlsx-FUNCTION.R")

load("plot_design_excel_data.RData")

design2xlsx(x = Designs_List[[1]], NumLocs = F, FieldDim = F, NumDupl = F)
