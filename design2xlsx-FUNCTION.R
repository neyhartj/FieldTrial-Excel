###
#Title: Plot design output to formatted Excel file
#Working function name: design2xlsx
#Author: Jeff Neyhart
#Purpose: This code takes the output from the FldTrial package and creates a formatted Excel file wherein the first three data.frames 
##are positioned for effective use by the U of M Barley Breeding Group. This code is meant to be an extension of that package
#Version: 1.6
#Now on GitHub!
#Latest updates:
## Version 1.6: Improved functionality with RCBD designs and wrapped up the script into a function
## v 1.5.3: Cleaned up functionality for adding the .dgn data.frame
## v 1.5.2: Cleaned up variables
## v 1.5.1: Added functionality to highlight the checks in the design data.frame, the plot layout data.frame, and the check.layout data.frame
#Upcoming features:



##### Function creation and usage #####
# Function creation and argument designation
design2xlsx <- function (
  x, #The list of design outputs from the FldTrial package
  filename = "", #Filename to save the workbook as. Need to check to make sure .xlsx is appended as the file extension; defaults to "design_output" + system date/time
  SheetNames = NULL, #Optional character vector of names to call the sheets in the workbook. Must adhere to excel sheet name rules (i.e. no special characters). Defaults to the environment name of each list. Length must be the same as the length of the list
  CheckCol = NULL, #Optional character vector of color names for the checks in order of the check number orders (defaults to the first n colors of rainbow)
  BlkBorders = TRUE, #Optional logical indicating whether borders should be drawn around the blocks
  Source = TRUE, #Optional logical indicating whether a "source" column should be included. Defaults to the 5th column
  DgnNames = NULL, #Optional numeric vector of numbers indicating the parameters and order of such parameters to include in the dgn data.frame section; all parameters are available for all three designs (aibd, madii, rcbd) except where indicated; defaults to  (1 = environment, 2 = trial, 3 = plot.id, 4 = plot, 5 = rep, 6 = row, 7 = column, 8 = blk (aibd, madii), 9 = row.blk (aibd, madii), 10 = col.blk, 11 = line.code, 12 = line.name, 13 = entry, 14 = duplicates)
  SheetTitle = NULL, #Optional vector of names for the title within each sheet (e.g. "2015 S2TP St. Paul Randomization"). Defaults to the trial name
  NumLocs = TRUE, #Optional logical indicating whether the Number of Locations subtitle should be included
  DgnType = TRUE, #Optional logical indicating whether the Design Type (e.g. AIBD) subtitle should be included
  TotalPlots = TRUE, #Logical indicating whether the Total Test Plots subtitle should be included
  FieldDim = TRUE, #Logical indicating whether the Field Dimensions subitle should be included
  NumDupl = TRUE #Logical indicating whether the Number of Duplicates subtitle should be included  
) {


  #Package requirements
  require(xlsx)
  
  ##### Pre-Pre-Formatting#####
  #Reassign x to the name Designs_List
  Designs_List <- x
  #Remove x
  rm(x)
  
   
  ##### Usage checks, errors, return #####
  if (filename == "") {filename = paste("design_output_", format(Sys.time(), "%d-%m-%Y-%H-%M-%S"), ".xlsx", sep = "")}
  #Check to make sure the file name has the ".xlsx" extension
  if (grepl(pattern = ".xlsx", x = filename) == FALSE) stop("File does not have the .xlsx file extension")
  #Make sure the entry file is a list
  if (!is.list(Designs_List)) stop("The x argument is not a list!")
  #CheckCol errors
  if (!is.null(CheckCol)) {
    #Make sure that the length of the user-defined CheckCol vector is the same as the number of checks
    if (length(CheckCol) != max(Designs_List[[1]][[1]]$line_code)) stop("The CheckCol vector is not of the same length as the number of checks")
    #Make sure that the vector is a character
    if (!is.character(CheckCol)) stop("The CheckCol vector is not of class character")
  }
  #SheetNames errors
  if (!is.null(SheetNames)) {
    #Make sure the vector is of appropriate length
    if (length(SheetNames) != length(Designs_List)) stop("The length of the SheetNames vector is not the same as the length of the x list input")
    #Make sure the vector is character
    if (!is.character(SheetNames)) stop("The SheetNames vector is not of class character")
  }
  #BlkBorders errors
  if (!is.logical(BlkBorders)) stop("The BlkBorders argument is not of class logical")
  #Source errors
  if (!is.logical(Source)) stop("The Source argument is not of class logical")
  #DgnNames errors
  if (!is.null(DgnNames)) {
    if (!is.numeric(DgnNames)) stop("The DgnNames vector is not of class numeric")
  }
  #SheetTitle errors
  if (!is.null(SheetTitle)) {
    if (length(SheetTitle != legnth(Designs_List))) stop("The SheetTitle vector is not of the same length as the x list input")
    if (!is.character(SheetTitle)) stop("The SheetTitle vector is not of class character")
  }
  #NumLocs errors
  if (!is.logical(NumLocs)) stop("The NumLocs argument is not of class logical")
  #DgnType errors
  if (!is.logical(DgnType)) stop("The DgnType argument is not of class logical")
  #TotalPlots errors
  if (!is.logical(TotalPlots)) stop("The TotalPlots argument is not of class logical")
  #FieldDim errors
  if (!is.logical(FieldDim)) stop("The FieldDim argument is not of class logical")
  #NumDupl errors
  if (!is.logical(NumDupl)) stop("The NumDupl argument is not of class logical")

  #Test if the list is just a list of one design plan, or is a list of lists. If it is the former, add another level to the list for compliance
  if (is.factor(Designs_List[[1]][[1]])) {
    Designs_List <- list(Designs_List)
  }
  
  
  
  ##### Workbook setup #####
  #Create the new workbook
  pdwb <- createWorkbook()
  
  #Create a list for the sheets
  sheetlist <- list()
  
  #If the SheetTitle is not user-defined, set it to the default
  if (is.null(SheetTitle)) {
    for (l in 1:length(Designs_List)) {
    SheetTitle[l] <- as.character(unique(Designs_List[[l]][[1]]$trial))
    }
  }
  
  #Loop through each design output in the list
  for (l in 1:length(Designs_List)) { 
  
    #If the DgnNames vector is not defined, set it to the default
    if (is.null(DgnNames)) {
      DgnNames <- c(1,4,12,11,6,7,8,9,10,14)
    }
    
    #Rename the .dgn data.frame in the list member l so that there are easier to work with
    #TALK TO TYLER ABOUT THE NAMING OUTPUT SO THAT THIS IS UNECESSARY
    names(Designs_List[[l]][[1]]) <- sub(pattern = "_", replacement = ".", x = names(Designs_List[[l]][[1]]))
        
    ###### Pre-formatting #####
    #Set the names of the Designs_List list to the environment
    for (p in 1:length(Designs_List)) {
      names(Designs_List)[p] <- as.character(unique(Designs_List[[p]][[1]]$environment))
    }
    #Add "Row" to the column containing row numbers for the plot layout data.frame
    #Using an if statement to check to see if "Row" had already been added
    if (length(grep(pattern = "Row", x = Designs_List[[l]][[3]][,1][1])) == 0) {
      Designs_List[[l]][[3]][,1] <- c(paste("Row",Designs_List[[l]][[3]][-length(Designs_List[[l]][[3]][,1]),1], sep = " "),"Columns:")
    }
    #Extract the checks
    checks_df <- data.frame()
    for (i in 1:max(Designs_List[[l]][[1]]$line.code)) {
      checks_df[i,1] <- sum(Designs_List[[l]][[1]]$line.code == i)
      checks_df[i,2] <- unique(Designs_List[[l]][[1]]$line.name[which(Designs_List[[l]][[1]]$line.code == i)])
    }
    
    
    #Extract names for the design data.frame in the form of a vector
    names_vec <- character()
    #Extract the names and capitalize the first letter
    for (i in 1:length(names(Designs_List[[l]][[1]]))) {
      names_vec[i] <- paste(toupper(substring(text = names(Designs_List[[l]][[1]])[i], first = 1, last = 1)), substring(text = names(Designs_List[[l]][[1]])[i], first = 2), sep = "")
      names_vec[i] <- sub(pattern = "_", replacement = ".", x = names_vec[i]) #Replace underscores with periods
    }
    
    #Vector containing all of the column names
    design_dfnames <- c("environment", "trial", "plot.id", "plot", "rep", "row", "column", "blk", "row.blk", "col.blk", "line.code", "line.name", "entry", "duplicates")
    
    #If the SheetNames vector is empty, create the sheet and name it according to the environment
    if (is.null(SheetNames)) {
      sheetlist[[l]] <- createSheet(wb = pdwb, sheetName = paste(unique(Designs_List[[l]][[1]]$environment), substring(text = names(Designs_List[[l]])[1], first = 1, last = gregexpr(pattern = ".dgn", text = names(Designs_List[[l]])[1])[[1]][1]-1), "rand", sep = "_"))
    #If not, then use the names given by the user
    } else {
      sheetlist[[l]] <- createSheet(wb = pdwb, sheetName = SheetNames[l])
    }
    
    
    ##### Workbook titles and subtitles #####
    #Add pertinent information that is located at the topleft section of the final plan design workbook
      #Create the baseline cellstyle
      CS <- CellStyle(wb = pdwb)
    
      
      #Title Cellstyle
      CS <- CS + Font(wb = pdwb, color = "#C65911", heightInPoints = 26, name = "Times New Roman", isBold = TRUE, underline = 1)
      #Add the title
      rows <- createRow(sheet = sheetlist[[l]], rowIndex = 1)
      pf_cell <- createCell(row = rows, colIndex = 1)
      setCellValue(pf_cell[[1,1]], value = SheetTitle[l])
      setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
    
    
    
      #Add subtitles
        #First create a vector with all of the subtitle variables/arguments
        Subtitles <- c(NumLocs, DgnType, TotalPlots, FieldDim, NumDupl)
        SubtitlesVec <- ((1:sum(Subtitles))+1)
    
        #Create the cellstyle for the first two subtitles
        CS <- CS + Font(wb = pdwb, color = "#0070C0", heightInPoints = 14, name = "Times New Roman", isBold = TRUE)
    
    
        #Check if the subtitle is wanted
        if (NumLocs == TRUE) {
        #Number of locations
          rows <- createRow(sheet = sheetlist[[l]], rowIndex = 2)
          pf_cell <- createCell(row = rows, colIndex = 2)
          setCellValue(pf_cell[[1,1]], value = paste(length(Designs_List), "Locations:", paste(names(Designs_List), collapse = ", "), sep = " "))
          setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
        }
        
        #Same cell style
        #Check if the subtitle is wanted
        if (DgnType == TRUE) {
          #Type of design (ie aibd, madii, rcbd)
          rows <- createRow(sheet = sheetlist[[l]], rowIndex = 3 + (sum(Subtitles[1]) - 1))#This needs to be changed
          pf_cell <- createCell(row = rows, colIndex = 2)
          setCellValue(pf_cell[[1,1]], value = paste("Design:", toupper(substring(text = names(Designs_List[[l]])[1], first = 1, last = gregexpr(pattern = ".dgn", text = names(Designs_List[[l]])[1])[[1]][1]-1)), sep = " "))
          setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
        }
        
        #Create the cellstyle for the next few subtitles
        CS <- CS + Font(wb = pdwb, heightInPoints = 14, name = "Times New Roman", isBold = TRUE)#General specs, like plot dimension, number of entries, and number of duplicates
        
        #Check if subtitle is wanted
        if (TotalPlots == TRUE) {
          #Number of test plots
          rows <- createRow(sheet = sheetlist[[l]], rowIndex = 3 + (sum(Subtitles[1:2]) - 1))#This needs to be changed
          pf_cell <- createCell(row = rows, colIndex = 2)
          setCellValue(pf_cell[[1,1]], value = paste(dim(Designs_List[[l]][[1]])[1], "Total test plots", sep = " "))
          setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
        }
        
        #Same cellstyle
        #Check if subtitle is wanted
        if (FieldDim == TRUE) {
          #Dimensions of plot
          rows <- createRow(sheet = sheetlist[[l]], rowIndex = 3 + (sum(Subtitles[1:3]) - 1))#This needs to be changed
          pf_cell <- createCell(row = rows, colIndex = 2)
          setCellValue(pf_cell[[1,1]], value = paste(max(Designs_List[[l]][[1]]$row), "rows x", max(Designs_List[[l]][[1]]$column), "columns", sep = " "))
          setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
        }
        
        #Check if the subtitle is wanted
        if (NumDupl == TRUE) {
        #If statement to check if the duplicates column is actually present (it potentially could not be present)
          if (!is.null(x = Designs_List[[l]][[1]]$duplicates)) {
            #If the duplicates column is present, a number of duplicates cell group will be added
            #Number of duplicates
            rows <- createRow(sheet = sheetlist[[l]], rowIndex = 3 + (sum(Subtitles[1:4]) - 1)) #This needs to be changed
            pf_cell <- createCell(row = rows, colIndex = 2)
            setCellValue(pf_cell[[1,1]], value = paste(sum(Designs_List[[l]][[1]]$duplicates == "D"), "duplicate entries", sep = " "))
            setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
          }
        }
    
        #This subtitle is always included
        #Cell specifying Number of checks
        rows <- createRow(sheet = sheetlist[[l]], rowIndex = 3 + (sum(Subtitles[1:5]) - 1))#This needs to be changed
        pf_cell <- createCell(row = rows, colIndex = 2)
        setCellValue(pf_cell[[1,1]], value = paste(sum(Designs_List[[l]][[1]]$line.code != 0), "check plots:", sep = " "))
        setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS + Font(wb = pdwb, heightInPoints = 14, name = "Times New Roman", isItalic = TRUE, isBold = TRUE))
        
        
    #Create separately formatted cells for the checks
    #Create the cellstyle
    CS <- CS +
        Font(wb = pdwb, heightInPoints = 12, name = "Times New Roman", isBold = TRUE)
        
    checks_df[,2] <- as.vector(checks_df[,2]) #Designate the names of this list to be the same as the checks
    
    #If CheckCol is not user-defined, set it as the first n colors of rainbow
    if (is.null(CheckCol)) {
      CheckCol <- rainbow(n = dim(checks_df)[1])
    }
    
    for (i in 1:dim(checks_df)[1]) { #For loop to iterate through the number of checks
      rows <- createRow(sheet = sheetlist[[l]], rowIndex = (sum(Subtitles) + 2) + i) #This 7 needs to be substituted with the number of rows requested in arguments
      checksumcell <- createCell(row = rows, colIndex = 3)
      checknamecell <- createCell(row = rows, colIndex = 4)
      setCellValue(checksumcell[[1,1]], value = sum(Designs_List[[l]][[1]]$line.code == i))
      setCellValue(checknamecell[[1,1]], value = unique(Designs_List[[l]][[1]]$line.name[which(Designs_List[[l]][[1]]$line.code == i)]))
      setCellStyle(cell = checksumcell[[1,1]], cellStyle = CS + Fill(foregroundColor = CheckCol[i], backgroundColor = "white"))
      setCellStyle(cell = checknamecell[[1,1]], cellStyle = CS + Fill(foregroundColor = CheckCol[i], backgroundColor = "white"))
    }
    
    
    
    ##### Adding design list, plot and check layouts #####
    #Write the design dataset (first item in the output list) to the sheet
    #Create the cellstyle
      CS <- CS + #Column names for the design data.frame
        Font(wb = pdwb, isBold = TRUE) + 
        Border(position = c("TOP", "BOTTOM", "LEFT", "RIGHT")) + 
        Alignment(horizontal = "ALIGN_CENTER", vertical = "VERTICAL_CENTER") #Alignment of cell text; type HALIGN_STYLES_ for other character values
    
    #Add data
    #Create the data.frame
    #This matches the names that are called by the numeric vector DgnNames to the names present in the design data.frame that will be added to the worksheet. If the names do not match, they will not be added (note that rep does not match because the RCBD design output has "replication" instead)
    design_tempdf <- Designs_List[[l]][[1]][,match(x = design_dfnames[DgnNames], table = names(Designs_List[[l]][[1]]))[which(!is.na(x = match(x = design_dfnames[DgnNames], table = names(Designs_List[[l]][[1]]))))]]
    #Add names to the data.frame
    names(design_tempdf)[1:dim(design_tempdf)[2]] <- names_vec[match(x = names(design_tempdf)[1:dim(design_tempdf)[2]], table = names(Designs_List[[l]][[1]]))]
    
    #If the Source argument is T, then add the source column
    if (Source == TRUE) {
      #Add empty data as filler for the "source" column
      design_tempdf[,(dim(design_tempdf)[2]+1)] <- rep("", dim(design_tempdf)[1])
      #Add the "Source" column name
      names(design_tempdf)[dim(design_tempdf)[2]] <- "Source"
      
      #Rearrange the data.frame so that the "source" column follows the "line.code" column or is in the 5th column position
      if(11 %in% DgnNames) {
        #Add the column number of the last column (and therefore the source column) to the position following the number referring to the "line.code" column
        design_tempdf <- design_tempdf[,append(x = 1:(length(design_tempdf)-1), values = dim(design_tempdf)[2], after = which(DgnNames == 11))]
        #If "line.code" is not present, then the "source" column is defaulted to the 5th column position or last position
      } else {
        if(dim(design_tempdf)[2] > 4) {
          design_tempdf <- design_tempdf[,append(x = 1:(length(design_tempdf)-1), values = dim(design_tempdf)[2], after = 4)]
        }
      }
    }
    
    #Add the dgn data.frame
    addDataFrame(x = design_tempdf, sheet = sheetlist[[l]], col.names = TRUE, row.names = FALSE, startRow = (sum(Subtitles) + 2) + dim(checks_df)[1] + 2, colnamesStyle = CS)  #This 7 needs to be substituted with the number of rows requested in arguments
    
    #Write the plot layout data to the worksheet
    addDataFrame(x = as.data.frame(Designs_List[[l]][[3]]), sheet = sheetlist[[l]], col.names = FALSE, row.names = FALSE, startRow = 2, startColumn = dim(design_tempdf)[2] + 4)
    
    #Write the check layout data to the worksheet
    addDataFrame(x = as.data.frame(Designs_List[[l]][[4]]), sheet = sheetlist[[l]], col.names = FALSE, row.names = FALSE, startRow = (1 + dim(Designs_List[[l]][[3]])[1] + 4), startColumn = dim(design_tempdf)[2] + 4, colStyle = CellStyle(wb = pdwb, dataFormat = DataFormat(x = "0.0")))
    
    
    
    ##### Borders, highlighting, and other formatting #####
    
    #Only perform highlighting if the line.name column is present
    if (!is.na(match(12, DgnNames))) {
    
      #Highlighting the cells in the design data.frame that are checks
      #First need to redesignate the cellstyle for checks to eliminate boldness
      CS <- CellStyle(wb = pdwb)
      
      #Next, iterate through all rows of the design data.frame, check to see if the cell value is a check, then assign the appropriate cell style
      for (n in 1:dim(Designs_List[[l]][[1]])[1]) { #For loop to iterate through every row in the design data.frame
        rows <- getRows(sheet = sheetlist[[l]], rowIndex = n + (sum(Subtitles) + 2) + dim(checks_df)[1] + 2) # 7 = number of rows taken up by pertinant primary info, dim(checks_df) = rows taken up by checks, 2 = space between checks and the data.frame  #This 7 needs to be substituted with the number of rows requested in arguments
        #If line.code, line.name, and plot are not together, then only highlight line.name
        if (!is.na(match(NA, match(c(4,11,12), DgnNames)) > 0) || max(match(c(4,11,12), DgnNames)) - min(match(c(4,11,12), DgnNames)) > 2) {
          colin <- match(12, DgnNames)
          pf_cell <- getCells(row = rows, colIndex = sort(colin))
          pf_value <- getCellValue(cell = pf_cell[[1]])
          
          #Highlight just line.name
          if (sum(pf_value == checks_df[,2]) == 1) { #If statement to ask whether the entry names matches any of the checks
            setCellStyle(cell = pf_cell[[1]], cellStyle = CS + Fill(foregroundColor = CheckCol[which(pf_value == checks_df[,2])], backgroundColor = "white")) #If the value of the cell matches any of the names of csCheckdflist (and therefore any of the checks), the cellstyle will be set to the style designated for that check
          }
        } else {
          #If source is not included, only highlight line.code, line.name, and plot
          if (Source == FALSE) {
            colin <- match(c(4,11,12), DgnNames)
            pf_cell <- getCells(row = rows, colIndex = sort(colin)) 
            pf_value <- getCellValue(cell = pf_cell[[2]])
            
            if (sum(pf_value == checks_df[,2]) == 1) { #If statement to ask whether the entry names matches any of the checks
              setCellStyle(cell = pf_cell[[1]], cellStyle = CS + Fill(foregroundColor = CheckCol[which(pf_value == checks_df[,2])], backgroundColor = "white")) #If the value of the cell matches any of the names of csCheckdflist (and therefore any of the checks), the cellstyle will be set to the style designated for that check
              setCellStyle(cell = pf_cell[[2]], cellStyle = CS + Fill(foregroundColor = CheckCol[which(pf_value == checks_df[,2])], backgroundColor = "white"))
              setCellStyle(cell = pf_cell[[3]], cellStyle = CS + Fill(foregroundColor = CheckCol[which(pf_value == checks_df[,2])], backgroundColor = "white"))
            }
          } else {
            #If Source is true and the three parameters are together, then highlight all four
            colin <- append(x = match(c(4,11,12), DgnNames), values = max(match(c(4,11,12), DgnNames)) + 1)
            pf_cell <- getCells(row = rows, colIndex = sort(colin)) 
            pf_value <- getCellValue(cell = pf_cell[[2]])
            
            if (sum(pf_value == checks_df[,2]) == 1) { #If statement to ask whether the entry names matches any of the checks
              setCellStyle(cell = pf_cell[[1]], cellStyle = CS + Fill(foregroundColor = CheckCol[which(pf_value == checks_df[,2])], backgroundColor = "white")) #If the value of the cell matches any of the names of csCheckdflist (and therefore any of the checks), the cellstyle will be set to the style designated for that check
              setCellStyle(cell = pf_cell[[2]], cellStyle = CS + Fill(foregroundColor = CheckCol[which(pf_value == checks_df[,2])], backgroundColor = "white"))
              setCellStyle(cell = pf_cell[[3]], cellStyle = CS + Fill(foregroundColor = CheckCol[which(pf_value == checks_df[,2])], backgroundColor = "white"))
              setCellStyle(cell = pf_cell[[4]], cellStyle = CS + Fill(foregroundColor = CheckCol[which(pf_value == checks_df[,2])], backgroundColor = "white"))
            }
          }
        }          
      }
      
      #Highlighting check positions in the ploy.layout data.frame
      for (r in 1:(dim(Designs_List[[l]][[3]])[1]-1)) { #For loop to cycle through the number of rows
        for (k in 1:(dim(Designs_List[[l]][[3]])[2]-1)) { #For loop to cycle through the number of columns
          rows <- getRows(sheet = sheetlist[[l]], rowIndex = r + 1) #
          pf_cell <- getCells(row = rows, colIndex = (k + dim(design_tempdf)[2] + 4)) #MAY NEED TO CHANGE COLUMN INDEX FOR SOFT-CODING
          pf_value <- getCellValue(cell = pf_cell[[1]])
          
          if (sum(Designs_List[[l]][[1]]$line.name[which(Designs_List[[l]][[1]]$plot == pf_value)] == checks_df[,2]) == 1) {
            setCellStyle(cell = pf_cell[[1]], cellStyle = CS + Fill(foregroundColor = CheckCol[which(Designs_List[[l]][[1]]$line.name[which(Designs_List[[l]][[1]]$plot == pf_value)] == checks_df[,2])], backgroundColor = "white"))  
          } #Close the if statement
        }} #Close the for loops
      
      #Highlighting check positions in the check.layout data.frame
      for (r in 1:(dim(Designs_List[[l]][[4]])[1]-1)) { #For loop to cycle through the number of rows
        for (k in 1:(dim(Designs_List[[l]][[4]])[2]-1)) { #For loop to cycle through the number of columns
          rows <- getRows(sheet = sheetlist[[l]], rowIndex = (r + dim(Designs_List[[l]][[3]])[1] + 4)) #
          pf_cell <- getCells(row = rows, colIndex = (k + dim(design_tempdf)[2] + 4)) #MAY NEED TO CHANGE COLUMN INDEX FOR SOFT-CODING
          pf_value <- getCellValue(cell = pf_cell[[1]])
          
          if (pf_value != 0) { #If statement to proceed on the condition that the value is not 0 (and is therefore a check)
            setCellStyle(cell = pf_cell[[1]], cellStyle = CS + Fill(foregroundColor = CheckCol[as.numeric(pf_value)], backgroundColor = "white"))
          } #Close the if statement
        }} #Close the for loops
    } #Close the initial if statement
    
    
    #Manipulating the blocks created in the design
    #First determine if blocks were requested
    if (BlkBorders == TRUE) {
      
      #If statement to check what design format is being used
      if (substring(text = names(Designs_List[[l]])[1], first = 1, last = gregexpr(pattern = ".dgn", text = names(Designs_List[[l]])[1])[[1]][1]-1) == "aibd" || substring(text = names(Designs_List[[l]])[1], first = 1, last = gregexpr(pattern = ".dgn", text = names(Designs_List[[l]])[1])[[1]][1]-1) == "mad") {
      
        #Create numeric vector containing the block dimensions  
        blk_dim <- numeric()
        blk_dim[1] <- max(Designs_List[[l]][[1]]$row)/max(Designs_List[[l]][[1]][,9]) #Add the number of rows in each block
        blk_dim[2] <- max(Designs_List[[l]][[1]]$column)/max(Designs_List[[l]][[1]][,10]) #Add the column dimension (this is done by finding various maxima and minima in the dgn data.frame)
        
        for (r in 1:max(Designs_List[[l]][[1]][,9])) { #For loop to cycle through the number of row_blk
          
          for (k in 1:max(Designs_List[[l]][[1]][,10])) { #For loop to cycle through the number of col_blk
            
            #Manipulate the border of the blocks by creating cell blocks of the cells
            #Create a cell block that includes the cells in block 9 (for example)
            cellblk1 <- CellBlock(sheet = sheetlist[[l]], #Designate the sheet
                                  startRow = (2 + ((r-1) * blk_dim[1])), #The plot layout will usually start in row 2
                                  startColumn = ((dim(design_tempdf)[2] + 4 + 1) + ((k-1) * blk_dim[2])), #Start column is the same as the start column that was used for placing the plot.layout data + 1 to exclude the row names
                                  noRows = blk_dim[1],
                                  noColumns = blk_dim[2],
                                  create = FALSE)
            
            #Set borders, first the top left
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("TOP", "LEFT"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = 1)
            #Top right
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("TOP", "RIGHT"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = blk_dim[2])
            #Bottom right
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("BOTTOM", "RIGHT"), pen = "BORDER_MEDIUM"), rowIndex = blk_dim[1], colIndex = blk_dim[2])
            #Bottom left
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("BOTTOM", "LEFT"), pen = "BORDER_MEDIUM"), rowIndex = blk_dim[1], colIndex = 1)
            #Top row
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("TOP"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = 2:(blk_dim[2]-1)) #Columns are added by starting at the second column (2), then including all columns up until the last one (blk_dim[2]-1)
            #Right side
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("RIGHT"), pen = "BORDER_MEDIUM"), rowIndex = 2:(blk_dim[1]-1), colIndex = blk_dim[2]) #Rows are added by starting at the second row (2), then including all rows up until the last one (blk_dim[1]-1)
            #Bottom row
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("BOTTOM"), pen = "BORDER_MEDIUM"), rowIndex = blk_dim[1], colIndex = 2:(blk_dim[2]-1))
            #Left side
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("LEFT"), pen = "BORDER_MEDIUM"), rowIndex = 2:(blk_dim[1]-1), colIndex = 1)
          }}  #Close the for loops
        
      } else { #Close the previous if statement and open and else + if statement
        
          #Border formatting for RCBD designs
          if (substring(text = names(Designs_List[[l]])[1], first = 1, last = gregexpr(pattern = ".dgn", text = names(Designs_List[[l]])[1])[[1]][1]-1) == "rcbd") {
            
            #Create a border around the whole plot map first to simplify things later
            #First create a cellblock containing all of the cells
            cellblk1 <- CellBlock(sheet = sheetlist[[l]], #Designate the sheet
                                  startRow = 2, #The plot layout will usually start in row 2 
                                  startColumn = ((dim(design_tempdf)[2] + 4 + 1)), #Start column is the same as the start column that was used for placing the plot.layout data + 1 to exclude the row names
                                  noRows = (dim(Designs_List[[l]][[3]]) - 1)[1], #Number of rows is the number of rows in the field design
                                  noColumns = (dim(Designs_List[[l]][[3]]) - 1)[2], #Number of cols is the number of cols in the field design
                                  create = FALSE)
            
            #Add the perimeter border
            #Set borders, first the top left
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("TOP", "LEFT"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = 1)
            #Top right
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("TOP", "RIGHT"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = (dim(Designs_List[[l]][[3]]) - 1)[2])
            #Bottom right
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("BOTTOM", "RIGHT"), pen = "BORDER_MEDIUM"), rowIndex = (dim(Designs_List[[l]][[3]]) - 1)[1], colIndex = (dim(Designs_List[[l]][[3]]) - 1)[2])
            #Bottom left
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("BOTTOM", "LEFT"), pen = "BORDER_MEDIUM"), rowIndex = (dim(Designs_List[[l]][[3]]) - 1)[1], colIndex = 1)
            #Top row
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("TOP"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = 2:((dim(Designs_List[[l]][[3]]) - 1)[2]-1)) #Columns are added by starting at the second column (2), then including all columns up until the last one (blk_dim[2]-1)
            #Right side
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("RIGHT"), pen = "BORDER_MEDIUM"), rowIndex = 2:((dim(Designs_List[[l]][[3]]) - 1)[1]-1), colIndex = (dim(Designs_List[[l]][[3]]) - 1)[2]) #Rows are added by starting at the second row (2), then including all rows up until the last one (blk_dim[1]-1)
            #Bottom row
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("BOTTOM"), pen = "BORDER_MEDIUM"), rowIndex = (dim(Designs_List[[l]][[3]]) - 1)[1], colIndex = 2:((dim(Designs_List[[l]][[3]]) - 1)[2]-1))
            #Left side
            CB.setBorder(cellBlock = cellblk1, border = Border(position = c("LEFT"), pen = "BORDER_MEDIUM"), rowIndex = 2:((dim(Designs_List[[l]][[3]]) - 1)[1]-1), colIndex = 1)
          
            
            #Iterating through all of the plots to draw inside borders (RCBDs can have funky borders for the replicates)
            #For loop through each row
            for (r in 1:max(Designs_List[[l]][[1]]$row)) { #For loop to cycle through the number of rows
              for (k in 1:max(Designs_List[[l]][[1]]$column)) { #For loop to cycle through the number of cols
                
                #Create a cellblock with just one cell
                cellblk1 <- CellBlock(sheet = sheetlist[[l]], #Designate the sheet
                                      startRow = (r + 1), #The plot layout will usually start in row 2 
                                      startColumn = (k + dim(design_tempdf)[2] + 4), #Start column is the same as the start column that was used for placing the plot.layout data + 1 to exclude the row names
                                      noRows = 1, #Number of rows is the number of rows in the field design
                                      noColumns = 1, #Number of cols is the number of cols in the field design
                                      create = FALSE)
                
                #Check if the cell is the last in the column
                if (k != max(Designs_List[[l]][[1]]$column)) {
              
                  #Check if the cell to the right of the current cell is the same rep
                  if (Designs_List[[l]][[5]][r,(k+1)] != Designs_List[[l]][[5]][r,(k+1)+1]) {
                    
                    #Add a right-side border to the cellstyle
                    CB.setBorder(cellBlock = cellblk1, border = Border(position = c("RIGHT"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = 1)
                  }} #Close the if statements
                
                #Check if the cell is the last in the row
                if (r != max(Designs_List[[l]][[1]]$row)) {
                  
                  #Check if the cell to the bottom of the current cell is the same rep
                  if (Designs_List[[l]][[5]][r,(k+1)] != Designs_List[[l]][[5]][r+1,(k+1)]) {
                    
                    #Add a bottom-side border to the cellstyle
                    CB.setBorder(cellBlock = cellblk1, border = Border(position = c("BOTTOM"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = 1)
                  }} #Close the if statements

              }} #Column loop and Row loop
            
          } #If statement for RCBD
      } #Close the else statement
    } #Close the BlkBorders if statement



    ##### Post-Formatting #####
    #Autosize columns
    autoSizeColumn(sheet = sheetlist[[l]], colIndex = c(3:(dim(Designs_List[[l]][[4]])[2] + dim(design_tempdf)[2] + 4)))
    
    #Set the zoom
    setZoom(sheet = sheetlist[[l]], numerator = 85, denominator = 100)
    
    #Manual column widths
    setColumnWidth(sheet = sheetlist[[l]], colIndex = 1, colWidth = 11)
      
    
  
  } #Close the per-design list for loop



  ##### Closing time #####
  #Save the workbook to the disk
  saveWorkbook(wb = pdwb, file = filename)
  
  
} #Close the function
  

