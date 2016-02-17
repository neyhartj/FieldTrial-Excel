#' Generate a randomized complete block design
#'
#' @description Randomized complete block designs (RCBD) are exceedingly common experimental designs in many field, including agricultural reserach. Details on RCBDs can easily be found in basic experimental design textbooks or an internet search; briefly, a set of experimental entries along with check lines, if desired, comprise a single block. This block is then replicated in its entirety at least once (i.e. two blocks) and up to any number of times.
#'              Blocks can be replicated within an environment, across environments, or both. Under the common circumstance of limited resources, Bernardo (2010; p.184) suggest prioritizing replication across environments rather than within environments. More is explained in \code{Details}.
#' @param enviro Optional descriptor of the environment which the field trial will be grown in. Takes single-string alpha-numeric arguments, e.g. "MN15", "MN_15", "DisneyWorld", etc...). Default is the current date and time.
#' @param exp.name Required input of an identifying name for the experiment. For example, if the trial was Gregor Mendel's yield trial 1 the \code{exp.name} may be "GM_YT1". Accepts alpha-numeric arguments.
#' @param nBlks Optional \code{Integer} indicating the number of complete blocks desired. Default is \code{2}. If a single replication is desired to grow across multiple environments (aka locations, see \code{Details} for further information) then set nBlks to \code{1}.
#' @param entries A \code{vector} of experimental (i.e. not check) entries. Accepts alpha-numeric entry names. Default is \code{NULL}, meaning this \strong{OR} \code{nEntries} (see next) requires an input.
#' @param nEntries \code{Integer} argument indicating how many experimental entries are to be included in the trial. If \code{nEntries} = \emph{m}, then \emph{m} generic entry names will be generated and included in the output. Default is \code{NULL}, meaning this \strong{OR} \code{nEntries} requires an input. \code{nEntries} will be superceded if \code{entries} is provided.
#' @param chks A \code{vector} including the set of check lines to be included. Accepts alpha-numeric check names. The  Default is \code{NULL}, meaning this \strong{OR} \code{nChks} requires an input.
#' @param nChks \code{Integer} argument indicating how many check lines (see \code{Details}) are to be included in the trial. Default is NULL, meaning this \strong{OR} \code{chks} requires an input.
#' @param nChkReps Optional \code{Integer} argument indicating how many times each check line should be replicated \strong{per block}. Default is \code{1} meaning that each check line will appear in each block twice (i.e. replicated once). 
#' @param nFieldRows \code{Integer} argument indicating the number of field rows (dimension 'a' in Figure 1). Since researchers typically know the depth of their field and the lenght of their field plots, this is the only required field dimension input.
#' @param nFieldCols Optional \code{integer} argument indicating the number of field columns (dimension 'b',Figure 1). If this is known and will allow enough field plots to accomodate the desired number of entries and checks, it can be supplied. If not provided the default is \code{NULL} and the necessary dimension will be calculated within the function.
#' @param plot.start Optional \code{integer} indicating the identification number of the first plot. \code{design.aibd} places the first plot in the "bottom-left" corner of the field and numbering then progresses from left-to-right, front-to-back (Figure 1). Default is \code{1001}.
#' @param fillWithEntry Optional \code{Logical}. \code{TRUE} by default. When \code{TRUE}, fill plots (i.e. plots remaining after the defined sets of experimental entries and check lines are assigned) will be replaced by randomly-selected experimental entries. A "D" will appear in the output next to these replicated entries.
#' @param dup.list Optional \code{character vector} containing the candidates that could be replicated if \code{fillWithEntry = TRUE}. This is useful if there is a low amount of source seed for some experimental entries; those, for example, would be excluded from \code{dup.list}. Default is \code{NULL}, meaning that all entries are candidates for duplication.
#' @param fillWithChk Optional \code{Logical}. If \code{TRUE} this will supercede \code{fillWithEntry} and any fill plots will be replaced by a balanced set of check lines. Default is \code{FALSE}.
#' @parem output.excel Optional \code{Logical}. If \code{TRUE} an excel worksheet that visualizes the design - including color-coded checks and a map of the field - is written to the current directory.
#' @details Figure 1 presents a simplified version of an RCBD with two replications within an environment where each block is a rectangle and there are no fill plots. In reality a block may end and the next start in the middle of a row; this is done to minimize the size of the experiment while meeting all requirements.
#'          
#'          As mentioned, for a given set of resources (i.e. field plots planted), planting a single replication across more environents will provide a more precision than planting multiple replications within an environent \cite{(Bernardo, 2010; p.184)}. This is based off of the calculation of least significant difference (LSD) between the means of two genotypes, which is presented in Bernardo (2010; eq. 8.5) as:
#'          
#'          \if{html}{\figure{LSD_eq_Bernardo_8.5_rcbd.jpg}}
#'          
#'          It can readily be seen that \emph{e} (numbr of environments, aka locations) appears twice in denominators of the expression below the radial whereas \emph{r} (number of reps per environment) appears in just one denominator. Logically, this fact results in a greater reduction in \eqn{LSD} for every additional environment than with every additional replication within an environment.
#'          Another way of saying it would be that to attain a set level of precision, fewer plots would be required if the experiment were grown in more environments rather than in replications within environments. Granted, it is often more practical to to plant replications within environments rather than across many environments, which is of course why both scenarios are facilitated by \code{design.rcbd}.
#'          
#'          
#'          
#'          
#'          \if{html}{\figure{rcbd_Fig1.jpg}}
#'  RCBD may not be well-suited for all circumstances; in these cases consider using other design functions included in \link{FldTrial}: \link{design.aibd} or \link{design.mad}.
#'  @return A \code{list} containing:
#'          \itemize{
#'            \item \code{rcbd.dgn} A \code{dataframe} of the resulting design. A line code of 0 denotes an experimental entry, codes of 1, 2, ..., (nChks) denote the first, second, up to the (nChks)-th check, respectively. The order of the check lines and thus line code is arbitrary and has no bearing on placement or randomization. 
#'                                  If \code{fillWithEntry} is \code{TRUE} then a  "D" will be placed in the "Duplicates" column for all duplicated entries.
#'                                  Note that the column name "replication" is used in \code{rcbd.dgn} rather than "block" or "blk" for reasons related to compatibility with a preferred data entry software.
#'            \item \code{field.book} An abbreviated version of \code{rcbd.dgn} for use with eletronic data entry systems (i.e. tablets).
#'            \item \code{plot.layout} A \code{matrix} of plot numbers reflecting the field layout. The first column contains field row numberis (i.e. 1:\code{nFieldRows}) and the last row contains field column numbers (i.e. 1:\code{nFieldCols}). See Figure 1.
#'            \item \code{check.layout} A \code{matrix} of line codes reflecting the field layout. See \code{plot.layout} for further details.
#'            \item \code{rep.layout} A \code{matirx} of block assignments reflecting the field layout. See \code{plot.layout} for further details.
#'            \item \code{field.dim} The dimensions of the resulting field design; nFieldRows x nFieldCols, or a x b (Figure 1).
#'            \item \code{nDuplicates} The number of duplicated experimental entries. Will only be non-zero if \code{fillWithEntry} is \code{TRUE}.
#'            \item \code{rlzPerChks} The final percentage of the designed represented by check lines. This can be affected by nChkReps and fillWithCheck
#'          } 
#' @references
#'        Bernardo, Rex. 2010. Breeding for Quantitative Traits in Plants. Stemma Press. Woodbury, MN.
#' 
#' @examples
#' \dontrun{
#' ## Example 1 - Generic layout with 2 reps
#' rcbd.ex1 <- design.rcbd(exp.name = "ex1", nBlks = 2, 
#'                nEntries = 80, nChks = 3, nChkReps = 3, 
#'                nFieldRows = 8)
#'                
#' ## Example 2 - Layout with defined lines and checks
#' rcbd.ex2 <- design.rcbd(exp.name = "ex2", nBlks = 2, 
#'                entries = paste("Line", 1:75, sep="."), 
#'                chks = c("Larry", "Curly", "Moe"), 
#'                nChkReps = 4, nFieldRows = 9)
#' } 
#' @export


## Should there be a rect.blocks options? So each block is its own rectangle, i.e. transition isn't in the middle of a row
design.rcbd <- function(enviro=format(Sys.Date(), "%x"), exp.name=NULL, nBlks=2, entries= NULL, nEntries= NULL, chks=NULL, nChks=NULL, nChkReps=1, nFieldRows=NULL, nFieldCols=NULL, plot.start=1001, fillWithEntry=T, dup.list=NULL, fillWithChk=F, output.excel = F) {
  
  ## QC
  if(is.null(entries) & is.null(nEntries)) stop("Must provide an entry list (entries=) OR the number of entries desired (nEntries=).")
  
  if(is.null(chks) & is.null(nChks)) stop("Must provide a list of check names (chk.names=) with the primary check listed first\n OR the number of SECONDARY checks desired (nSecChk=).")
  
  if(is.null(nFieldRows)) stop("Provide the number of rows (sometimes called ranges or beds) (nFieldRows=)")
  
  ## Develop other non-input functions parameters
  if(is.null(exp.name)) stop("Must provide a name for the experiment (exp.name=)"); trial <- paste(exp.name, enviro, sep="_")
  if(!is.null(entries)){entries <- as.character(as.matrix(entries)); nEntries <- length(entries)} else entries <- paste("entry", 1:nEntries, sep="_")
  if(!is.null(chks)) chks <- as.character(as.matrix(chks)) else chks <- paste("chk", 1:nChks, sep="_")
  if(!is.null(dup.list)) dup.list <- as.character(as.matrix(dup.list)) else{dup.list <- entries}
  
  nChks <- length(chks)
  
  entry.mat <- cbind(line.name=c(unique(entries), chks, "FILLER"), entry=c(1:(length(unique(entries)) + nChks), NA)) ## This is for field book
  
  if(all(fillWithEntry, fillWithChk)) stop("Select either fillWithEntry or fillWithCheck")
  if(fillWithEntry) fillWithChk <- FALSE
  if(fillWithChk) fillWithEntry <- FALSE
  
  nTotal <- nBlks*(nEntries + nChks*nChkReps)
  
  if(!is.null(nFieldCols)){
    if(nFieldRows * nFieldCols < nTotal) stop("The field dimensions provided are not large enough to accomodate all of the defined entries and checks.")
  }
  
  if(is.null(nFieldCols)) nFieldCols <- ceiling(nTotal / nFieldRows)
  
  exp.size <- nFieldCols * nFieldRows ## This includes all blks (i.e. reps)
  
  nFill.exp <- exp.size - nTotal
  nFill.blk <- floor(nFill.exp/nBlks)
  nFill.remainder <- nFill.exp - (nFill.blk*nBlks)
  
  ## Build field design
  fld.dgn <- data.frame(environment=enviro, trial=trial, plot.id=paste(trial, plot.start:(plot.start+(exp.size-1)), sep="_"), plot=plot.start:(plot.start+(exp.size-1)), replication=c(rep(1:nBlks, each=(exp.size/nBlks)), rep("FILLER", times=nFill.remainder)), row=rep(1:nFieldRows, each=nFieldCols), column=rep(c(1:nFieldCols, nFieldCols:1), length.out=exp.size), line.code=rep(NA, times=exp.size), line.name=rep(NA, times=exp.size))
    
  ## Fills in FILL plots, and randomizes chks, entries, etc..
  if(fillWithChk) entries.col <- unlist(lapply(1:nBlks, function(b){tmp <- c(entries, rep(chks, times=nChkReps), rep(chks[sample(1:nChks, size = nChks, replace = FALSE)], times=nFill.blk/nChks, length.out=nFill.blk)); tmp[sample(1:length(tmp), replace = FALSE)]}))
  if(fillWithEntry) entries.col <- unlist(lapply(1:nBlks, function(b){tmp <- c(entries, rep(chks, times=nChkReps), rep(dup.list[sample(1:length(dup.list), size = length(dup.list), replace = FALSE)], times= ceiling(nFill.blk/length(dup.list)), length.out=nFill.blk)); tmp[sample(1:length(tmp), replace = FALSE)]}))
  if(!fillWithEntry & !fillWithChk) entries.col <- unlist(lapply(1:nBlks, function(b){tmp <- c(entries, rep(chks, times=nChkReps)); tmp <- c(tmp[sample(1:length(tmp), replace = FALSE)], rep("FILLER", times=nFill.blk))}))
  
  fld.dgn$line.name <- c(entries.col, rep("FILLER", times=nFill.remainder))
  fld.dgn$line.code <- sapply(entries.col, function(X){if(X %in% entries) return(0) ; if(X %in% chks) return(which(chks == X)); if(X == "FILLER") return(NA)})
  
  fld.dgn <- merge(fld.dgn, entry.mat, by = "line.name", all.x = TRUE); fld.dgn <- fld.dgn[order(fld.dgn$plot), ]
  fld.book <- cbind(fld.dgn[,c(3,4,2,9,6,7,10,5)], notes=" ")
  
  if(fillWithEntry){
    fld.dgn$duplicates  <- ""
    for(b in 1:nBlks){
      blk.tmp <- subset(fld.dgn, replication == b)
      dup.entries <- blk.tmp[which(duplicated(blk.tmp$entry) & blk.tmp$line.code %in% 0), "entry"]
      fld.dgn[which(fld.dgn$replication == b & fld.dgn$entry %in% dup.entries), "duplicates"]  <- "D"
    }
  }
  nDups <- length(which(fld.dgn$duplicates == "D"))/2

  per.chks <- round(length(which(fld.dgn[,"line.code"] %in% (1:nChks))) / length(!is.na(fld.dgn[,"line.code"])), digits = 3)
  
  ## Make field.layout with plot numbers
  for.layout.mats <- fld.dgn[order(fld.dgn$row, fld.dgn$column), ]
  
  fld.plot <- apply(t(matrix(for.layout.mats$plot, nrow = nFieldRows, ncol=nFieldCols, byrow=T)), 1, rev)
  fld.plot <- cbind(nFieldRows:1, fld.plot)
  fld.plot <- rbind(fld.plot, c("", 1:nFieldCols))
  
  ## Make field.layout with line codes
  line.code <- apply(t(matrix(for.layout.mats$line.code, nrow = nFieldRows, ncol=nFieldCols, byrow=T)), 1, rev)
  line.code <- cbind(nFieldRows:1, line.code)
  line.code <- rbind(line.code, c("", 1:nFieldCols))
  
  ## Make field.layout with blk codes
  blk.code <- apply(t(matrix(for.layout.mats$replication, nrow = nFieldRows, ncol=nFieldCols, byrow=T)), 1, rev)
  blk.code <- cbind(nFieldRows:1, blk.code)
  blk.code <- rbind(blk.code, c("", 1:nFieldCols))
  
  # Create the output list
  output <- list(rcbd.dgn=fld.dgn, field.book=fld.book, plot.layout=fld.plot, 
                 check.layout=line.code, rep.layout=blk.code, 
                 field.dim=paste(nFieldRows, "x", nFieldCols, sep=" "), nDuplicates=nDups, 
                 rlzPerChks=per.chks)
  
  # Proceed if the output.excel logical is TRUE
  if (output.excel) {
    
    # Notify the user
    cat(paste("Creating and writing a .xlsx file to: "), getwd(), sep = "")
    
    ## This is the version of the design2excel function to be integrated into existing functions
    create.xlsx <- function(fld.dgn, fld.book, fld.plot, line.code, blk.code, design.type = "rcbd") {
      
      # Function package requirements
      require(xlsx, quietly = T)
      
      ##### Workbook setup #####
      #Create the new workbook
      pdwb <- createWorkbook()
      
      # Set some initial variables
      # Worksheet attributes
      # The title of the sheet - set this to the trial name
      sheet.title <- as.character(unique(fld.dgn$trial))
      # Design attributes 
      design.env <- as.character(unique(fld.dgn$environment))
      
      # Some hard coding  
      # Set the column numbers to extract from the ...
      dgn.names <- c(1,4,12,11,6,7,8,9,10,14)
      #Vector containing all of the column names
      design_dfnames <- c("environment", "trial", "plot.id", "plot", "rep", "row", "column", "blk", "row.blk", "col.blk", "line.code", "line.name", "entry", "duplicates")
      
      # Create the main sheet
      # The name of the sheet (i.e. the name of the sheet tab at the bottom of the workbook) is set to
      ## a combination of the environment and the design type
      main.sheet <- createSheet(wb = pdwb, sheetName = paste(design.env, design.type, "rand", sep = "_"))
      
      
      ###### Pre-formatting #####
      
      #Add "Row" to the column containing row numbers for the plot layout data.frame
      #Using an if statement to check to see if "Row" had already been added
      if (length(grep(pattern = "Row", x = fld.plot[,1][1])) == 0) {
        fld.plot[,1] <- c(paste("Row",fld.plot[-length(fld.plot[,1]),1], sep = " "),"Columns:")
      }
      #Extract the checks
      checks_df <- data.frame()
      for (i in 1:max(fld.dgn$line.code, na.rm = T)) {
        checks_df[i,1] <- sum(fld.dgn$line.code == i, na.rm = T)
        checks_df[i,2] <- unique(fld.dgn$line.name[which(fld.dgn$line.code == i)])
      }
      
      #Extract names for the design data.frame in the form of a vector
      names_vec <- character()
      #Extract the names and capitalize the first letter
      for ( colname in names(fld.dgn) ) {
        name.to.add <- paste(toupper(substring(text = colname, first = 1, last = 1)), substring(text = colname, first = 2), sep = "")
        names_vec <- append(x = names_vec, values = sub(pattern = "_", replacement = ".", x = name.to.add))
      }
      
      ##### Workbook titles and subtitles #####
      #Add pertinent information that is located at the topleft section of the final plan design workbook
      #Create the baseline cellstyle
      CS <- CellStyle(wb = pdwb)
      
      
      #Title Cellstyle
      CS <- CS + Font(wb = pdwb, color = "#C65911", heightInPoints = 26, name = "Times New Roman", isBold = TRUE, underline = 1)
      #Add the title
      rows <- createRow(sheet = main.sheet, rowIndex = 1)
      pf_cell <- createCell(row = rows, colIndex = 1)
      setCellValue(pf_cell[[1,1]], value = sheet.title)
      setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
      
      #Create the cellstyle for the first two subtitles
      CS <- CS + Font(wb = pdwb, color = "#0070C0", heightInPoints = 14, name = "Times New Roman", isBold = TRUE)
      
      #Number of locations
      rows <- createRow(sheet = main.sheet, rowIndex = 2)
      pf_cell <- createCell(row = rows, colIndex = 2)
      setCellValue(pf_cell[[1,1]], value = paste("1 Location:", paste(design.env, collapse = ", "), sep = " "))
      setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
      
      #Same cell style
      #Type of design (ie aibd, madii, rcbd)
      rows <- createRow(sheet = main.sheet, rowIndex = 3)#This needs to be changed
      pf_cell <- createCell(row = rows, colIndex = 2)
      setCellValue(pf_cell[[1,1]], value = paste("Design:", toupper(design.type)) )
      setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
      
      #Create the cellstyle for the next few subtitles
      CS <- CS + Font(wb = pdwb, heightInPoints = 14, name = "Times New Roman", isBold = TRUE)#General specs, like plot dimension, number of entries, and number of duplicates
      
      #Number of test plots
      rows <- createRow(sheet = main.sheet, rowIndex = 4)#This needs to be changed
      pf_cell <- createCell(row = rows, colIndex = 2)
      setCellValue(pf_cell[[1,1]], value = paste(dim(fld.dgn)[1], "Total test plots", sep = " "))
      setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
      
      #Dimensions of plot
      rows <- createRow(sheet = main.sheet, rowIndex = 5)#This needs to be changed
      pf_cell <- createCell(row = rows, colIndex = 2)
      setCellValue(pf_cell[[1,1]], value = paste(max(fld.dgn$row), "rows x", max(fld.dgn$column), "columns", sep = " "))
      setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
      
      #If statement to check if the duplicates column is actually present (it potentially could not be present)
      if (!is.null(x = fld.dgn$duplicates)) {
        #If the duplicates column is present, a number of duplicates cell group will be added
        #Number of duplicates
        rows <- createRow(sheet = main.sheet, rowIndex = 6) #This needs to be changed
        pf_cell <- createCell(row = rows, colIndex = 2)
        setCellValue(pf_cell[[1,1]], value = paste(sum(fld.dgn$duplicates == "D"), "duplicate entries", sep = " "))
        setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS)
      }
      
      #This subtitle is always included
      #Cell specifying Number of checks
      rows <- createRow(sheet = main.sheet, rowIndex = 7)#This needs to be changed
      pf_cell <- createCell(row = rows, colIndex = 2)
      setCellValue(pf_cell[[1,1]], value = paste(sum(fld.dgn$line.code != 0, na.rm = T), "check plots:", sep = " "))
      setCellStyle(cell = pf_cell[[1,1]], cellStyle = CS + Font(wb = pdwb, heightInPoints = 14, name = "Times New Roman", isItalic = TRUE, isBold = TRUE))
      
      
      #Create separately formatted cells for the checks
      #Create the cellstyle
      CS <- CS +
        Font(wb = pdwb, heightInPoints = 12, name = "Times New Roman", isBold = TRUE)
      
      checks_df[,2] <- as.vector(checks_df[,2]) #Designate the names of this list to be the same as the checks
      
      #If check.color is not user-defined, set it as the first n colors of rainbow
      check.color <- rainbow(n = dim(checks_df)[1])
      
      for (i in 1:dim(checks_df)[1]) { #For loop to iterate through the number of checks
        rows <- createRow(sheet = main.sheet, rowIndex = 7 + i) #This 7 needs to be substituted with the number of rows requested in arguments
        checksumcell <- createCell(row = rows, colIndex = 3)
        checknamecell <- createCell(row = rows, colIndex = 4)
        setCellValue(checksumcell[[1,1]], value = sum(fld.dgn$line.code == i, na.rm = T))
        setCellValue(checknamecell[[1,1]], value = unique(fld.dgn$line.name[which(fld.dgn$line.code == i)]))
        setCellStyle(cell = checksumcell[[1,1]], cellStyle = CS + Fill(foregroundColor = check.color[i], backgroundColor = "white"))
        setCellStyle(cell = checknamecell[[1,1]], cellStyle = CS + Fill(foregroundColor = check.color[i], backgroundColor = "white"))
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
      #This matches the names that are called by the numeric vector dgn.names to the names present in the design data.frame that will be added to the worksheet. If the names do not match, they will not be added (note that rep does not match because the RCBD design output has "replication" instead)
      design_tempdf <- fld.dgn[,match(x = design_dfnames[dgn.names], table = names(fld.dgn))[which(!is.na(x = match(x = design_dfnames[dgn.names], table = names(fld.dgn))))]]
      #Add names to the data.frame
      names(design_tempdf)[1:dim(design_tempdf)[2]] <- names_vec[match(x = names(design_tempdf)[1:dim(design_tempdf)[2]], table = names(fld.dgn))]
      
      # Add the source column
      # Add empty data as filler for the "source" column
      design_tempdf[,(dim(design_tempdf)[2]+1)] <- rep("", dim(design_tempdf)[1])
      #Add the "Source" column name
      names(design_tempdf)[dim(design_tempdf)[2]] <- "Source"
      
      #Rearrange the data.frame so that the "source" column follows the "line.code" column or is in the 5th column position
      if (11 %in% dgn.names) {
        #Add the column number of the last column (and therefore the source column) to the position following the number referring to the "line.code" column
        design_tempdf <- design_tempdf[,append(x = 1:(length(design_tempdf)-1), values = dim(design_tempdf)[2], after = which(dgn.names == 11))]
        #If "line.code" is not present, then the "source" column is defaulted to the 5th column position or last position
      } else {
        if(dim(design_tempdf)[2] > 4) {
          design_tempdf <- design_tempdf[,append(x = 1:(length(design_tempdf)-1), values = dim(design_tempdf)[2], after = 4)]
        }}
      
      #Add the dgn data.frame
      addDataFrame(x = design_tempdf, sheet = main.sheet, col.names = TRUE, row.names = FALSE, startRow = 7 + dim(checks_df)[1] + 2, colnamesStyle = CS)  #This 7 needs to be substituted with the number of rows requested in arguments
      
      #Write the plot layout data to the worksheet
      addDataFrame(x = as.data.frame(fld.plot), sheet = main.sheet, col.names = FALSE, row.names = FALSE, startRow = 2, startColumn = dim(design_tempdf)[2] + 4)
      
      #Write the check layout data to the worksheet
      addDataFrame(x = as.data.frame(line.code), sheet = main.sheet, col.names = FALSE, row.names = FALSE, startRow = (1 + dim(fld.plot)[1] + 4), startColumn = dim(design_tempdf)[2] + 4, colStyle = CellStyle(wb = pdwb, dataFormat = DataFormat(x = "0.0")))
      
      
      
      ##### Borders, highlighting, and other formatting #####
      
      #Only perform highlighting if the line.name column is present
      if (!is.na(match(12, dgn.names))) {
        
        #Highlighting the cells in the design data.frame that are checks
        #First need to redesignate the cellstyle for checks to eliminate boldness
        CS <- CellStyle(wb = pdwb)
        
        #Next, iterate through all rows of the design data.frame, check to see if the cell value is a check, then assign the appropriate cell style
        for (n in 1:dim(fld.dgn)[1]) { #For loop to iterate through every row in the design data.frame
          rows <- getRows(sheet = main.sheet, rowIndex = n + 7 + dim(checks_df)[1] + 2) # 7 = number of rows taken up by pertinant primary info, dim(checks_df) = rows taken up by checks, 2 = space between checks and the data.frame  #This 7 needs to be substituted with the number of rows requested in arguments
          #If line.code, line.name, and plot are not together, then only highlight line.name
          if (!is.na(match(NA, match(c(4,11,12), dgn.names)) > 0) || max(match(c(4,11,12), dgn.names)) - min(match(c(4,11,12), dgn.names)) > 2) {
            colin <- match(12, dgn.names)
            pf_cell <- getCells(row = rows, colIndex = sort(colin))
            pf_value <- getCellValue(cell = pf_cell[[1]])
            
            #Highlight just line.name
            if (sum(pf_value == checks_df[,2]) == 1) { #If statement to ask whether the entry names matches any of the checks
              setCellStyle(cell = pf_cell[[1]], cellStyle = CS + Fill(foregroundColor = check.color[which(pf_value == checks_df[,2])], backgroundColor = "white")) #If the value of the cell matches any of the names of csCheckdflist (and therefore any of the checks), the cellstyle will be set to the style designated for that check
            }
          } else {
            #If Source is true and the three parameters are together, then highlight all four
            colin <- append(x = match(c(4,11,12), dgn.names), values = max(match(c(4,11,12), dgn.names)) + 1)
            pf_cell <- getCells(row = rows, colIndex = sort(colin)) 
            pf_value <- getCellValue(cell = pf_cell[[2]])
            
            if (sum(pf_value == checks_df[,2]) == 1) { #If statement to ask whether the entry names matches any of the checks
              setCellStyle(cell = pf_cell[[1]], cellStyle = CS + Fill(foregroundColor = check.color[which(pf_value == checks_df[,2])], backgroundColor = "white")) #If the value of the cell matches any of the names of csCheckdflist (and therefore any of the checks), the cellstyle will be set to the style designated for that check
              setCellStyle(cell = pf_cell[[2]], cellStyle = CS + Fill(foregroundColor = check.color[which(pf_value == checks_df[,2])], backgroundColor = "white"))
              setCellStyle(cell = pf_cell[[3]], cellStyle = CS + Fill(foregroundColor = check.color[which(pf_value == checks_df[,2])], backgroundColor = "white"))
              setCellStyle(cell = pf_cell[[4]], cellStyle = CS + Fill(foregroundColor = check.color[which(pf_value == checks_df[,2])], backgroundColor = "white"))
            }}}
        
        #Highlighting check positions in the ploy.layout data.frame
        for (r in 1:(dim(fld.plot)[1]-1)) { #For loop to cycle through the number of rows
          for (k in 1:(dim(fld.plot)[2]-1)) { #For loop to cycle through the number of columns
            rows <- getRows(sheet = main.sheet, rowIndex = r + 1) #
            pf_cell <- getCells(row = rows, colIndex = (k + dim(design_tempdf)[2] + 4)) #MAY NEED TO CHANGE COLUMN INDEX FOR SOFT-CODING
            pf_value <- getCellValue(cell = pf_cell[[1]])
            
            if (sum(fld.dgn$line.name[which(fld.dgn$plot == pf_value)] == checks_df[,2]) == 1) {
              setCellStyle(cell = pf_cell[[1]], cellStyle = CS + Fill(foregroundColor = check.color[which(fld.dgn$line.name[which(fld.dgn$plot == pf_value)] == checks_df[,2])], backgroundColor = "white"))  
            } #Close the if statement
          }} #Close the for loops
        
        #Highlighting check positions in the check.layout data.frame
        for (r in 1:(dim(line.code)[1]-1)) { #For loop to cycle through the number of rows
          for (k in 1:(dim(line.code)[2]-1)) { #For loop to cycle through the number of columns
            rows <- getRows(sheet = main.sheet, rowIndex = (r + dim(fld.plot)[1] + 4)) #
            pf_cell <- getCells(row = rows, colIndex = (k + dim(design_tempdf)[2] + 4)) #MAY NEED TO CHANGE COLUMN INDEX FOR SOFT-CODING
            pf_value <- getCellValue(cell = pf_cell[[1]])
            
            if (pf_value != 0) { #If statement to proceed on the condition that the value is not 0 (and is therefore a check)
              setCellStyle(cell = pf_cell[[1]], cellStyle = CS + Fill(foregroundColor = check.color[as.numeric(pf_value)], backgroundColor = "white"))
            } #Close the if statement
          }} #Close the for loops
      } #Close the initial if statement
      
      
      # Manipulating the blocks created in the design
      # If statement to check what design format is being used
      # This first procedure is used for MADII or AIBD
      if (design.type == "aibd" || design.type == "mad") {
        
        #Create numeric vector containing the block dimensions  
        blk_dim <- numeric()
        blk_dim[1] <- max(fld.dgn$row)/max(fld.dgn[,9]) #Add the number of rows in each block
        blk_dim[2] <- max(fld.dgn$column)/max(fld.dgn[,10]) #Add the column dimension (this is done by finding various maxima and minima in the dgn data.frame)
        
        for (r in 1:max(fld.dgn[,9])) { #For loop to cycle through the number of row_blk
          
          for (k in 1:max(fld.dgn[,10])) { #For loop to cycle through the number of col_blk
            
            #Manipulate the border of the blocks by creating cell blocks of the cells
            #Create a cell block that includes the cells in block 9 (for example)
            cellblk1 <- CellBlock(sheet = main.sheet, #Designate the sheet
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
        if (design.type == "rcbd") {
          
          #Create a border around the whole plot map first to simplify things later
          #First create a cellblock containing all of the cells
          cellblk1 <- CellBlock(sheet = main.sheet, #Designate the sheet
                                startRow = 2, #The plot layout will usually start in row 2 
                                startColumn = ((dim(design_tempdf)[2] + 4 + 1)), #Start column is the same as the start column that was used for placing the plot.layout data + 1 to exclude the row names
                                noRows = (dim(fld.plot) - 1)[1], #Number of rows is the number of rows in the field design
                                noColumns = (dim(fld.plot) - 1)[2], #Number of cols is the number of cols in the field design
                                create = FALSE)
          
          #Add the perimeter border
          #Set borders, first the top left
          CB.setBorder(cellBlock = cellblk1, border = Border(position = c("TOP", "LEFT"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = 1)
          #Top right
          CB.setBorder(cellBlock = cellblk1, border = Border(position = c("TOP", "RIGHT"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = (dim(fld.plot) - 1)[2])
          #Bottom right
          CB.setBorder(cellBlock = cellblk1, border = Border(position = c("BOTTOM", "RIGHT"), pen = "BORDER_MEDIUM"), rowIndex = (dim(fld.plot) - 1)[1], colIndex = (dim(fld.plot) - 1)[2])
          #Bottom left
          CB.setBorder(cellBlock = cellblk1, border = Border(position = c("BOTTOM", "LEFT"), pen = "BORDER_MEDIUM"), rowIndex = (dim(fld.plot) - 1)[1], colIndex = 1)
          #Top row
          CB.setBorder(cellBlock = cellblk1, border = Border(position = c("TOP"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = 2:((dim(fld.plot) - 1)[2]-1)) #Columns are added by starting at the second column (2), then including all columns up until the last one (blk_dim[2]-1)
          #Right side
          CB.setBorder(cellBlock = cellblk1, border = Border(position = c("RIGHT"), pen = "BORDER_MEDIUM"), rowIndex = 2:((dim(fld.plot) - 1)[1]-1), colIndex = (dim(fld.plot) - 1)[2]) #Rows are added by starting at the second row (2), then including all rows up until the last one (blk_dim[1]-1)
          #Bottom row
          CB.setBorder(cellBlock = cellblk1, border = Border(position = c("BOTTOM"), pen = "BORDER_MEDIUM"), rowIndex = (dim(fld.plot) - 1)[1], colIndex = 2:((dim(fld.plot) - 1)[2]-1))
          #Left side
          CB.setBorder(cellBlock = cellblk1, border = Border(position = c("LEFT"), pen = "BORDER_MEDIUM"), rowIndex = 2:((dim(fld.plot) - 1)[1]-1), colIndex = 1)
          
          
          #Iterating through all of the plots to draw inside borders (RCBDs can have funky borders for the replicates)
          #For loop through each row
          for (r in 1:max(fld.dgn$row)) { #For loop to cycle through the number of rows
            for (k in 1:max(fld.dgn$column)) { #For loop to cycle through the number of cols
              
              #Create a cellblock with just one cell
              cellblk1 <- CellBlock(sheet = main.sheet, #Designate the sheet
                                    startRow = (r + 1), #The plot layout will usually start in row 2 
                                    startColumn = (k + dim(design_tempdf)[2] + 4), #Start column is the same as the start column that was used for placing the plot.layout data + 1 to exclude the row names
                                    noRows = 1, #Number of rows is the number of rows in the field design
                                    noColumns = 1, #Number of cols is the number of cols in the field design
                                    create = FALSE)
              
              #Check if the cell is the last in the column
              if (k != max(fld.dgn$column)) {
                
                #Check if the cell to the right of the current cell is the same rep
                if (blk.code[r,(k+1)] != blk.code[r,(k+1)+1]) {
                  
                  #Add a right-side border to the cellstyle
                  CB.setBorder(cellBlock = cellblk1, border = Border(position = c("RIGHT"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = 1)
                }} #Close the if statements
              
              #Check if the cell is the last in the row
              if (r != max(fld.dgn$row)) {
                
                #Check if the cell to the bottom of the current cell is the same rep
                if (blk.code[r,(k+1)] != blk.code[r+1,(k+1)]) {
                  
                  #Add a bottom-side border to the cellstyle
                  CB.setBorder(cellBlock = cellblk1, border = Border(position = c("BOTTOM"), pen = "BORDER_MEDIUM"), rowIndex = 1, colIndex = 1)
                }} #Close the if statements
              
            }} #Column loop and Row loop
          
        } #If statement for RCBD
      } #Close the else statement
      
      
      # Create sheets for the fieldbook information
      field.design.sheet <- createSheet(wb = pdwb, sheetName = "field.design")
      addDataFrame(x = fld.dgn, sheet = field.design.sheet, col.names = T, row.names = F, startRow = 1, startColumn = 1)
      
      field.book.sheet <- createSheet(wb = pdwb, sheetName = "field.book")
      # Add the field.book data.frame
      addDataFrame(x = fld.book, sheet = field.book.sheet, col.names = T, row.names = F, startRow = 1, startColumn = 1)
      
      
      ##### Post-Formatting #####
      #Autosize columns
      autoSizeColumn(sheet = main.sheet, colIndex = c(3:(dim(line.code)[2] + dim(design_tempdf)[2] + 4)))
      autoSizeColumn(sheet = field.design.sheet, colIndex = 1:ncol(fld.dgn))
      autoSizeColumn(sheet = field.book.sheet, colIndex = 1:ncol(fld.book))
      
      #Set the zoom
      setZoom(sheet = main.sheet, numerator = 85, denominator = 100)
      setZoom(sheet = field.design.sheet, numerator = 85, denominator = 100)
      setZoom(sheet = field.book.sheet, numerator = 85, denominator = 100)
      
      #Manual column widths
      setColumnWidth(sheet = main.sheet, colIndex = 1, colWidth = 11)
      
      ##### Closing time #####
      
      # Choose a filename
      filename <- paste(sheet.title, sub("\\.", "_", design.type), "output.xlsx", sep = "_")
      
      #Save the workbook to the disk
      saveWorkbook(wb = pdwb, file = filename)
      
    } #Close the function
    
    # Implement the function
    create.xlsx(fld.dgn = fld.dgn, fld.book = fld.book, fld.plot = fld.plot, line.code = line.code, blk.code = blk.code, design.type = "rcbd")
    
  } # Close the if statement
    
  # Return the list
  return(output)
}



