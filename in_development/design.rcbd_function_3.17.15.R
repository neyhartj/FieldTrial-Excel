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
design.rcbd <- function(enviro=format(Sys.Date(), "%x"), exp.name=NULL, nBlks=2, entries= NULL, nEntries= NULL, chks=NULL, nChks=NULL, nChkReps=1, nFieldRows=NULL, nFieldCols=NULL, plot.start=1001, fillWithEntry=T, dup.list=NULL, fillWithChk=F){
  
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
  
  entry.mat <- cbind(line_name=c(unique(entries), chks, "FILLER"), entry=c(1:(length(unique(entries)) + nChks), NA)) ## This is for field book
  
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
  fld.dgn <- data.frame(environment=enviro, trial=trial, plot_id=paste(trial, plot.start:(plot.start+(exp.size-1)), sep="_"), plot=plot.start:(plot.start+(exp.size-1)), replication=c(rep(1:nBlks, each=(exp.size/nBlks)), rep("FILLER", times=nFill.remainder)), row=rep(1:nFieldRows, each=nFieldCols), column=rep(c(1:nFieldCols, nFieldCols:1), length.out=exp.size), line_code=rep(NA, times=exp.size), line_name=rep(NA, times=exp.size))
    
  ## Fills in FILL plots, and randomizes chks, entries, etc..
  if(fillWithChk) entries.col <- unlist(lapply(1:nBlks, function(b){tmp <- c(entries, rep(chks, times=nChkReps), rep(chks[sample(1:nChks, size = nChks, replace = FALSE)], times=nFill.blk/nChks, length.out=nFill.blk)); tmp[sample(1:length(tmp), replace = FALSE)]}))
  if(fillWithEntry) entries.col <- unlist(lapply(1:nBlks, function(b){tmp <- c(entries, rep(chks, times=nChkReps), rep(dup.list[sample(1:length(dup.list), size = length(dup.list), replace = FALSE)], times= ceiling(nFill.blk/length(dup.list)), length.out=nFill.blk)); tmp[sample(1:length(tmp), replace = FALSE)]}))
  if(!fillWithEntry & !fillWithChk) entries.col <- unlist(lapply(1:nBlks, function(b){tmp <- c(entries, rep(chks, times=nChkReps)); tmp <- c(tmp[sample(1:length(tmp), replace = FALSE)], rep("FILLER", times=nFill.blk))}))
  
  fld.dgn$line_name <- c(entries.col, rep("FILLER", times=nFill.remainder))
  fld.dgn$line_code <- sapply(entries.col, function(X){if(X %in% entries) return(0) ; if(X %in% chks) return(which(chks == X)); if(X == "FILLER") return(NA)})
  
  fld.dgn <- merge(fld.dgn, entry.mat, by = "line_name", all.x = TRUE); fld.dgn <- fld.dgn[order(fld.dgn$plot), ]
  fld.book <- cbind(fld.dgn[,c(3,4,2,9,6,7,10,5)], notes=" ")
  
  if(fillWithEntry){
    fld.dgn$duplicates  <- ""
    for(b in 1:nBlks){
      blk.tmp <- subset(fld.dgn, replication == b)
      dup.entries <- blk.tmp[which(duplicated(blk.tmp$entry) & blk.tmp$line_code %in% 0), "entry"]
      fld.dgn[which(fld.dgn$replication == b & fld.dgn$entry %in% dup.entries), "duplicates"]  <- "D"
    }
  }
  nDups <- length(which(fld.dgn$duplicates == "D"))/2

  per.chks <- round(length(which(fld.dgn[,"line_code"] %in% (1:nChks))) / length(!is.na(fld.dgn[,"line_code"])), digits = 3)
  
  ## Make field.layout with plot numbers
  for.layout.mats <- fld.dgn[order(fld.dgn$row, fld.dgn$column), ]
  
  fld.plot <- apply(t(matrix(for.layout.mats$plot, nrow = nFieldRows, ncol=nFieldCols, byrow=T)), 1, rev)
  fld.plot <- cbind(nFieldRows:1, fld.plot)
  fld.plot <- rbind(fld.plot, c("", 1:nFieldCols))
  
  ## Make field.layout with line codes
  line.code <- apply(t(matrix(for.layout.mats$line_code, nrow = nFieldRows, ncol=nFieldCols, byrow=T)), 1, rev)
  line.code <- cbind(nFieldRows:1, line.code)
  line.code <- rbind(line.code, c("", 1:nFieldCols))
  
  ## Make field.layout with blk codes
  blk.code <- apply(t(matrix(for.layout.mats$replication, nrow = nFieldRows, ncol=nFieldCols, byrow=T)), 1, rev)
  blk.code <- cbind(nFieldRows:1, blk.code)
  blk.code <- rbind(blk.code, c("", 1:nFieldCols))

  return(list(rcbd.dgn=fld.dgn, field.book=fld.book, plot.layout=fld.plot, check.layout=line.code, rep.layout=blk.code, field.dim=paste(nFieldRows, "x", nFieldCols, sep=" "), nDuplicates=nDups, rlzPerChks=per.chks))
}



