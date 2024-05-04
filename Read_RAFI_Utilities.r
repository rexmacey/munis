#' Find As Of Date in RAFI file
#'
#' @param fn Excel file name including path
#' @param vecSheetNames Character vector of potential names of sheets containing the returns.
#'
#' @return String containing as of date.
#' @export
findAsOfDate <- function(fn, vecSheetNames = c("Expected.Returns", "Expected Returns")) {
  erSheet <- getSheet(fn, vecSheetNames)
  tmp <- readxl::read_excel(fn, sheet = erSheet, .name_repair = "minimal")
  tmp1 <- which("As of Date:" == tmp)
  if (length(tmp1) == 0) return(NULL)
  tmp1 <- tmp1[1]
  colNum <- floor(tmp1/nrow(tmp))
  rowNum <- tmp1 - colNum * nrow(tmp)
  out <- as.character(tmp[rowNum, colNum + 2])
  return(out)
}

#' Get Sheet Named
#' 
#' Given a file name and a vector of potential sheet names, this function opens the
#' file and looks for any sheets with names in the vector, returning the name which is found.
#' 
#' If you are opening a file and expect it to contain a sheet named either "Expected Return"
#' or "Expected.Return", the function will return the name found.   If the name is not found, 
#' NA is returned.  
#'
#' @param fn Excel file name including path
#' @param vecSheetNames Character vector of potential names of sheets
#'
#' @return character string of name found or NA
getSheet <- function(fn, vecSheetNames) {
  sheetNames <- readxl::excel_sheets(fn)
  out <- which(sheetNames %in% vecSheetNames)
  if (length(out) == 0) return(NA)
  return(out[1, drop = TRUE])
}

#' Make Asset Classes (Segments) Uniform
#' 
#' A function to convert raw asset class (segment) names in the RAFI file to a  
#' name based on the map in the classNamesMap table.
#'
#' @param tbl A return table with an "Asset Class" column.
#' @param classNamesMap data frame with at least two columns.  First column is the original (from).   Second column is the new.
#'
#' @return A table with the original asset class names replaced.
makeClassesUniform <- function(tbl, classNamesMap, cName = "Asset Class") {
  # origClassNames <- unlist(tbl[, "Asset Class"])
  origClassNames <- unlist(tbl[, cName])
  # idx <- match(origClassNames, unlist(classNamesMap[, 1]))
  idx <- match(origClassNames, classNamesMap[, 1, drop = TRUE])
  newNames <- classNamesMap[idx, 2, drop = TRUE]
  tbl[, cName] <- newNames
  tbl <- tbl %>% filter(!!rlang::sym(cName) != "NA")
  return(tbl)
}

#' Make Asset Classes (Segments) Uniform
#' 
#' A function to convert raw asset class (segment) names in the RAFI file to a  
#' name based on the map in the classNamesMap table.
#'
#' @param tbl A correlation or covariance table.
#' @param classNamesMap data frame with at least two columns.  First column is the original (from).   Second column is the new.
#'
#' @return A table with the original asset class names replaced.
makeClassesUniformCorrCov <- function(tbl, classNamesMap, cName = "Asset Class") {
  origClassNames <- colnames(tbl)
  idx <- match(origClassNames, unlist(classNamesMap[, 1]))
  newNames <- classNamesMap[idx, 2, drop = TRUE]
  idx <- which(newNames != "NA")
  tbl <- tbl[idx, idx]
  colnames(tbl) <- newNames[idx]
  # tbl <- cbind(`Asset Class` = colnames(tbl), tbl)
  tbl <- cbind(`Asset Class` = colnames(tbl), tbl)
  return(tbl)
}

#' Make Asset Classes (Segments) Uniform
#' 
#' A function to convert raw column names in the RAFI file to a  
#' name based on the map in the colNamesMap table.  
#'
#' @param tbl A correlation or covariance table.
#' @param classNamesMap data frame with at least two columns.  First column is the original (from).   Second column is the new.
#'
#' @return A table with the original asset class names replaced.
makeColNamesUniform <- function(tbl, colNamesMap) {
  colTmp <- colnames(tbl)
  colTmp <- gsub("\r\n", " ", colTmp)
  colTmp <- gsub("\n", " ", colTmp)
  idx <-  match(colTmp, unlist(colNamesMap[,1]))
  newNames <- colNamesMap[idx, 2]
  tbl[, which(newNames == "NA")] <- NULL
  colnames(tbl) <- newNames[newNames != "NA"]
  return(tbl)
}

#' Massage Raw RAFI file
#' 
#' Takes a raw RAFI list (probably read using rafi.read.xls.v4) and makes it more
#' usable by converting asset class and column names to standard names.
#'
#' @param rafi List object containing at least ret, corr and cov data frames (or tibbles)
#' @param colNamesMap Data frame with column Names in RAFI file and standard names to use.
#' @param classNamesMap Data frame with asset class (segment) names in RAFI file and standard names to use. 
#'
#' @return The rafi input object massaged with potential modifications o the ret, corr, and cov items.
#' @export
#' @seealso [rafi.read.xls.v4()]
massageRafi <- function(rafi, colNamesMap, classNamesMap) {
  # check that cov and corr have same dimensions and Asset Class names
  if (!identical(dim(rafi$corr),dim(rafi$cov))) {stop("Covariance and Correlation Matrices do not have matching dimensions")}
  if (sum(rafi$corr$`Asset Class` %in% rafi$cov$`Asset Class`) != nrow(rafi$corr)) {stop("Asset Classs in cov and corr matrices do not match")}
  # select the data from ret and risk matrices for Asset Classes that match.
  segmentsToKeep <- intersect(rafi$ret$`Asset Class`, rafi$corr$`Asset Class`)
  rafi$ret <- rafi$ret[rafi$ret$`Asset Class` %in% segmentsToKeep, ]
  # each idx number represents the asset class in the rafi ret table corresponding
  # to the correlation or covariance Asset Class.
  idxCorr <- match(rafi$corr$`Asset Class`, rafi$ret$`Asset Class`)
  idxCov <- match(rafi$cov$`Asset Class`, rafi$ret$`Asset Class`) 
  rafi$corr <- rafi$corr[!is.na(idxCorr), c(TRUE, !is.na(idxCorr)),]
  rafi$cov <- rafi$cov[!is.na(idxCov), c(TRUE, !is.na(idxCov)),]
  rafi$corr$`Asset Class` <- NULL
  rafi$cov$`Asset Class` <- NULL
  # put in same order
  idxCorr <- match(colnames(rafi$corr),rafi$ret$`Asset Class`)
  idxCov <- match(colnames(rafi$cov),rafi$ret$`Asset Class`)
  rafi$corr <- rafi$corr[idxCorr, idxCorr]
  rafi$cov <- rafi$cov[idxCov, idxCov]
  # at this point, ret, corr, and cov, should have the same asset classes in the
  # same order
  rafi$ret <- makeColNamesUniform(rafi$ret, colNamesMap)
  rafi$ret <- makeClassesUniform(rafi$ret, classNamesMap)
  rafi$corr <- makeClassesUniformCorrCov(rafi$corr, classNamesMap)
  rafi$cov <- makeClassesUniformCorrCov(rafi$cov, classNamesMap)
  # rafi$corr$`Asset Class` <- NULL
  # rafi$cov$`Asset Class` <- NULL
  return(rafi)
}

#' Read RAFI File
#'
#' Reads the RAFI file without needing the ranges that specify where the as of date, returns, or 
#' risk matrices are.  This does little more than read the data from the file.   A separate
#' function massageRafi massages some of the data. 
#'
#' @param fn Excel file name including path
#' @param vecExpRetSheetNames Character vector of potential names of sheets containing the returns
#' @param vecRiskSheetNames Character vector of potential names of sheets containing the correlations
#'
#' @return list with as_of_date character string; ret tibble containing expected returns, correlation and covariances
#' @seealso [massageRafi()] Prepares the RAFI file for further analysis.
#' @export
rafi.read.xls.v4 <- function(fn, vecExpRetSheetNames = c("Expected.Returns", "Expected Returns"), 
                             vecRiskSheetNames = c("Expected.Risk.Matrix", "Expected Correlations")) {
  out <- list()
  out$as_of_date <- findAsOfDate(fn, vecExpRetSheetNames)
  out$ret <- readReturnData(fn, vecExpRetSheetNames)
  out$corr <- readCorrData(fn, vecRiskSheetNames)
  out$cov <- readCovData(fn, vecRiskSheetNames)
  return(out)
}

#' Read Correlation Data from RAFI
#' 
#' Reads the Correlation data from RAFI file.   Does not require a range. The input file
#' should have a tab (sheet) whose name is in the vecSheetNames.  The function looks
#' for two tables on that sheet and returns the top one.
#'
#' @param fn Excel file name including path
#' @param vecSheetNames Character vector of potential names of sheets containing the correlations  
#'
#' @return Tibble containing expected returns from RAFI file. 
readCorrData <- function(fn, vecSheetNames = c("Expected.Risk.Matrix", "Expected Correlations")) {
  riskSheet <- getSheet(fn, vecSheetNames)
  tbl1 <- readxl::read_excel(fn, sheet = riskSheet, .name_repair = "minimal")
  rowData <- apply(tbl1, 1, function(x) ncol(tbl1) - sum(is.na(x)))
  nFields <- median(rowData[rowData != 0])
  colData <- apply(tbl1, 2, function(x) nrow(tbl1) - sum(is.na(x)))
  nClasses <- median(colData[colData != 0])
  
  corRowStart <- min(which(rowData >= nFields))
  corRowEnd <- min(which(rowData[corRowStart + 1:length(rowData)] < nFields)) + corRowStart 
  corColStart <- pmax(1, min(which(colData >= nClasses)))
  corColEnd <- max(which(colData >= max(colData)))
  corRange <- paste0("R", corRowStart, "C", corColStart, ":R", corRowEnd, "C", corColEnd)
  out <- readxl::read_excel(path = fn, sheet = riskSheet, range = corRange, col_names = TRUE, .name_repair = "minimal")
  names(out)[1] <- 'Asset Class'
  return(out)
}
#' Read Covariance Data from RAFI
#' 
#' Reads the covariance data from RAFI file.   Does not require a range. The input file
#' should have a tab (sheet) whose name is in the vecSheetNames.  The function looks
#' for two tables on that sheet and returns the bottom one.
#'
#' @param fn Excel file name including path
#' @param vecSheetNames Character vector of potential names of sheets containing the covariances  
#'
#' @return Tibble containing expected returns from RAFI file. 
readCovData <- function(fn, vecSheetNames = c("Expected.Risk.Matrix", "Expected Correlations")) {
  riskSheet <- getSheet(fn, vecSheetNames)
  tbl1 <- readxl::read_excel(fn, sheet = riskSheet, .name_repair = "minimal")
  rowData <- apply(tbl1, 1, function(x) ncol(tbl1) - sum(is.na(x)))
  nFields <- median(rowData[rowData != 0])
  colData <- apply(tbl1, 2, function(x) nrow(tbl1) - sum(is.na(x)))
  nClasses <- median(colData[colData != 0])
  
  corRowStart <- min(which(rowData >= nFields))
  corRowEnd <- min(which(rowData[corRowStart + 1:length(rowData)] < nFields)) + corRowStart 
  corColStart <- pmax(1, min(which(colData >= nClasses)))
  corColEnd <- max(which(colData >= max(colData)))
  
  covRowStart <- min(which(rowData[corRowEnd + 1:length(rowData)] >= nFields)) + corRowEnd
  covRowEnd <- min(which(rowData[covRowStart + 1:length(rowData)] < nFields)) + covRowStart 
  covRange <- paste0("R", covRowStart, "C", corColStart, ":R", covRowEnd, "C", corColEnd)
  out <- readxl::read_excel(path = fn, sheet = riskSheet, range = covRange, col_names = TRUE, .name_repair = "minimal")
  names(out)[1] <- 'Asset Class'
  return(out)
}

#' Read Return Data from RAFI
#' 
#' Reads the return data from RAFI file.   Does not require a range. The input file
#' should have a tab (sheet) whose name is in the vecSheetNames.  The function looks
#' for a table on that sheet.
#'
#' @param fn Excel file name including path
#' @param vecSheetNames Character vector of potential names of sheets containing the returns.  
#'
#' @return Tibble containing expected returns from RAFI file. 
readReturnData <- function(fn, vecSheetNames = c("Expected.Returns", "Expected Returns")) {
  erSheet <- getSheet(fn, vecSheetNames)
  # riskSheet <- which(sheetNames %in% riskSheetNames)
  
  tbl1 <- readxl::read_excel(fn, sheet = erSheet, .name_repair = "minimal")
  rowData <- apply(tbl1, 1, function(x) ncol(tbl1) - sum(is.na(x)))
  nFields <- median(rowData[rowData != 0])
  colData <- apply(tbl1, 2, function(x) nrow(tbl1) - sum(is.na(x)))
  nClasses <- median(colData[colData != 0])
  retRange <- paste0("R", min(which(rowData >= nFields)) + 1,"C", min(which(colData >= nClasses)), ":R", max(which(rowData >= nFields)) + 1,"C", max(which(colData >= nClasses)))
  
  out <- readxl::read_excel(path = fn, sheet = erSheet, range = retRange, col_names = TRUE, .name_repair = "unique")
  out <- out %>% mutate(`Asset Type` = NULL, Category = NULL)
  return(out)
}

  
# todo. Fix massageRafi.  Probably in make class names
# after massageRafi there are duplicate asset class names in the corr item
# US Treasury Intermediate is renamed US Aggregate

