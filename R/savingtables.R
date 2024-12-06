###################################################################################################
# SAVING THE FINAL SPREADSHEET

#' @title accessibletablesR::savingtables
#' 
#' @description Saving the final output
#' 
#' @details 
#' The savingtables function only requires that the location and name of the spreadsheet be 
#' specified.
#' A xls file can be saved but it is recommended to use a xlsx file instead.
#' Cannot save directly to ods, instead first a xlsx file is saved and then converted.
#' Default setting is not to create the ods file from the xlsx file.
#' If only xlsx file is wanted keep odsfile = "No" but if both ods and xlsx file wanted set 
#' odsfile = "Yes" and deletexlsx = "No".
#' If only the ods file wanted set odsfile = "Yes" and deletexlsx = "Yes".
#' 
#' @param filename File path and file name of final output, including file type
#' @param odsfile Define whether to convert output to an ods file (optional)
#' @param deletexlsx Define whether to delete the xlsx file output (optional)
#' 
#' @returns A workbook saved to the network drive
#' 
#' @examples 
#' accessibletablesR::workbook(
#'    covertab = "Yes", contentstab = "Yes", notestab = "Yes", definitionstab = "Yes", 
#'    autonotes = "Yes", 
#'    title = "Fuel consumption and aspects of car design and performance for various cars",
#'    creator = "An organisation")
#'                             
#' accessibletablesR::creatingtables(
#'    title = "Fuel consumption and aspects of car design and performance for various cars C",
#'    subtitle = "Cars",
#'    extraline1 = "Link to contents",
#'    extraline2 = "Link to notes",
#'    extraline3 = "Link to definitions",
#'    sheetname = "Table_3", table_data = dummydf, tablename = "thirdtable", headrowsize = 40,
#'    numdatacols = c(2:8,11:13), numdatacolsdp = c(1,0,1,0,2,1,2,0,0,3),
#'    othdatacols = c(9,10), columnwidths = "specified",
#'    colwid_spec = c(18,18,18,15,17,15,12,17,12,13,23,22,12))
#'                                   
#' accessibletablesR::contentstable()
#' 
#' accessibletablesR::addnote(notenumber = "note1", 
#'    notetext = "Google is an internet search engine", applictabtext = "All", linktext1 = "Google",
#'                linktext2 = "https://www.ons.google.co.uk") 
#' 
#' accessibletablesR::notestab()
#' 
#' accessibletablesR::adddefinition(term = "Usual resident", 
#'    definition = "A usual resident is anyone who, on Census Day, 21 March 2021 was in the UK and 
#'                  had stayed or intended to stay in the UK for a period of 12 months or more, or 
#'                  had a permanent UK address and was outside the UK and intended to be outside the
#'                  UK for less than 12 months.")
#'
#' accessibletablesR::definitionstab()
#' 
#' accessibletablesR::coverpage(
#'   title = "Fuel consumption and aspects of car design and performance for various cars",
#'   intro = "Some made up data about cars",
#'   about = "The output of an example of how to use accessibletablesR",
#'   source = "R mtcars",
#'   relatedlink = "https://www.rdocumentation.org/packages/datasets/versions/3.6.2/topics/mtcars)",
#'   relatedtext = "mtcars: Motor trend car road tests",
#'   dop = "26 October 2023",
#'   blank = "There should be no blank cells",
#'   names = "Your name",
#'   email = "yourname@emailprovider.com",
#'   phone = "01111 1111111111111",
#'   reuse = "Yes", govdept = NULL)
#'                              
#' accessibletablesR::savingtables("D:/mtcarsexample.xlsx", odsfile = "Yes", deletexlsx = "No")
#' 
#' @export

savingtables <- function(filename, odsfile = "No", deletexlsx = NULL) {
  
  if (!("openxlsx" %in% utils::installed.packages()) |
      !("stringr" %in% utils::installed.packages()) |
      !("conflicted" %in% utils::installed.packages())) {
    
    stop(base::strwrap("Not all required packages installed. Run the \"workbook\" function first to 
         ensure packages are installed.", prefix = " ", initial = ""))
    
  } else if (utils::packageVersion("openxlsx") < "4.2.5.2" |
             utils::packageVersion("stringr") < "1.5.0" |
             utils::packageVersion("conflicted") < "1.2.0") {
    
    stop(base::strwrap("Older versions of packages detected. Run the \"workbook\" function first to 
         ensure up to date packages are installed.", prefix = " ", initial = ""))
    
  }
  
  conflicted::conflict_prefer_all("base", quiet = TRUE)
  `%>%` <- dplyr::`%>%`
  
  # Create some objects only for purpose of satisfying R CMD check
  
  tabcontents <- NULL
  notesdf <- NULL
  definitionsdf <- NULL
  autonotes2 <- NULL
  covernumrow <- NULL
  table_data2 <- NULL
  fontsz <- NULL
  fontszst <- NULL
  fontszt <- NULL
  
  if (!(exists("wb", envir = as.environment(acctabs)))) {
    
    stop("Run the \"workbook\" function first to ensure that a workbook named wb exists")
    
  }
  
  wb <- acctabs$wb
  
  if (length(filename) > 1) {
    
    stop("filename is not populated properly. It should be a single entity and not a vector.")
    
  }
  
  if (substr(filename, nchar(filename) - 4, nchar(filename)) != ".xlsx" & 
      substr(filename, nchar(filename) - 3, nchar(filename)) != ".xls") {
    
    stop("filename needs to end with \".xlsx\" or \".xls\"")
    
  }
  
  if (substr(filename, nchar(filename) - 3, nchar(filename)) == ".xls") {
    
    warning(strwrap("Ideally filename should end with \".xlsx\". Check if file extension can be 
            changed to \".xlsx\".", prefix = " ", initial = ""))
    
  }
  
  if (stringr::str_detect(filename, " ")) {
    
    warning(strwrap("GSS guidance for spreadsheets includes not using spaces in file names, instead 
            consider using dashes", prefix = " ", initial = ""))
    
  }
  
  if (is.null(odsfile)) {
    
    odsfile <- "Yes"
    
  } else if (tolower(odsfile) == "yes" | tolower(odsfile) == "y") {
    
    odsfile <- "Yes"
    
  } else if (tolower(odsfile) == "no" | tolower(odsfile) == "n") {
    
    odsfile <- "No"
    
  }
  
  if (odsfile != "Yes" & odsfile != "No") {
    
    stop("odsfile not set to \"Yes\" or \"No\"")
    
  }
  
  if (length(odsfile) > 1) {
    
    stop("odsfile is more than a single entity")
    
  }
  
  if (is.null(deletexlsx)) {
    
    deletexlsx <- "Yes"
    
  } else if (tolower(deletexlsx) == "yes" | tolower(deletexlsx) == "y") {
    
    deletexlsx <- "Yes"
    
  } else if (tolower(deletexlsx) == "no" | tolower(deletexlsx) == "n") {
    
    deletexlsx <- "No"
    
  }
  
  if (deletexlsx != "Yes" & deletexlsx != "No") {
    
    stop("deletexlsx not set to \"Yes\" or \"No\"")
    
  }
  
  if (length(deletexlsx) > 1) {
    
    stop("deletexlsx is more than a single entity")
    
  }
  
  # reverse function from https://www.geeksforgeeks.org/how-to-reverse-a-string-in-r/
  
  reverse <- function(str) {
    
    reversedstr <- ""
    
    while (nchar(str) > 0) {
      
      reversedstr <- paste0(reversedstr, substr(str, nchar(str), nchar(str)))
      str <- substr(str, 1, nchar(str) - 1)
      
    }
    
    return(reversedstr)
    
  }
  
  str <- filename
  reversedstr <- reverse(str)
  
  filename2 <- stringr::str_split(reversedstr, c("/"))
  filename3 <- unlist(filename2)
  filename4 <- filename3[[1]]
  filename5 <- substr(filename4, 6, nchar(filename4))
  
  if (grepl("[A-Z]", filename5, perl = TRUE) == TRUE) {
    
    warning("GSS guidance for spreadsheets includes not using upper case letters in file names")
    
  }
  
  rm(filename2, filename3, filename4, filename5)
  
  if ("Cover" %in% names(wb)) {
    
    if (suppressWarnings(length(openxlsx::readWorkbook(wb, "Cover")) == 0)) {
      
      warning("The cover page is empty")
      
    }
    
  }
  
  if ("Contents" %in% names(wb)) {
    
    if (suppressWarnings(length(openxlsx::readWorkbook(wb, "Contents")) == 0)) {
      
      warning("The contents page is empty")
      
    }
    
  }
  
  if ("Notes" %in% names(wb)) {
    
    if (suppressWarnings(length(openxlsx::readWorkbook(wb, "Notes")) == 0)) {
      
      warning("The notes page is empty")
      
    }
    
  }
  
  if ("Definitions" %in% names(wb)) {
    
    if (suppressWarnings(length(openxlsx::readWorkbook(wb, "Definitions")) == 0)) {
      
      warning("The definitions page is empty")
      
    }
    
  }
  
  sheetnames <- names(wb) 
  sheetnames2 <- sheetnames[-which(sheetnames %in% c("Cover", "Contents", "Notes", "Definitions"))]
  
  tablestarts <- unlist(mget(paste0(sheetnames2, "_tablestart"), envir = as.environment(acctabs)))
  
  if (length(unique(tablestarts)) > 1) {
    
    warning(strwrap("The row number of the data table headings is not the same on all worksheets. 
            This might be frustrating for anyone reading the tables into a programming language. 
            Consider whether it would be possible to make the tables start on the same row of each 
            worksheet.", prefix = " ", initial = ""))
    
  }
  
  rm(sheetnames, sheetnames2, tablestarts)
  
  openxlsx::saveWorkbook(wb, filename, overwrite = TRUE)
  
  if (odsfile == "Yes" & deletexlsx == "Yes" & 
      substr(filename, nchar(filename) - 4, nchar(filename)) == ".xlsx") {
    
    if (file.exists(paste0(substr(filename, 1, nchar(filename) - 5), ".ods")) == TRUE) {
      
      file.remove(paste0(substr(filename, 1, nchar(filename) - 5), ".ods"))
      
    }
    
    convert_to_ods(filename)
    
    file.remove(filename)
    
  } else if (odsfile == "Yes" & deletexlsx == "No" & 
             substr(filename, nchar(filename) - 4, nchar(filename)) == ".xlsx") {
    
    if (file.exists(paste0(substr(filename, 1, nchar(filename) - 5), ".ods")) == TRUE) {
      
      file.remove(paste0(substr(filename, 1, nchar(filename) - 5), ".ods"))
      
    }
    
    convert_to_ods(filename)
    
  } else if (odsfile == "Yes" & substr(filename, nchar(filename) - 3, nchar(filename)) == ".xls") {
    
    warning("In order to produce an ods output, filename needs to be a \".xlsx\" extension")
    
  }
  
  if (odsfile == "No") {
    
    warning("For accessibility reasons, consider converting the workbook to an ods file")
    
  }
  
  # Remove data frames and variables from the global environment in case accessible tables needs ...
  # ... to be run again
  
  if (exists("wb", envir = as.environment(acctabs))) {
    
    rm(wb, envir = as.environment(acctabs))
    
  }
  
  if (exists("tabcontents", envir = as.environment(acctabs))) {
    
    rm(tabcontents, envir = as.environment(acctabs))
    
  }
  
  if (exists("notesdf", envir = as.environment(acctabs))) {
    
    rm(notesdf, envir = as.environment(acctabs))
    
  }
  
  if (exists("definitionsdf", envir = as.environment(acctabs))) {
    
    rm(definitionsdf, envir = as.environment(acctabs))
    
  }
  
  if (exists("autonotes2", envir = as.environment(acctabs))) {
    
    rm(autonotes2, envir = as.environment(acctabs))
    
  }
  
  if (exists("covernumrow", envir = as.environment(acctabs))) {
    
    rm(covernumrow, envir = as.environment(acctabs))
    
  }
  
  if (exists("table_data2", envir = as.environment(acctabs))) {
    
    rm(table_data2, envir = as.environment(acctabs))
    
  }
  
  if (length(ls(pattern = "_startrow", envir = as.environment(acctabs)))) {
    
    rm(list = ls(pattern = "_startrow", envir = as.environment(acctabs)), 
       envir = as.environment(acctabs))
    
  }
  
  if (length(ls(pattern = "_tablestart", envir = as.environment(acctabs)))) {
    
    rm(list = ls(pattern = "_tablestart", envir = as.environment(acctabs)), 
       envir = as.environment(acctabs))
    
  }
  
  if (exists("fontsz", envir = as.environment(acctabs))) {
    
    rm(fontsz, envir = as.environment(acctabs))
    
  }
  
  if (exists("fontszst", envir = as.environment(acctabs))) {
    
    rm(fontszst, envir = as.environment(acctabs))
    
  }
  
  if (exists("fontszt", envir = as.environment(acctabs))) {
    
    rm(fontszt, envir = as.environment(acctabs))
    
  }
  
}

###################################################################################################