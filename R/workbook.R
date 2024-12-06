###################################################################################################
# CREATE WORKBOOK

#' @title accessibletablesR::workbook
#' 
#' @description Create the openxlsx workbook where data tables are added.
#' 
#' @details 
#' The workbook function creates a new workbook with the required metadata worksheets and defines 
#' the workbook's font name, colour and sizes.
#' All parameters are optional and preset.
#' If a cover page, contents page, notes page or definitions page are required then set the 
#' parameter to "Yes" when calling the function.
#' autonotes is required if a line is wanted towards the top of the worksheet which lists all the 
#' note numbers associated with the worksheet (set to "Yes" if wanted).
#' Default font is Arial with a black colour and size ranging from 12 to 16 - change if want to when
#' calling the function.
#' title, creator, subject and category refer to the document information properties displayed in 
#' the final Excel workbook.
#' 
#' @param covertab Define whether a cover page is required  (optional)
#' @param contentstab Define whether a contents page is required (optional)
#' @param notestab Define whether a notes page is required (optional)
#' @param autonotes Define whether automated listing of notes associated with a table is required 
#'                  (optional)
#' @param definitionstab Define whether a definitions page is required (optional)
#' @param fontnm Define the font name used in the final output (optional)
#' @param fontcol Define the font colour used in the final output (optional)
#' @param fontsz Define the general font size used in the final output (optional)
#' @param fontszst Define the font size for subtitles used in the final output (optional)
#' @param fontszt Define the font size for titles used in the final output (optional)
#' @param title Define the title to go into the document information in the final output (optional)
#' @param creator Define the creator to go into the document information in the final output 
#'                (optional)
#' @param subject Define the subject to go into the document information in the final output 
#'                (optional)
#' @param category Define the category to go into the document information in the final output  
#'                 (optional)
#'                 
#' @returns 
#' A workbook called zzz_wb_zzz will appear in the global environment. Necessary R packages will be
#' installed.
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

workbook <- function(covertab = NULL, contentstab = NULL, notestab = NULL, autonotes = NULL,
                     definitionstab = NULL, fontnm = "Arial", fontcol = "black",
                     fontsz = 12, fontszst = 14, fontszt = 16, title = NULL, creator = NULL,
                     subject = NULL, category = NULL) {
  
  # Install the required packages if they are not already installed, then load the packages
  
  listofpackages <- base::c("openxlsx", "conflicted", "dplyr", "stringr", "purrr", "rlang")
  packageversions <- base::c("4.2.5.2", "1.2.0", "1.1.2", "1.5.0", "1.0.1", "1.1.0")
  
  for (i in base::seq_along(listofpackages)) {
    
    if (!(listofpackages[i] %in% utils::installed.packages())) {
      
      utils::install.packages(listofpackages[i], dependencies = TRUE, type = "binary")
      
    } else if (listofpackages[i] %in% utils::installed.packages() & 
               utils::packageVersion(listofpackages[i]) < packageversions[i]) {
      
      base::unloadNamespace(listofpackages[i])
      utils::install.packages(listofpackages[i], dependencies = TRUE, type = "binary")
      
    } 
    
  }
  
  # When functions are used in this script, the package from which the function comes from is... 
  # ...specified e.g., dplyr::filter
  # The exception to this is if the functions come from the R base package
  # To ensure there is no unintentional masking of base functions, conflict_prefer_all will set...
  # ...it so base is the package used unless otherwise specified
  
  conflicted::conflict_prefer_all("base", quiet = TRUE)
  `%>%` <- dplyr::`%>%`
  
  # Create tabcontents, covernumrow and table_data2 only for purpose of satisfying R CMD check
  
  tabcontents <- NULL
  covernumrow <- NULL
  table_data2 <- NULL
  
  # Cleaning some of the parameters to be either "Yes" or "No"
  
  if (is.null(covertab)) {
    
    covertab <- "No"
    
  } else if (tolower(covertab) == "no" | tolower(covertab) == "n") {
    
    covertab <- "No"
    
  } else if (tolower(covertab) == "yes" | tolower(covertab) == "y") {
    
    covertab <- "Yes"
    
  }
  
  if (is.null(contentstab)) {
    
    contentstab <- "No"
    
  } else if (tolower(contentstab) == "no" | tolower(contentstab) == "n") {
    
    contentstab <- "No"
    
  } else if (tolower(contentstab) == "yes" | tolower(contentstab) == "y") {
    
    contentstab <- "Yes"
    
  }
  
  if (is.null(notestab)) {
    
    notestab <- "No"
    
  } else if (tolower(notestab) == "no" | tolower(notestab) == "n") {
    
    notestab <- "No"
    
  } else if (tolower(notestab) == "yes" | tolower(notestab) == "y") {
    
    notestab <- "Yes"
    
  }
  
  if (is.null(definitionstab)) {
    
    definitionstab <- "No"
    
  } else if (tolower(definitionstab) == "no" | tolower(definitionstab) == "n") {
    
    definitionstab <- "No"
    
  } else if (tolower(definitionstab) == "yes" | tolower(definitionstab) == "y") {
    
    definitionstab <- "Yes"
    
  }
  
  if (is.null(autonotes)) {
    
    autonotes <- "No"
    
  } else if (tolower(autonotes) == "no" | tolower(autonotes) == "n") {
    
    autonotes <- "No"
    
  } else if (tolower(autonotes) == "yes" | tolower(autonotes) == "y") {
    
    autonotes <- "Yes"
    
  }
  
  # Checking some of the parameters to ensure they are properly populated, if not the function...
  # ...will error
  
  if (length(covertab) > 1 | length(contentstab) > 1 | length(notestab) > 1 | 
      length(autonotes) > 1 | length(definitionstab) > 1) {
    
    stop(strwrap("One or more of covertab, contentstab, notestab, definitionstab and autnotes not 
         populated with a single word (\"Yes\", \"No\")", prefix = " ", initial = ""))
    
  }
  
  if (covertab != "No" & covertab != "Yes" & !is.null(covertab)) {
    
    stop("covertab not set to \"Yes\" or \"No\" or NULL")
    
  }
  
  if (contentstab != "No" & contentstab != "Yes" & !is.null(contentstab)) {
    
    stop("contentstab not set to \"Yes\" or \"No\" or NULL")
    
  }
  
  if (notestab != "No" & notestab != "Yes" & !is.null(notestab)) {
    
    stop("notestab not set to \"Yes\" or \"No\" or NULL")
    
  }
  
  if (definitionstab != "No" & definitionstab != "Yes" & !is.null(definitionstab)) {
    
    stop("definitionstab not set to \"Yes\" or \"No\" or NULL")
    
  }
  
  if (autonotes != "No" & autonotes != "Yes" & !is.null(autonotes)) {
    
    stop("autonotes not set to \"Yes\" or \"No\" or NULL")
    
  }
  
  if (is.null(fontnm) | is.null(fontcol) | is.null(fontsz) | is.null(fontszst) | is.null(fontszt)) {
    
    stop("One or more of fontnm, fontcol, fontsz, fontszst and fontszt has been set to NULL")
    
  }
  
  if (!is.character(fontnm) | !is.character(fontcol)) {
    
    stop("One or both of fontnm and fontcol is not of type character")
    
  }
  
  if (!is.numeric(fontsz) | !is.numeric(fontszst) | !is.numeric(fontszt)) {
    
    stop("One or more of fontsz, fontszst and fontszt os not of type numeric")
    
  }
  
  if (length(fontnm) > 1 | length(fontcol) > 1 | length(fontsz) > 1 | length(fontszst) > 1 | 
      length(fontszt) > 1) {
    
    stop(strwrap("One or more of fontnm, fontcol, fontsz, fontszst and fontszt is more than a single 
         entity", prefix = " ", initial = ""))
    
  }
  
  if (!is.null(title) & !is.character(title)) {
    
    stop("title needs to be of type character")
    
  }
  
  if (!is.null(creator) & !is.character(creator)) {
    
    stop("creator needs to be of type character")
    
  }
  
  if (!is.null(subject) & !is.character(subject)) {
    
    stop("subject needs to be of type character")
    
  }
  
  if (!is.null(category) & !is.character(category)) {
    
    stop("category needs to be of type character")
    
  }
  
  if (length(title) > 1 | length(creator) > 1 | length(subject) > 1 | length(category) > 1) {
    
    stop("One of more of title, creator, subject and category is more than a single entity")
    
  }
  
  if (fontsz < 12 | fontszst < 12 | fontszt < 12) {
    
    warning("GSS accessibility guidelines suggest minimum font size should be 12")
    
  }
  
  # Remove any pre-existing objects from earlier, aborted runs of package
  
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
  
  # Create new workbook
  
  wb <- openxlsx::createWorkbook(title = title, creator = creator, subject = subject,
                                 category = category)
  
  # Add required metadata worksheets to the workbook and create new notes and definitions data...
  # ...frames if wanted
  
  if (covertab == "Yes") {
    
    openxlsx::addWorksheet(wb, "Cover")
    
  }
  
  if (contentstab == "Yes") {
    
    openxlsx::addWorksheet(wb, "Contents")
    
  }
  
  if (notestab == "Yes") {
    
    openxlsx::addWorksheet(wb, "Notes")
    
    notesdf <- data.frame() %>%
      dplyr::mutate("Note number" = "", "Note text" = "", "Applicable tables" = "", "Link1" = "", 
                    "Link2" = "")
    
    assign("notesdf", notesdf, envir = as.environment(acctabs))
    rm(notesdf)
    
  }
  
  if (definitionstab == "Yes") {
    
    openxlsx::addWorksheet(wb, "Definitions")
    
    definitionsdf <- data.frame() %>%
      dplyr::mutate("Term" = "", "Definition" = "", "Link1" = "", "Link2" = "")
    
    assign("definitionsdf", definitionsdf, envir = as.environment(acctabs))
    rm(definitionsdf)
    
  }
  
  # Create a variable zzz_autonotes2_zzz in the global environment for later use
  
  if (autonotes == "Yes") {
    
    autonotes2 <- "Yes"
    
  } else {
    
    autonotes2 <- "No"
    
  }
  
  assign("autonotes2", autonotes2, envir = as.environment(acctabs))
  rm(autonotes2)
  
  # Create variables for font sizes in the global environment for later use
  
  assign("fontsz", fontsz, envir = as.environment(acctabs))
  assign("fontszst", fontszst, envir = as.environment(acctabs))
  assign("fontszt", fontszt, envir = as.environment(acctabs))
  
  openxlsx::modifyBaseFont(wb, fontSize = fontsz, fontColour = fontcol, fontName = fontnm)
  
  assign("wb", wb, envir = as.environment(acctabs))
  rm(wb)
  
}

###################################################################################################