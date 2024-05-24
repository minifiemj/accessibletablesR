###################################################################################################
# CONTENTS

#' @title accessibletablesR::contentstable
#' 
#' @description Create a contents page for the workbook.
#' 
#' @details 
#' contentstable function creates a table of contents for the workbook.
#' If no contents page wanted, then do not run the contentstable function.
#' gridlines is by default set to "Yes", change to "No" if gridlines are not wanted.
#' Column widths are automatically set unless user defines specific values in colwid_spec.
#' Extra columns can be added, need to set extracols to "Yes" and create a dataframe 
#' extracols_contents with the desired extra columns.
#' 
#' @param gridlines Define whether gridlines are present (optional)
#' @param colwid_spec Define widths of columns (optional)
#' @param extracols Define whether additional columns required (optional)
#' 
#' @returns A worksheet with a contents page of tables in the workbook.
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
#' @importFrom rlang .data
#' 
#' @export

contentstable <- function(gridlines = "Yes", colwid_spec = NULL, extracols = NULL) {
  
  if (!("dplyr" %in% utils::installed.packages()) |
      !("conflicted" %in% utils::installed.packages()) |
      !("openxlsx" %in% utils::installed.packages()) | 
      !("rlang" %in% utils::installed.packages())) {
    
    stop(base::strwrap("Not all required packages installed. Run the \"workbook\" function first to 
         ensure packages are installed.", prefix = " ", initial = ""))
    
  } else if (utils::packageVersion("dplyr") < "1.1.2" |
             utils::packageVersion("conflicted") < "1.2.0" |
             utils::packageVersion("openxlsx") < "4.2.5.2" |
             utils::packageVersion("rlang") < "1.1.0") {
    
    stop(base::strwrap("Older versions of packages detected. Run the \"workbook\" function first to 
         ensure up to date packages are installed.", prefix = " ", initial = ""))
    
  }
  
  conflicted::conflict_prefer_all("base", quiet = TRUE)
  `%>%` <- dplyr::`%>%`
  
  if (!(exists("wb", envir = as.environment(acctabs)))) {
    
    stop("Run the \"workbook\" function first to ensure that a workbook named wb exists")
    
  }
  
  wb <- acctabs$wb
  tabcontents <- acctabs$tabcontents
  fontszt <- acctabs$fontszt
  fontsz <- acctabs$fontsz
  
  # Check to see that a contents page is wanted, based on whether a worksheet was created in the ...
  # ... initial workbook
  
  if (!("Contents" %in% names(wb))) {
    
    stop("contentstab cannot have been set to \"Yes\" in the workbook function call")
    
  }
  
  # Checking some of the parameters to ensure they are properly populated, if not the function ...
  # ... will error or display a warning in the console
  
  if (is.null(gridlines)) {
    
    gridlines <- "Yes"
    
  } else if (tolower(gridlines) == "no" | tolower(gridlines) == "n") {
    
    gridlines <- "No"
    
  } else if (tolower(gridlines) == "yes" | tolower(gridlines) == "y") {
    
    gridlines <- "Yes"
    
  }
  
  if (gridlines != "Yes" & gridlines != "No") {
    
    stop("gridlines has not been set to either \"Yes\" or \"No\"")
    
  }
  
  if (length(gridlines) > 1) {
    
    stop(strwrap("gridlines has not been populated properly. It must be a single word, either 
         \"Yes\" or \"No\".", prefix = " ", initial = ""))
    
  }
  
  if (is.null(extracols)) {
    
    extracols <- "No"
    
  } else if (tolower(extracols) == "no" | tolower(extracols) == "n") {
    
    extracols <- "No"
    
  } else if (tolower(extracols) == "yes" | tolower(extracols) == "y") {
    
    extracols <- "Yes"
    
  }
  
  if (extracols != "Yes" & extracols != "No") {
    
    stop("extracols has not been set to either \"Yes\" or \"No\"")
    
  }
  
  if (length(extracols) > 1) {
    
    stop(strwrap("extracols has not been populated properly. It must be a single word, either 
         \"Yes\" or \"No\".", prefix = " ", initial = ""))
    
  }
  
  # Automatically detecting if notes and definitions tabs are required
  
  if ("Notes" %in% names(wb)) {
    
    notestab <- "Yes"
    
  } else {
    
    notestab <- "No"
    
  }
  
  if ("Definitions" %in% names(wb)) {
    
    definitionstab <- "Yes"
    
  } else {
    
    definitionstab <- "No"
    
  }
  
  # Title and second row of worksheet
  
  title <- "Table of contents"
  extraline1 <- "This worksheet contains one table."
  
  # Creating a data frame with a record for all the worksheets in the workbook
  # Notes and definitions worksheets will be listed before the main data worksheets
  
  if (notestab == "Yes") {
    
    notesdf2a <- data.frame() %>%
      dplyr::mutate("Sheet name" = "", "Table description" = "") %>%
      dplyr::add_row("Sheet name" = "Notes", "Table description" = "Notes")
    
  } else {
    
    notesdf2a <- data.frame() %>%
      dplyr::mutate("Sheet name" = "", "Table description" = "")
    
  }
  
  if (definitionstab == "Yes") {
    
    notesdf2b <- data.frame() %>%
      dplyr::mutate("Sheet name" = "", "Table description" = "") %>%
      dplyr::add_row("Sheet name" = "Definitions", "Table description" = "Definitions")
    
  } else {
    
    notesdf2b <- data.frame() %>%
      dplyr::mutate("Sheet name" = "", "Table description" = "")
    
  }
  
  # To insert additional columns which are not default columns allowed by the function, a ...
  # ... dataframe called "extracols_contents" needs to be created with the extra columns
  
  if (extracols == "Yes" & exists("extracols_contents", envir = .GlobalEnv)) {
    
    extracols_contents <- get("extracols_contents", envir = .GlobalEnv)
    
    if ((nrow(tabcontents) + nrow(notesdf2a) + nrow(notesdf2b)) != nrow(extracols_contents)) {
      
      stop(strwrap("The number of rows in the table of contents is not the same as in the dataframe 
           of extra columns", prefix = " ", initial = ""))
      
    }
    
    if ("Sheet name" %in% colnames(extracols_contents) | 
        "Table description" %in% colnames(extracols_contents)) {
      
      warning(strwrap("There is at least one duplicate column name in the contents table and the 
              extracols_contents dataframe", prefix = " ", initial = ""))
      
    }
    
    df_temp <- dplyr::bind_rows(notesdf2a, notesdf2b, tabcontents) %>%
      dplyr::bind_cols(extracols_contents)
    
    assign("tabcontents", df_temp, envir = as.environment(acctabs))
    rm(df_temp)
    
  } else if (!(exists("extracols_contents", envir = .GlobalEnv))) {
    
    df_temp <- dplyr::bind_rows(notesdf2a, notesdf2b, tabcontents)
    
    assign("tabcontents", df_temp, envir = as.environment(acctabs))
    rm(df_temp)
    
    if (extracols == "Yes") {
      
      warning(strwrap("extracols has been set to \"Yes\" but the dataframe extracols_contents does 
              not exist. No extra columns will be added.", prefix = " ", initial = ""))
      
    }
    
  } else if (extracols == "No" & exists("extracols_contents", envir = .GlobalEnv)) {
    
    warning(strwrap("extracols has been set to \"No\" but a dataframe extracols_contents exist. 
            Check if extra columns are wanted. No extra columns have been added.", prefix = " ",
                    initial = ""))
    
  }
  
  tabcontents <- acctabs$tabcontents
  
  tabcontents2 <- tabcontents %>%
    dplyr::rename(sheet_name = "Sheet name") %>%
    dplyr::group_by(.data$sheet_name) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(.data$count) / dplyr::n())
  
  tabcontents3 <- tabcontents %>%
    dplyr::rename(table_description = "Table description") %>%
    dplyr::group_by(.data$table_description) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(.data$count) / dplyr::n())
  
  # Checks to make sure there is no duplication of tables in the contents
  
  if (tabcontents2$check > 1) {
    
    stop("Duplicated sheet name(s)")
    
  }
  
  if (tabcontents3$check > 1) {
    
    warning("Duplicated table description(s). Explore to see if this is an issue.")
    
  }
  
  rm(tabcontents2, tabcontents3)
  
  # Define some formatting to be used later
  
  titleformat <- openxlsx::createStyle(fontSize = fontszt, textDecoration = "bold")
  extralineformat <- openxlsx::createStyle(wrapText = FALSE, valign = "top")
  normalformat <- openxlsx::createStyle(wrapText = TRUE, valign = "top")
  linkformat <- openxlsx::createStyle(fontColour = "blue", wrapText = TRUE, valign = "top", 
                                      textDecoration = "underline")
  headingsformat <- openxlsx::createStyle(textDecoration = "bold", wrapText = TRUE, border = NULL, 
                                          valign = "top")
  
  openxlsx::addStyle(wb, "Contents", normalformat, rows = 1:(nrow(tabcontents) + 3), 
                     cols = 1:ncol(tabcontents), gridExpand = TRUE)
  
  openxlsx::writeData(wb, "Contents", title, startCol = 1, startRow = 1)
  
  openxlsx::addStyle(wb, "Contents", titleformat, rows = 1, cols = 1)
  
  openxlsx::writeData(wb, "Contents", extraline1, startCol = 1, startRow = 2)
  
  openxlsx::addStyle(wb, "Contents", extralineformat, rows = 2, cols = 1)
  
  openxlsx::addStyle(wb, "Contents", headingsformat, rows = 3, cols = 1:ncol(tabcontents))
  
  openxlsx::writeDataTable(wb, "Contents", tabcontents, tableName = "contents_table", 
                           startRow = 3, startCol = 1, withFilter = FALSE, tableStyle = "none")
  
  numchars <- max(nchar(tabcontents$"Sheet name"))
  
  if (is.null(colwid_spec) & ncol(tabcontents) == 2) {
    
    openxlsx::setColWidths(wb, "Contents", cols = c(1,2), widths = c(max(15, numchars + 3), 100))
    
  } else if (is.null(colwid_spec) & ncol(tabcontents) > 2) {
    
    openxlsx::setColWidths(wb, "Contents", cols = c(1,2,3:ncol(tabcontents)), 
                           widths = c(max(15, numchars + 3), 100, "auto"))
    
  } else if (!is.numeric(colwid_spec) | length(colwid_spec) != ncol(tabcontents)) {
    
    warning(strwrap("colwid_spec has either been provided as non-numeric or a vector of length not 
                    equal to the number of columns in tabcontents. The default column widths have 
                    been used instead.", prefix = " ", initial = ""))
    
    openxlsx::setColWidths(wb, "Contents", cols = c(1,2,3:max(ncol(tabcontents),3)), 
                           widths = c(max(15, numchars + 3), 100, "auto"))
    
  } else if (!is.null(colwid_spec) & is.numeric(colwid_spec) & 
             length(colwid_spec) == ncol(tabcontents)) {  
    
    openxlsx::setColWidths(wb, "Contents", cols = 1:ncol(tabcontents), widths = colwid_spec)            
    
  }
  
  openxlsx::setRowHeights(wb, "Contents", 2, fontsz * (25/12))
  
  contentrows <- nrow(tabcontents)
  
  # Creating hyperlinks so user can quickly navigate through the spreadsheet
  
  for (i in c(4:(3 + contentrows))) {
    
    openxlsx::writeFormula(wb, "Contents", startRow = i, 
                           x = openxlsx::makeHyperlinkString(paste0(tabcontents[i-3, 1]), 
                                                             row = 1, col = 1, 
                                                             text = paste0(tabcontents[i-3, 1])))
    openxlsx::addStyle(wb, "Contents", linkformat, rows = i, cols = 1)
    
  }
  
  # Remove gridlines if they are not wanted
  
  if (gridlines == "No") {
    
    openxlsx::showGridLines(wb, "Contents", showGridLines = FALSE)
    
  }
  
  assign("wb", wb, envir = as.environment(acctabs))
  
}  

###################################################################################################