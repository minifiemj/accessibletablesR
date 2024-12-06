###################################################################################################
# MAIN TABLES

#' @title accessibletablesR::creatingtables
#' 
#' @description Create a worksheet, formatted to meet accessibility criteria, containing data table.
#' 
#' @details 
#' The creatingtables function will create a worksheet with all the data and annotations.
#' title, sheetname and table_data are the only compulsory parameters.
#' All other parameters are optional and most are preset to NULL, so only need to be defined if 
#' they are wanted.
#' sheetname is what you want the sheet to be called in the published workbook
#' table_data is the name of the R dataframe containing the data to be included in the published 
#' workbook.
#' headrowsize is the height of the row containing the table column names.
#' numdatacols is the column position(s) of columns containing number data values (character or 
#' numeric class) - it is useful for right aligning data columns and inserting thousand commas.
#' numdatacolsdp is the number of desired decimal places for columns with numbers (character or 
#' numeric class).
#' othdatacols is the column position(s) of columns containing data values that are not numbers 
#' (e.g., text, dates) - it is useful for formatting (although at present it seems dates do not 
#' obey the given formatting).
#' Character class data columns will have thousand commas inserted as long as the column position 
#' is identified in numdatacols.
#' Numeric class data columns will only have thousand commas inserted if numdatacolsdp is populated.
#' numdatacolsdp either should be one value which will be applied to all numdatacols columns or a 
#' vector the same length as numdatacols.
#' For character variables, the figure in Excel will only be the value rounded to the specified 
#' number of decimal places. For numeric variables, the figure in Excel will be maintained but the 
#' displayed figure will be the value rounded to the specified number of decimal places.
#' Enter 0 in numdatacolsdp if no decimal places wanted. If an element in numdatacols represents a 
#' non-character and non-numeric class column, enter 0 in the corresponding position in 
#' numdatacolsdp.
#' tablename is the name of the table within the worksheet that a screen reader will detect. It is 
#' automatically selected to be the same as the sheetname unless tablename is populated.
#' gridlines is preset to "Yes", change to "No" if gridlines are not wanted.
#' columnwidths is preset to "R_auto" which allows openxlsx to automatically determine column 
#' widths. If automatic width determination is not wanted, set to NULL.
#' columnwidths can alternatively be set to "characters" which will base the column widths on the 
#' number of characters in a column cell.
#' If columnwidths = "characters" then width_adj can be modified. width_adj adds an additional few 
#' spaces to the number of characters in a column cell.
#' width_adj can either be one value which will be applied to all columns or a vector the same 
#' length as the number of columns in the table.
#' If you want to specify the exact width of each column, set columnwidths = "specified" and provide 
#' the widths in colwid_spec (e.g., colwid_spec = c(3,4,5)).
#' If a link to the contents page is required, set one of the extralines to "Link to contents".
#' If a link to the notes page is required, set one of the extralines to "Link to notes".
#' If a link to the definitions page is required, set one of the extralines to "Link to 
#' definitions".
#' extralines1-6 can be set to hyperlinks - e.g., extraline5 = "[BBC](https://www.bbc.co.uk)".
#' 
#' @param title Title of worksheet
#' @param subtitle Subtitle of worksheet (optional)
#' @param extraline1 First extra line above main data (optional)
#' @param extraline2 Second extra line above main data (optional)
#' @param extraline3 Third extra line above main data (optional)
#' @param extraline4 Fourth extra line above main data (optional)
#' @param extraline5 Fifth extra line above main data (optional)
#' @param extraline6 Sixth extra line above main data (optional)
#' @param sheetname Tab name
#' @param table_data Name of table within R global environment
#' @param headrowsize Height of row containing column headings (optional)
#' @param numdatacols Position of columns in table containing number data (optional)
#' @param numdatacolsdp Number of decimal places wanted for each column of number data (optional)
#' @param othdatacols Position of columns in table containing non-number data (optional)
#' @param tablename Name for table in final output (optional)
#' @param gridlines Define whether gridlines are present (optional)
#' @param columnwidths Define method for assigning widths of columns (optional)
#' @param width_adj Additional width adjustment for columns (optional)
#' @param colwid_spec Define widths of columns (optional)
#' 
#' @returns A worksheet with data formatted to meet accessibility criteria.
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

creatingtables <- function(title, subtitle = NULL, extraline1 = NULL, extraline2 = NULL, 
                           extraline3 = NULL, extraline4 = NULL, extraline5 = NULL, 
                           extraline6 = NULL, sheetname, table_data, headrowsize = NULL, 
                           numdatacols = NULL, numdatacolsdp = NULL, othdatacols = NULL, 
                           tablename = NULL, gridlines = "Yes", columnwidths = "R_auto", 
                           width_adj = NULL, colwid_spec = NULL) {
  
  if (!("dplyr" %in% utils::installed.packages()) | 
      !("conflicted" %in% utils::installed.packages()) | 
      !("openxlsx" %in% utils::installed.packages()) | 
      !("stringr" %in% utils::installed.packages()) | 
      !("purrr" %in% utils::installed.packages()) | !("rlang" %in% utils::installed.packages())) {
    
    stop(base::strwrap("Not all required packages installed. Run the \"workbook\" function first to 
         ensure packages are installed.", prefix = " ", initial = ""))
    
  } else if (utils::packageVersion("dplyr") < "1.1.2" | 
             utils::packageVersion("conflicted") < "1.2.0" |
             utils::packageVersion("openxlsx") < "4.2.5.2" | 
             utils::packageVersion("stringr") < "1.5.0" |
             utils::packageVersion("purrr") < "1.0.1" |
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
  autonotes2 <- acctabs$autonotes2
  fontsz <- acctabs$fontsz
  fontszst <- acctabs$fontszst
  fontszt <- acctabs$fontszt
  
  if (exists("tabcontents", envir = as.environment(acctabs))) {
    
    tabcontents <- acctabs$tabcontents
    
  }
  
  
  # Checking some of the parameters to ensure they are properly populated, if not the function...
  # ...will error or display a warning in the console
  
  table_data_temp <- table_data %>%
    dplyr::mutate(dplyr::across(dplyr::everything(), ~ dplyr::case_when(.x == "" ~ NA,
                                                                        TRUE ~ .x)))
  
  if (nrow(table_data_temp) == 0) {
    
    warning("There are no rows in the data. Check to make sure there is not a problem.")
    
  }
  
  table_data_temp2 <- table_data_temp %>%
    dplyr::mutate(numcellmiss = rowSums(is.na(.))) %>%
    dplyr::mutate(numcellmiss2 = dplyr::case_when(numcellmiss == ncol(table_data_temp) ~ 1,
                                                  TRUE ~ 0)) %>%
    dplyr::summarise(numcellmiss3 = max(.data$numcellmiss2, na.rm = TRUE))
  
  if (table_data_temp2[["numcellmiss3"]] == 1) {
    
    warning(strwrap("There is at least one row missing data in all columns. Check to make sure there 
            is not a problem. This code is not designed to produce worksheets with more than one 
            table in them.", prefix = " ", initial = ""))
    
  }
  
  rm(table_data_temp2)
  
  colsmiss <- colSums(is.na(table_data_temp))
  colsmiss2 <- 0
  
  for (i in seq_along(colsmiss)) {
    
    if (colsmiss[i] == nrow(table_data_temp)) {colsmiss2 <- 1}
    
  }
  
  if (colsmiss2 == 1) {
    
    warning(strwrap("There is at least one column missing data. Check to make sure there is not a 
            problem. This code is not designed to produce worksheets with more than one table in 
            them.", prefix = " ", initial = ""))
    
  }
  
  rm(colsmiss, colsmiss2)
  
  if (any(is.na(table_data_temp)) == TRUE) {
    
    warning(strwrap("There are some blank cells present in the data. Check to make sure these are 
            not a problem.", prefix = " ", initial = ""))
    
  }
  
  rm(table_data_temp)
  
  if (is.null(title) | is.null(sheetname) | is.null(table_data)) {
    
    stop("No title or sheetname or table_data entered. Must have title, sheetname and table_data.")
    
  }
  
  if (title == "" | sheetname == "") {
    
    stop("No title or sheetname. Must have title and sheetname.")
    
  }
  
  if (length(title) > 1 | length(subtitle) > 1 | length(sheetname) > 1 | length(headrowsize) > 1 | 
      length(tablename) > 1 | length(gridlines) > 1) {
    
    stop(strwrap("One or more of title, subtitle, sheetname, headrowsize, tablename and gridlines 
         are not populated properly. They must be a single entity and not a vector.", prefix = " ",
                 initial = ""))
    
  }
  
  if (length(extraline1) > 1 | length(extraline2) > 1 | length(extraline3) > 1 | 
      length(extraline4) > 1 | length(extraline5) > 1 | length(extraline6) > 1) {
    
    warning(strwrap("One or more of extraline1, extraline2, extraline3, extraline4, extraline5 and 
            extraline6 is a vector. Check that this is intentional.", prefix = " ", initial = ""))
    
  }
  
  if (any(duplicated(c(title, subtitle, extraline1, extraline2, extraline3, extraline4,
                       extraline5, extraline6))) == TRUE) {
    
    warning(strwrap("There is duplicated text somewhere in the title, subtitle and extralines1-6. 
            Check that this is intentional.", prefix = " ", initial = ""))
    
  }
  
  if (nchar(sheetname) > 31) {
    
    stop("The number of characters in sheetname must not exceed 31")
    
  }
  
  if (!is.null(sheetname) & is.numeric(sheetname)) {
    
    sheetname <- as.character(sheetname)
    warning("sheetname is numeric and has been changed to character class")
    
  } else if (!is.null(sheetname) & !is.character(sheetname)) {
    
    stop("sheetname must be of character class, ideally not with a number as the first character")
    
  }
  
  if ("Cover" %in% names(wb) & (tolower(sheetname) == "cover")) {
    
    stop("sheetname cannot be set to \"Cover\" if a cover page is desired")
    
  }
  
  if ("Contents" %in% names(wb) & (tolower(sheetname) == "contents")) {
    
    stop("sheetname cannot be set to \"Contents\" if a contents page is desired")
    
  }
  
  if ("Notes" %in% names(wb) & (tolower(sheetname) == "notes")) {
    
    stop("sheetname cannot be set to \"Notes\" if a notes page is desired")
    
  }
  
  if ("Definitions" %in% names(wb) & (tolower(sheetname) == "definitions")) {
    
    stop("sheetname cannot be set to \"Definitions\" if a definitions page is desired")
    
  }
  
  if (!grepl("\\D", sheetname, perl = TRUE) == TRUE) {
    
    warning(strwrap("sheetname is only comprised of numbers - this can cause an issue when opening 
            up the final spreadsheet. Ideally sheetname should be a character string, which can
            contain numbers, though the first character should not be a number.", prefix = " ", 
                    initial = ""))
    
  } else if (grepl("\\d", substr(sheetname, 1, 1), perl = TRUE) == TRUE) {
    
    warning(strwrap("sheetname should not start with a number - this can cause an issue when opening 
            up the final spreadsheet. Ideally sheetname should be a character string, which can
            contain numbers, though the first character should not be a number.", prefix = " ", 
                    initial = ""))
    
  }
  
  if (stringr::str_detect(sheetname, " ")) {
    
    sheetname <- stringr::str_replace(sheetname, " ", "_")
    
  }
  
  if ("zzz_temp_zzz" %in% colnames(table_data) | "zzz_temp_zzz2" %in% colnames(table_data) | 
      "zzz_temp_zzz3" %in% colnames(table_data))  { 
    
    stop(strwrap("The variable(s) zzz_temp_zzz or zzz_temp_zzz2 or zzz_temp_zzz3 exist on the table
         data file. The code needs to create temporary variables of the same name. The columns on
         the table data file will have to named differently.", prefix = " ", initial = ""))
    
  }
  
  if (!is.null(numdatacols) & !is.numeric(numdatacols)) {
    
    stop(strwrap("numdatacols either needs to be numeric (e.g., numdatacols = 6 or numdatacols = 
         c(2,5) or NULL", prefix = " ", initial = ""))
    
  }
  
  if (!is.null(numdatacols)) {
    
    if (any(numdatacols <= 0)) {
      
      stop("Column positions for numdatacols cannot be 0 or negative")
      
    } else if (any(numdatacols > ncol(table_data))) {
      
      stop("Column positions for numdatacols should not exceed the number of columns in the data")
      
    }
    
  }
  
  if (!is.null(numdatacolsdp) & !is.numeric(numdatacolsdp)) {
    
    stop(strwrap("numdatacolsdp either needs to be numeric (e.g., numdatacolsdp = 6 or 
         numdatacolsdp = c(2,5) or NULL", prefix = " ", initial = ""))
    
  }
  
  if (!is.null(numdatacolsdp)) {
    
    if (any(numdatacolsdp < 0)) {
      
      stop("numdatacolsdp cannot be negative")
      
    }
    
  }
  
  if (is.null(numdatacols) & !is.null(numdatacolsdp)) {
    
    stop(strwrap("numdatacols has not been populated but numdatacolsdp has. Need to also populate 
         numdatacols or set numdatacolsdp to NULL.", prefix = " ", initial = ""))
    
  }
  
  if (!is.null(numdatacols) & !is.null(numdatacolsdp) & length(numdatacols) > 1 & 
      length(numdatacolsdp) == 1) {
    
    numdatacolsdp <- rep(numdatacolsdp, length(numdatacols))
    warning(strwrap("numdatacols specifies more than one column. numdatacolsdp has only one value 
            and so it has been assumed that this one value represents the number of decimal places 
            required for each column specified by numdatacols.", prefix = " ", initial = ""))
    
  } else if (!is.null(numdatacols) & !is.null(numdatacolsdp) & 
             length(numdatacols) != length(numdatacolsdp)) {
    
    stop(strwrap("The number of elements in numdatacols and numdatacolsdp needs to be the same 
         (e.g., if numdatacols = c(x,y,z) then numdatacolsdp = c(a,b,c)) or numdatacolsdp set to 
         one value to be applied to all columns in numdatacols", prefix = " ", initial = ""))
    
  }
  
  if (any(duplicated(numdatacols)) == TRUE) {
    
    stop("There is at least one column number entered multiple times in numdatacols")
    
  }
  
  if (!is.null(othdatacols) & !is.numeric(othdatacols)) {
    
    stop(strwrap("othdatacols either needs to be numeric (e.g., othdatacols = 6 or othdatacols = 
         c(2,5) or NULL", prefix = " ", initial = ""))
    
  }
  
  if (!is.null(othdatacols)) {
    
    if (any(othdatacols <= 0)) {
      
      stop("Column positions for othdatacols cannot be 0 or negative")
      
    } else if (any(othdatacols > ncol(table_data))) {
      
      stop("Column positions for othdatacols should not exceed the number of columns in the data")
      
    }
    
  }
  
  if (any(duplicated(othdatacols)) == TRUE) {
    
    stop("There is at least one column number entered multiple times in othdatacols")
    
  }
  
  if (any(is.element(numdatacols, othdatacols)) == TRUE) {
    
    stop("There is at least one column number entered in both numdatacols and othdatacols")
    
  }
  
  numcharcols <- NULL
  numcharcolsdp <- NULL
  numericcols <- NULL
  numericcolsdp <- NULL
  
  for (i in seq_along(numdatacols)) {
    
    if (!is.null(numdatacols) & is.character(table_data[[numdatacols[i]]]) == FALSE & 
        is.numeric(table_data[[numdatacols[i]]]) == FALSE & 
        is.integer(table_data[[numdatacols[i]]]) == FALSE) {
      
      warning(strwrap("A column identified as a number column is not of class character or numeric. 
              Check that is intentional.", prefix = " ", initial = ""))
      othdatacols <- append(othdatacols, numdatacols[i])
      
    } else if (is.character(table_data[[numdatacols[i]]]) == TRUE) {
      
      numcharcols <- append(numcharcols, numdatacols[i])
      numcharcolsdp <- append(numcharcolsdp, numdatacolsdp[i])
      
    } else if (is.numeric(table_data[[numdatacols[i]]]) == TRUE | 
               is.integer(table_data[[numdatacols[i]]]) == TRUE) {
      
      numericcols <- append(numericcols, numdatacols[i])
      numericcolsdp <- append(numericcolsdp, numdatacolsdp[i])
      
    }
    
  }
  
  if (!is.null(numericcols) & is.null(numericcolsdp)) {
    
    warning(strwrap("There are data columns of class numeric but the number of decimal places 
            desired has not been specified. This means that thousand commas cannot be inserted 
            automatically. If these commas are desired then consider entering 0 in the appropriate 
            position in numdatacolsdp.", prefix = " ", initial = ""))
    
  }
  
  if (is.null(tablename)) {tablename <- sheetname}
  
  if (!is.null(tablename) & !is.character(tablename)) {
    
    stop("tablename has to be a string if it is not NULL")
    
  } else if (!is.null(tablename) & is.character(tablename) & length(tablename) > 1) {
    
    stop("tablename should be a single expression, not a vector with length > 1")
    
  }
  
  if (!is.null(tablename) & stringr::str_detect(tablename, " ")) {
    
    tablename <- stringr::str_replace(tablename, " ", "_")
    
  }
  
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
  
  if (!is.null(columnwidths)) {columnwidths <- tolower(columnwidths)}
  
  if (is.null(columnwidths)) {
    
    columnwidths <- "Default"
    
  } else if (length(columnwidths) > 1) {
    
    columnwidths <- "R_auto"
    warning(strwrap("columnwidths should be a single word, not a vector. It will be changed back to 
            the default of \"R_auto\".", prefix = " ", initial = ""))
    
  } else if (columnwidths == "none" | columnwidths == "no" | columnwidths == "n" | 
             columnwidths == "default") {
    
    columnwidths <- "Default"
    
  } else if (!is.null(columnwidths) & !is.character(columnwidths)) {
    
    columnwidths <- "R_auto"
    warning(strwrap("columnwidths should be a character string. It will be changed back to the 
            default of \"R_auto\".", prefix = " ", initial = ""))
    
  } else if (columnwidths == "r_auto") {
    
    columnwidths <- "R_auto"
    
  } else if (columnwidths == "character") {
    
    columnwidths <- "characters"
    
  } else if (columnwidths != "r_auto" & columnwidths != "characters" & 
             columnwidths != "specified") {
    
    columnwidths <- "R_auto"
    warning(strwrap("columnwidths has not been set to \"R_auto\" or \"characters\" or \"specified\" 
            or NULL. It will be changed back to the default of \"R_auto\".", prefix = " ",
                    initial = ""))
    
  }
  
  if (!is.null(colwid_spec)) {
    
    if (any(colwid_spec <= 0)) {
      
      stop("Column widths cannot be 0 or negative")
      
    }
    
  }
  
  if (columnwidths == "specified" & is.null(colwid_spec)) {
    
    stop(strwrap("The option to specify column widths has been selected but the widths have not 
         been provided", prefix = " ", initial = ""))
    
  } else if (columnwidths != "specified" & !is.null(colwid_spec)) {
    
    stop(strwrap("The option to specify column widths has not been selected but the widths have been 
         provided", prefix = " ", initial = ""))
    
  } else if (columnwidths == "specified" & length(colwid_spec) == 1 & 
             length(colnames(table_data)) > 1) {
    
    colwid_spec <- rep(colwid_spec, length(colnames(table_data)))
    warning(strwrap("There is more than one column in the table. colwid_spec has only one value and 
            so it has been assumed that this one value represents the widths of all columns.",
                    prefix = " ", initial = ""))
    
  } else if (columnwidths == "specified" & length(colwid_spec) != length(colnames(table_data))) {
    
    stop(strwrap("The number of elements in colwid_spec and the number of columns in the table need 
         to be the same, or colwid_spec set to one value to be applied to all columns in the table",
                 prefix = " ", initial = ""))
    
  }
  
  if (!is.null(width_adj)) {
    
    if (!is.numeric(width_adj)) {
      
      stop("width_adj must be a numeric value")
      
    } else if (length(width_adj) > 1 & length(width_adj) != length(colnames(table_data)) & 
               columnwidths == "characters") {
      
      stop(strwrap("The number of elements in width_adj is not equal to the number of columns in the 
           table data. The number of elements and columns should either be equal or width_adj should 
           be set to only a single value.", prefix = " ", initial = ""))
      
    }
    
  }
  
  # In addition to the title and subtitle, six other fields are permitted above the main data - ...
  # ... these extra fields can be provided as vectors and so there is really no limit to the ...
  # ... number of rows that can come before the main data
  # If a line with information on notes is wanted, this is initially created and existing rows ...
  # ... with information are shifted down one row position
  
  extralines1 <- c(extraline1, extraline2, extraline3, extraline4, extraline5, extraline6)
  
  if ("Notes" %in% names(wb) & autonotes2 == "Yes") {
    
    for (i in seq_along(extralines1)) {
      
      if (stringr::str_detect(extralines1[i], 
                              "This worksheet contains one table|this worksheet contains one table|\\[note")) {
        
        warning(strwrap("If autonotes2 is set to \"Yes\" then the information about the worksheet 
                containing one table or the notes tab will automatically be inserted and so there is
                no need to have one of the extralines already stating this", prefix = " ", 
                        initial = ""))
        
      }
      
    }
    
    extraline6 <- NULL
    
    for (i in 2:(length(extralines1) + 1)) {
      
      if (i < 6) {
        
        x <- extralines1[i-1]
        
        assign(paste0("extraline", i), x)
        
        rm(x)
        
      } else if (i >= 6) {
        
        x <- extralines1[i-1]
        
        extraline6 <- append(extraline6, x)
        
        rm(x)
        
      }
      
    }
    
    extraline1 <- "Temporary holder"
    
    temp <- length(title) + length(subtitle) + 1
    assign(paste0(sheetname, "_startrow"), temp, envir = as.environment(acctabs))
    rm(temp)
    
  } else if ("Notes" %in% names(wb) & autonotes2 == "No") {
    
    onetablenote <- 0
    notescolumn <- 0
    
    for (i in seq_along(extralines1)) {
      
      if (stringr::str_detect(extralines1[i], 
                              "This worksheet contains one table|this worksheet contains one table")) {
        
        onetablenote <- 1
        
      }
      
      if ("Notes" %in% colnames(table_data) | "Note" %in% colnames(table_data) |
          stringr::str_detect(extralines1[i], "\\[note")) {
        
        notescolumn <- 1
        
      }
      
    }
    
    if (onetablenote == 0) {
      
      warning(strwrap("There is no recognisable reference to the worksheet containing one table. 
              Consider whether you want to make a reference to this in one of the extra lines above 
              the main data.", prefix = " ", initial = ""))
      
    }
    
    if (notescolumn == 0) {
      
      warning(strwrap("There is no recognisable notes column or reference to notes. Check whether 
              this is OK.", prefix = " ", initial = ""))
      
    }
    
    rm(onetablenote, notescolumn)
    
  }
  
  # Function to deal with columns containing numbers stored as text, likely as some cells ...
  # ... contain character values (e.g., [c] to indicate some form of statistical disclosure control)
  # The function recognises characters accepted by the GSS as symbols or shorthand applicable ...
  # ... for use in tables (b, c, e, er, f, low, p, r, u, w, x, z)
  # Thousand commas will be inserted if necessary (e.g., 1,340)
  # Function will be called only when a specific number of decimal places is not given
  
  assign("table_data2", table_data, envir = as.environment(acctabs))
  
  numcharvars <- function(numcharcols) {
    
    table_data2 <- acctabs$table_data2
    
    dfx <- table_data2 %>%
      dplyr::mutate(zzz_temp_zzz = 
                      dplyr::case_when(.[[numcharcols]] %in% c("[b]", "[c]", "[e]", "[er]", "[f]",
                                                               "[low]", "[p]", "[r]", "[u]", "[w]",
                                                               "[x]", "[z]", "") ~ "0",
                                       is.na(.[[numcharcols]]) ~ "0",
                                       !is.na(.[[numcharcols]]) 
                                       ~ gsub(",", "", .[[numcharcols]]))) %>%
      dplyr::mutate(zzz_temp_zzz = as.numeric(.data$zzz_temp_zzz)) %>%
      dplyr::mutate(zzz_temp_zzz2 = format(.data$zzz_temp_zzz, big.mark = ",", 
                                           scientific = FALSE)) %>%
      dplyr::mutate(zzz_temp_zzz3 = 
                      dplyr::case_when(.[[numcharcols]] %in% c("[b]", "[c]", "[e]", "[er]", "[f]",
                                                               "[low]", "[p]", "[r]", "[u]", "[w]",
                                                               "[x]", "[z]", "") 
                                       ~ as.character(.[[numcharcols]]),
                                       is.na(.[[numcharcols]]) ~ "",
                                       TRUE ~ as.character(.data$zzz_temp_zzz2)))
    
    dfx[[numcharcols]] <- dfx$zzz_temp_zzz3
    
    dfx_temp <- dfx %>%
      dplyr::select(-.data$zzz_temp_zzz, -.data$zzz_temp_zzz2, -.data$zzz_temp_zzz3)
    
    assign("table_data2", dfx_temp, envir = as.environment(acctabs))
    rm(dfx_temp)
    
  }
  
  # Function to deal with columns containing numbers stored as text, likely as some cells ...
  # ... contain character values (e.g., [c] to indicate some form of statistical disclosure control)
  # The function recognises characters accepted by the GSS as symbols or shorthand applicable ...
  # ... for use in tables (b, c, e, er, f, low, p, r, u, w, x, z)
  # Thousand commas will be inserted if necessary (e.g., 1,340.54)
  # Function will be called only when a specific number of decimal places is given
  
  numcharvars2 <- function(numcharcols, numcharcolsdp) {
    
    table_data2 <- acctabs$table_data2
    
    dfx <- table_data2 %>%
      dplyr::mutate(zzz_temp_zzz = 
                      dplyr::case_when(.[[numcharcols]] %in% c("[b]", "[c]", "[e]", "[er]", "[f]",
                                                               "[low]", "[p]", "[r]", "[u]", "[w]",
                                                               "[x]", "[z]", "") ~ "0",
                                       is.na(.[[numcharcols]]) ~ "0",
                                       !is.na(.[[numcharcols]]) 
                                       ~ gsub(",", "", .[[numcharcols]]))) %>%
      dplyr::mutate(zzz_temp_zzz = if (numcharcolsdp >= 2) as.numeric(.data$zzz_temp_zzz) else 
        round(as.numeric(.data$zzz_temp_zzz), digits = numcharcolsdp)) %>%
      dplyr::mutate(zzz_temp_zzz2 = if (numcharcolsdp >= 2) 
        format(.data$zzz_temp_zzz, big.mark = ",", scientific = FALSE, nsmall = numcharcolsdp) else 
          format(.data$zzz_temp_zzz, big.mark = ",", scientific = FALSE)) %>%
      dplyr::mutate(zzz_temp_zzz3 = 
                      dplyr::case_when(.[[numcharcols]] %in% c("[b]", "[c]", "[e]", "[er]", "[f]",
                                                               "[low]", "[p]", "[r]", "[u]", "[w]",
                                                               "[x]", "[z]", "") 
                                       ~ as.character(.[[numcharcols]]),
                                       is.na(.[[numcharcols]]) ~ "",
                                       TRUE ~ as.character(.data$zzz_temp_zzz2)))
    
    dfx[[numcharcols]] <- dfx$zzz_temp_zzz3
    
    dfx_temp <- dfx %>%
      dplyr::select(-.data$zzz_temp_zzz, -.data$zzz_temp_zzz2, -.data$zzz_temp_zzz3)
    
    assign("table_data2", dfx_temp, envir = as.environment(acctabs))
    rm(dfx_temp)
    
  }
  
  # If there are columns with numbers stored as text then one of the two functions above will ...
  # ... be ran
  # Which function depends on whether the numbers stored as text should have a specific number ...
  # ... of decimal places or not
  # If there are no columns with numbers stored as text then the data are left alone
  
  if (!is.null(numcharcols) & !is.null(numcharcolsdp)) {
    
    purrr::pmap(list(numcharcols, numcharcolsdp), numcharvars2)
    
  } else if (!is.null(numcharcols) & is.null(numcharcolsdp)) {
    
    purrr::pmap(list(numcharcols), numcharvars)
    
  } else if (is.null(numcharcols)) {
    
    assign("table_data2", table_data, envir = as.environment(acctabs))
    
  }
  
  table_data2 <- acctabs$table_data2
  
  # Add the worksheet to the workbook and define various formatting to be used at some point
  
  openxlsx::addWorksheet(wb, sheetname)
  
  extralines2 <- c(extraline1, extraline2, extraline3, extraline4, extraline5, extraline6)
  
  tablestart <- (length(title) + length(subtitle) + length(extralines2) + 1)
  assign(paste0(sheetname, "_tablestart"), tablestart, envir = as.environment(acctabs))
  
  titleformat <- openxlsx::createStyle(fontSize = fontszt, textDecoration = "bold")
  subtitleformat <- openxlsx::createStyle(fontSize = fontszst)
  normalformat <- openxlsx::createStyle(valign = "top")
  linkformat <- openxlsx::createStyle(fontColour = "blue", textDecoration = "underline")
  topformat <- openxlsx::createStyle(valign = "bottom")
  headingsformat <- openxlsx::createStyle(textDecoration = "bold", wrapText = TRUE, border = NULL, 
                                          valign = "top")
  headingsformat2 <- openxlsx::createStyle(textDecoration = "bold", wrapText = TRUE, border = NULL, 
                                           valign = "top", halign = "right")
  dataformat <- openxlsx::createStyle(halign = "right", valign = "top")
  
  openxlsx::addStyle(wb, sheetname, normalformat, rows = 1:(nrow(table_data2) + tablestart),
                     cols = 1:ncol(table_data2), gridExpand = TRUE)
  openxlsx::addStyle(wb, sheetname, topformat, 
                     rows = 1:(length(title) + length(subtitle) + length(extralines2)), cols = 1, 
                     gridExpand = TRUE)
  
  openxlsx::writeData(wb, sheetname, title, startCol = 1, startRow = 1)
  
  openxlsx::addStyle(wb, sheetname, titleformat, rows = 1, cols = 1)
  
  if (!is.null(subtitle)) {
    
    openxlsx::writeData(wb, sheetname, subtitle, startCol = 1, startRow = 2)
    openxlsx::addStyle(wb, sheetname, subtitleformat, rows = 2, cols = 1)
    
  }
  
  # If a link is wanted to the contents or notes or definitions page then the code below will ...
  # ... create the hyperlink
  
  for (i in seq_along(extralines2)) {
    
    if (tolower(extralines2[i]) == "link to notes" | tolower(extralines2[i]) == "notes") {
      
      extralines2[i] <- "Link to notes"
      
    }
    
    if (extralines2[i] == "Link to notes" & !("Notes" %in% names(wb))) {
      
      stop(strwrap("Cannot put a link in to the notes tab unless notestab set to \"Yes\" in the 
           workbook function call", prefix = " ", initial = ""))
      
    }
    
    if (tolower(extralines2[i]) == "link to contents" | tolower(extralines2[i]) == "contents") {
      
      extralines2[i] <- "Link to contents"
      
    }
    
    if (extralines2[i] == "Link to contents" & !("Contents" %in% names(wb))) {
      
      stop(strwrap("Cannot put a link in to the contents tab unless contentstab set to \"Yes\" in 
           the workbook function call", prefix = " ", initial = ""))
      
    }
    
    if (tolower(extralines2[i]) == "link to definitions" | 
        tolower(extralines2[i]) == "definitions") {
      
      extralines2[i] <- "Link to definitions"
      
    }
    
    if (extralines2[i] == "Link to definitions" & !("Definitions" %in% names(wb))) {
      
      stop(strwrap("Cannot put a link in to the definitions tab unless definitionstab set to \"Yes\" 
           in the workbook function call", prefix = " ", initial = ""))
      
    }
    
    if (extralines2[i] == "Link to notes") {
      
      openxlsx::writeFormula(wb, sheetname, startRow = length(title) + length(subtitle) + i, 
                             x = openxlsx::makeHyperlinkString("Notes", row = 1, col = 1, 
                                                               text = "Link to notes"))
      openxlsx::addStyle(wb, sheetname, linkformat, 
                         rows = length(title) + length(subtitle) + i, cols = 1, stack = TRUE)
      
    } else if (extralines2[i] == "Link to contents") {
      
      openxlsx::writeFormula(wb, sheetname, startRow = length(title) + length(subtitle) + i, 
                             x = openxlsx::makeHyperlinkString("Contents", row = 1, col = 1, 
                                                               text = "Link to contents"))
      openxlsx::addStyle(wb, sheetname, linkformat, 
                         rows = length(title) + length(subtitle) + i, cols = 1, stack = TRUE)
      
    } else if (extralines2[i] == "Link to definitions") {
      
      openxlsx::writeFormula(wb, sheetname, startRow = length(title) + length(subtitle) + i, 
                             x = openxlsx::makeHyperlinkString("Definitions", row = 1, col = 1, 
                                                               text = "Link to definitions"))
      openxlsx::addStyle(wb, sheetname, linkformat, 
                         rows = length(title) + length(subtitle) + i, cols = 1, stack = TRUE)
      
    } else if (!is.null(extralines2[i])) {
      
      hyper_rx <- "\\[(([[:graph:]]|[[:space:]])+)\\]\\([[:graph:]]+\\)"
      
      if (grepl(hyper_rx, extralines2[i]) == TRUE) {
        
        if (substr(extralines2[i], 1, 1) != "[" | 
            substr(extralines2[i], nchar(extralines2[i]), nchar(extralines2[i])) != ")") {
          
          warning(strwrap(paste0(extralines2[i], " - if this is meant to be a hyperlink, it needs to 
                  be in the format \"[xxx](xxxxxx)\""), prefix = " ", initial = ""))
          
        }
        
        if ("Link to contents" %in% extralines2 & 
            stringr::str_detect(tolower(extralines2[i]), "\\[link to contents|\\[contents")) {
          
          warning(strwrap(paste0(extralines2[i], " - this appears to be duplicating a link to the 
                  contents page in another extraline parameter"), prefix = " ", initial = ""))
          
        } else if ("Link to notes" %in% extralines2 & 
                   stringr::str_detect(tolower(extralines2[i]), "\\[link to notes|\\[notes")) {
          
          warning(strwrap(paste0(extralines2[i], " - this appears to be duplicating a link to the 
                  notes page in another extraline parameter"), prefix = " ", initial = ""))
          
        } else if ("Link to definitions" %in% extralines2 & 
                   stringr::str_detect(tolower(extralines2[i]), 
                                       "\\[link to definitions|\\[definitions")) {
          
          warning(strwrap(paste0(extralines2[i], " - this appears to be duplicating a link to the 
                  definitions page in another extraline parameter"), prefix = " ", initial = ""))
          
        }
        
        if (stringr::str_detect(tolower(extralines2[i]), "\\[link to contents|\\[link to notes|
                                \\[link to definitions|\\[contents|\\[notes|\\[definitions")) {
          
          warning(strwrap("If you want an internal link to the contents, notes or definitions page, 
                  then set one of extraline1-6 to \"Link to contents\" or \"Link to notes\" or 
                  \"Link to definitions\"", prefix = " ", initial = ""))
          
        }
        
        x <- extralines2[i]
        
        # Hyperlink code taken from Matt Dray's a11ytables
        
        md_rx <- "\\[(([[:graph:]]|[[:space:]])+?)\\]\\([[:graph:]]+?\\)"
        md_match <- regexpr(md_rx, x, perl = TRUE)
        md_extract <- regmatches(x, md_match)[[1]]
        
        url_rx <- "(?<=\\]\\()([[:graph:]])+(?=\\))"
        url_match <- regexpr(url_rx, md_extract, perl = TRUE)
        url_extract <- regmatches(md_extract, url_match)[[1]]
        
        string_rx <- "(?<=\\[)([[:graph:]]|[[:space:]])+(?=\\])"
        string_match <- regexpr(string_rx, md_extract, perl = TRUE)
        string_extract <- regmatches(md_extract, string_match)[[1]]
        
        string_extract <- gsub(md_rx, string_extract, x)
        
        y <- stats::setNames(url_extract, string_extract)
        class(y) <- "hyperlink"
        
        rm(x, md_rx, md_match, md_extract, url_rx, url_match, url_extract, string_rx, string_match, 
           string_extract)
        
      } else {
        
        y <- extralines2[i]
        
      }
      
      openxlsx::writeData(wb, sheetname, y, startCol = 1, 
                          startRow = length(title) + length(subtitle) + i)
      
      if (grepl(hyper_rx, extralines2[i]) == TRUE) {
        
        openxlsx::addStyle(wb, sheetname, linkformat, 
                           rows = length(title) + length(subtitle) + i, cols = 1, stack = TRUE)
        
      }
      
      rm(y)
      
    }
    
  }
  
  openxlsx::addStyle(wb, sheetname, normalformat, rows = tablestart - 1, cols = 1, stack = TRUE)
  
  openxlsx::addStyle(wb, sheetname, headingsformat, rows = tablestart, cols = 1:ncol(table_data2))
  
  # Applying specific formatting to data columns
  
  if (!is.null(numericcols)) {
    
    openxlsx::addStyle(wb, sheetname, headingsformat2, rows = tablestart, 
                       cols = numericcols)
    openxlsx::addStyle(wb, sheetname, dataformat, 
                       rows = (tablestart + 1):(nrow(table_data2) + tablestart + 1), 
                       cols = numericcols, stack = TRUE, gridExpand = TRUE)
    
  }
  
  if (!is.null(numcharcols)) {
    
    openxlsx::addStyle(wb, sheetname, headingsformat2, rows = tablestart, 
                       cols = numcharcols)
    openxlsx::addStyle(wb, sheetname, dataformat, 
                       rows = (tablestart + 1):(nrow(table_data2) + tablestart + 1), 
                       cols = numcharcols, stack = TRUE, gridExpand = TRUE)
    
  }
  
  if (!is.null(othdatacols)) {
    
    openxlsx::addStyle(wb, sheetname, headingsformat2, rows = tablestart, 
                       cols = othdatacols)
    openxlsx::addStyle(wb, sheetname, dataformat, 
                       rows = (tablestart + 1):(nrow(table_data2) + tablestart + 1), 
                       cols = othdatacols, stack = TRUE, gridExpand = TRUE)
    
  }
  
  # If a specific number of decimal places is wanted for numeric columns, the code below will ...
  # ... do this as well as inserting thousand commas
  
  if (!is.null(numericcolsdp)) {
    
    for (i in seq_along(numericcolsdp)) {
      
      if (numericcolsdp[i] > 0) {
        
        fmta <- paste0("#,##0.", strrep("0", numericcolsdp[i]))
        fmt <- openxlsx::createStyle(numFmt = fmta)
        openxlsx::addStyle(wb, sheetname, fmt, 
                           rows = (tablestart + 1):(nrow(table_data2) + tablestart + 1), 
                           cols = numericcols[i], stack = TRUE, gridExpand = TRUE)
        rm(fmta, fmt)
        
      } else if (numericcolsdp[i] == 0) {
        
        fmt <- openxlsx::createStyle(numFmt = "#,##0")
        openxlsx::addStyle(wb, sheetname, fmt, 
                           rows = (tablestart + 1):(nrow(table_data2) + tablestart + 1), 
                           cols = numericcols[i], stack = TRUE, gridExpand = TRUE)
        rm(fmt)
        
      }
      
    }
    
  } 
  
  # Ensure table cell text is wrapped
  
  wrapformat <- openxlsx::createStyle(wrapText = TRUE)
  
  openxlsx::addStyle(wb, sheetname, wrapformat, 
                     rows = (tablestart + 1):(nrow(table_data2) + tablestart + 1), 
                     cols = 1:ncol(table_data2), stack = TRUE, gridExpand = TRUE)
  
  # tablename2 will be the name of the table accessible in Excel
  # If no specific name is given, then the name of the table will be the same as the sheetname
  
  if (!is.null(tablename) & is.character(tablename) & length(tablename) == 1) {
    
    tablename2 <- tablename
    
  } else {
    
    tablename2 <- sheetname
    
  }
  
  # Setting some specific row heights based in part on the font size
  
  openxlsx::writeDataTable(wb, sheetname, table_data2, tableName = tablename2, 
                           startRow = tablestart, startCol = 1, withFilter = FALSE, 
                           tableStyle = "none")
  
  if (length(extralines2) > 0) {
    
    openxlsx::setRowHeights(wb, sheetname, tablestart - 1, fontsz * (25/12))
    
  } else if (length(subtitle) == 1) {
    
    openxlsx::setRowHeights(wb, sheetname, tablestart - 1, fontszst * (25/12))
    
  } else if (length(title) == 1) {
    
    openxlsx::setRowHeights(wb, sheetname, tablestart - 1, fontszt * (25/12))
    
  }
  
  if (!is.null(headrowsize) & is.numeric(headrowsize)) {
    
    openxlsx::setRowHeights(wb, sheetname, tablestart, headrowsize)
    
  }
  
  # Updating the data frame to be used to create a table of contents
  
  if (sheetname != "Contents" & exists("tabcontents", envir = as.environment(acctabs))) {
    
    df_temp <- tabcontents %>%
      dplyr::add_row("Sheet name" = sheetname, "Table description" = title)
    
    assign("tabcontents", df_temp, envir = as.environment(acctabs))
    rm(df_temp)
    
  } else if (sheetname != "Contents" & !exists("tabcontents", envir = as.environment(acctabs))) {
    
    df_temp <- data.frame() %>%
      dplyr::mutate("Sheet name" = "", "Table description" = "") %>%
      dplyr::add_row("Sheet name" = sheetname, "Table description" = title)
    
    assign("tabcontents", df_temp, envir = as.environment(acctabs))
    rm(df_temp)
    
  }
  
  tabcontents <- acctabs$tabcontents
  
  if (exists("tabcontents", envir = as.environment(acctabs))) {
    
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
    
    if (tabcontents2$check > 1) {
      
      stop("Duplicated sheet name(s)")
      
    }
    
    if (tabcontents3$check > 1) {
      
      warning("Duplicated table description(s). Explore to see if this is an issue.")
      
    }
    
    rm(tabcontents2, tabcontents3)
    
  }
  
  # If gridlines are not required they will be turned off
  
  if (gridlines == "No") {
    
    openxlsx::showGridLines(wb, sheetname, showGridLines = FALSE)
    
  }
  
  # Automatically determining column widths
  # Automatic column widths can be hit and miss, so may need to sort these after running the ...
  # ... accessible tables script
  
  if (columnwidths == "R_auto") {
    
    numchars <- max(nchar(as.character(table_data2[[1]]))) + 2
    columns <- colnames(table_data2)
    col1name <- columns[1]
    col1chars <- nchar(col1name) + 2
    
    openxlsx::setColWidths(wb, sheetname, cols = 1, widths = max(numchars, col1chars))
    openxlsx::setColWidths(wb, sheetname, cols = 2:ncol(table_data2), widths = "auto")
    
  } else if (columnwidths == "characters") {
    
    if (is.null(width_adj)) {
      
      width_adj <- 2
      
    }
    
    width_vec <- apply(table_data2, MARGIN = 2, 
                       FUN = function(x) max(nchar(as.character(x)), na.rm = TRUE))
    width_vec <- width_vec + width_adj
    
    width_vec_header <- nchar(colnames(table_data2)) + width_adj
    
    max_vec_header <- pmax(width_vec, width_vec_header)
    
    openxlsx::setColWidths(wb, sheetname, cols = 1:ncol(table_data2), widths = max_vec_header)
    
  } else if (columnwidths == "specified") {
    
    openxlsx::setColWidths(wb, sheetname, cols = 1:ncol(table_data2), widths = colwid_spec)
    
  }
  
  rm(table_data2, envir = as.environment(acctabs))
  
  assign("wb", wb, envir = as.environment(acctabs))
  
}

###################################################################################################