###################################################################################################################
###################################################################################################################
# CREATE WORKBOOK

# The workbook function creates a new workbook with the required metadata worksheets and defines the workbook's font name, colour and sizes
# All parameters are optional and preset
# If a cover page, contents page, notes page or definitions page are required then set the parameter to "Yes" when calling the function
# autonotes is required if a line is wanted towards the top of the worksheet which lists all the note numbers associated with the worksheet (set to "Yes" if wanted)
# Default font is Arial with a black colour and size ranging from 12 to 16 - change if want to when calling the function
# title, creator, subject and category refer to the document information properties displayed in the final Excel workbook


workbook <- function(covertab = NULL, contentstab = NULL, notestab = NULL, autonotes = NULL,
                     definitionstab = NULL, fontnm = "Arial", fontcol = "black",
                     fontsz = 12, fontszst = 14, fontszt = 16, title = NULL, creator = NULL,
                     subject = NULL, category = NULL) {
  
  # Install the required packages if they are not already installed, then load the packages
  
  listofpackages <- base::c("openxlsx", "conflicted", "tidyverse")
  packageversions <- base::c("4.2.5.2", "1.2.0", "2.0.0")
  
  for (i in base::seq_along(listofpackages)) {
    
    if (!(listofpackages[i] %in% utils::installed.packages())) {
      
      utils::install.packages(listofpackages[i], dependencies = TRUE, type = "binary")
     
    } else if (listofpackages[i] %in% utils::installed.packages() & utils::packageVersion(listofpackages[i]) < packageversions[i]) {
      
      base::unloadNamespace(listofpackages[i])
      utils::install.packages(listofpackages[i], dependencies = TRUE, type = "binary")
      
    } 
    
  }
  
  base::library("conflicted")
  
  # When functions are used in this script, the package from which the function comes from is specified e.g., dplyr::filter
  # The exception to this is if the functions come from the R base package
  # To ensure there is no unintentional masking of base functions, conflict_prefer_all will set it so base is the package used unless otherwise specified
  
  conflicted::conflict_prefer_all("base", quiet = TRUE)
  conflicted::conflict_prefer("%>%", "dplyr", quiet = TRUE)
  
  library("tidyverse")
  library("openxlsx")
  
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
  
  # Checking some of the parameters to ensure they are properly populated, if not the function will error
  
  if (length(covertab) > 1 | length(contentstab) > 1 | length(notestab) > 1 | length(autonotes) > 1 | length(definitionstab) > 1) {
    
    stop("One or more of covertab, contentstab, notestab, definitionstab and autnotes not populated with a single word (\"Yes\", \"No\")")
    
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
  
  if (length(fontnm) > 1 | length(fontcol) > 1 | length(fontsz) > 1 | length(fontszst) > 1 | length(fontszt) > 1) {
    
    stop("One or more of fontnm, fontcol, fontsz, fontszst and fontszt is more than a single entity")
    
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
  
  if (exists("xxx_table_data2_xxx", envir = .GlobalEnv)) {
    
    stop("A data frame called xxx_table_data2_xxx exists in the global environment. This needs to be renamed. The code will overwrite any file named xxx_table_data2_xxx.")
    
  }
  
  if (fontsz < 12 | fontszst < 12 | fontszt < 12) {
    
    warning("GSS accessibility guidelines suggest minimum font size should be 12")
    
  }
  
  # If workbook already exists, delete it and create a new, blank one
  
  if (exists("wb", envir = .GlobalEnv)) {
    
    rm(wb, envir = .GlobalEnv)
    warning("wb has been removed from the global environment. If wb is a remnant from a previous run of the table code then it is not a problem. However, if wb is a data frame or variable that you have created then you will need to shut R down and start again but rename whatever you had called wb to something else.")
    
  }
  
  wb <<- openxlsx::createWorkbook(title = title, creator = creator, subject = subject,
                                  category = category)
  
  # If old contents, notes and definitions data frames exist, delete them
  
  if (exists("tabcontents", envir = .GlobalEnv)) {
    
    rm(tabcontents, envir = .GlobalEnv)
    warning("tabcontents has been removed from the global environment. If tabcontents is a remnant from a previous run of the table code then it is not a problem. However, if tabcontents is a data frame or variable that you have created then you will need to shut R down and start again but rename whatever you had called tabcontents to something else.")
    
  }
  
  if (exists("notesdf", envir = .GlobalEnv)) {
    
    rm(notesdf, envir = .GlobalEnv)
    warning("notesdf has been removed from the global environment. If notesdf is a remnant from a previous run of the table code then it is not a problem. However, if notesdf is a data frame or variable that you have created then you will need to shut R down and start again but rename whatever you had called notesdf to something else.")
    
  }
  
  if (exists("definitionsdf", envir = .GlobalEnv)) {
    
    rm(definitionsdf, envir = .GlobalEnv)
    warning("definitionsdf has been removed from the global environment. If definitionsdf is a remnant from a previous run of the table code then it is not a problem. However, if definitionsdf is a data frame or variable that you have created then you will need to shut R down and start again but rename whatever you had called definitionsdf to something else.")
    
  }
  
  if (exists("covernumrow", envir = .GlobalEnv)) {
    
    rm(covernumrow, envir = .GlobalEnv)
    warning("covernumrow has been removed from the global environment. If covernumrow is a remnant from a previous run of the table code then it is not a problem. However, if covernumrow is a data frame or variable that you have created then you will need to shut R down and start again but rename whatever you had called covernumrow to something else.")
    
  }
  
  if (length(ls(pattern = "_startrow", envir = .GlobalEnv)) > 0) {
    
    rm(list = ls(pattern = "_startrow", envir = .GlobalEnv), envir = .GlobalEnv)
    warning("Strings containing \"_startrow\" have been removed from the global environment. If these are remnants from a previous run of the table code then it is not a problem. However, if they were data frames or variables that you have created then you will need to shut R down and start again but rename whatever you had called these objects to something else.")
    
  }
  
  if (length(ls(pattern = "_tablestart", envir = .GlobalEnv)) > 0) {
    
    rm(list = ls(pattern = "_tablestart", envir = .GlobalEnv), envir = .GlobalEnv)
    warning("Strings containing \"_tablestart\" have been removed from the global environment. If these are remnants from a previous run of the table code then it is not a problem. However, if they were data frames or variables that you have created then you will need to shut R down and start again but rename whatever you had called these objects to something else.")
    
  }
  
  # Add required metadata worksheets to the workbook and create new notes and definitions data frames if wanted
  
  if (covertab == "Yes") {
    
    openxlsx::addWorksheet(wb, "Cover")
    
  }
  
  if (contentstab == "Yes") {
    
    openxlsx::addWorksheet(wb, "Contents")
    
  }
  
  if (exists("autonotes2", envir = .GlobalEnv)) {
    
    rm(autonotes2, envir = .GlobalEnv)
    warning("autonotes2 has been removed from the global environment. If autonotes2 is a remnant from a previous run of the table code then it is not a problem. However, if autonotes2 is a data frame or variable that you have created then you will need to shut R down and start again but rename whatever you had called autonotes2 to something else.")
    
  }
  
  if (notestab == "Yes") {
    
    openxlsx::addWorksheet(wb, "Notes")
    
    notesdf <<- data.frame() %>%
      dplyr::mutate("Note number" = "", "Note text" = "", "Applicable tables" = "", "Link1" = "", "Link2" = "")
    
  }
  
  if (definitionstab == "Yes") {
    
    openxlsx::addWorksheet(wb, "Definitions")
    
    definitionsdf <<- data.frame() %>%
      dplyr::mutate("Term" = "", "Definition" = "", "Link1" = "", "Link2" = "")
    
  }
  
  # Create a variable autonotes2 in the global environment for later use
  
  if (autonotes == "Yes") {
    
    autonotes2 <<- "Yes"
    
  } else {
    
    autonotes2 <<- "No"
    
  }
  
  # Create variables for font sizes in the global environment for later use
  
  if (exists("fontsz", envir = .GlobalEnv)) {
    
    rm(fontsz, envir = .GlobalEnv)
    warning("fontsz has been removed from the global environment. If fontsz is a remnant from a previous run of the table code then it is not a problem. However, if fontsz is a data frame or variable that you have created then you will need to shut R down and start again but rename whatever you had called fontsz to something else.")
    
  }
 
  if (exists("fontszst", envir = .GlobalEnv)) {
    
    rm(fontszst, envir = .GlobalEnv)
    warning("fontszst has been removed from the global environment. If fontszst is a remnant from a previous run of the table code then it is not a problem. However, if fontszst is a data frame or variable that you have created then you will need to shut R down and start again but rename whatever you had called fontszst to something else.")
    
  }
  
  if (exists("fontszt", envir = .GlobalEnv)) {
    
    rm(fontszt, envir = .GlobalEnv)
    warning("fontszt has been removed from the global environment. If fontszt is a remnant from a previous run of the table code then it is not a problem. However, if fontszt is a data frame or variable that you have created then you will need to shut R down and start again but rename whatever you had called fontszt to something else.")
    
  }
  
  fontsz <<- fontsz
  fontszst <<- fontszst
  fontszt <<- fontszt
  
  openxlsx::modifyBaseFont(wb, fontSize = fontsz, fontColour = fontcol, fontName = fontnm)
  
}

###################################################################################################################
###################################################################################################################
# MAIN TABLES

# The creatingtables function will create a worksheet with all the data and annotations
# title, sheetname and table_data are the only compulsory parameters
# All other parameters are optional and most are preset to NULL, so only need to be defined if they are wanted
# sheetname is what you want the sheet to be called in the published workbook
# table_data is the name of the R dataframe containing the data to be included in the published workbook
# headrowsize is the height of the row containing the table column names
# numdatacols is the column position(s) of columns containing number data values (character or numeric class) - it is useful for right aligning data columns and inserting thousand commas
# numdatacolsdp is the number of desired decimal places for columns with numbers (character or numeric class)
# othdatacols is the column position(s) of columns containing data values that are not numbers (e.g., text, dates) - it is useful for formatting (although at present it seems dates do not obey the given formatting)
# Character class data columns will have thousand commas inserted as long as the column position is identified in numdatacols
# Numeric class data columns will only have thousand commas inserted if numdatacolsdp is populated
# numdatacolsdp either should be one value which will be applied to all numdatacols columns or a vector the same length as numdatacols
# For character variables, the figure in Excel will only be the value rounded to the specified number of decimal places. For numeric variables, the figure in Excel will be maintained but the displayed figure will be the value rounded to the specified number of decimal places.
# Enter 0 in numdatacolsdp if no decimal places wanted. If an element in numdatacols represents a non-character and non-numeric class column, enter 0 in the corresponding position in numdatacolsdp
# tablename is the name of the table within the worksheet that a screen reader will detect. It is automatically selected to be the same as the sheetname unless tablename is populated.
# gridlines is preset to "Yes", change to "No" if gridlines are not wanted
# columnwidths is preset to "R_auto" which allows openxlsx to automatically determine column widths. If automatic width determination is not wanted, set to NULL.
# columnwidths can alternatively be set to "characters" which will base the column widths on the number of characters in a column cell.
# If columnwidths = "characters" then width_adj can be modified. width_adj adds an additional few spaces to the number of characters in a column cell.
# width_adj can either be one value which will be applied to all columns or a vector the same length as the number of columns in the table
# If you want to specify the exact width of each column, set columnwidths = "specified" and provide the widths in colwid_spec (e.g., colwid_spec = c(3,4,5))
# If a link to the contents page is required, set one of the extralines to "Link to contents"
# If a link to the notes page is required, set one of the extralines to "Link to notes"
# If a link to the definitions page is required, set one of the extralines to "Link to definitions"
# extralines1-6 can be set to hyperlinks - e.g., extraline5 = "[BBC](https://www.bbc.co.uk)"


creatingtables <- function(title, subtitle = NULL, extraline1 = NULL, extraline2 = NULL, extraline3 = NULL,
                           extraline4 = NULL, extraline5 = NULL, extraline6 = NULL, sheetname, table_data, 
                           headrowsize = NULL, numdatacols = NULL, numdatacolsdp = NULL, othdatacols = NULL, 
                           tablename = NULL, gridlines = "Yes", columnwidths = "R_auto", width_adj = NULL,
                           colwid_spec = NULL) {
  
  # Checking some of the parameters to ensure they are properly populated, if not the function will error or display a warning in the console
  
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
    dplyr::summarise(numcellmiss3 = max(numcellmiss2, na.rm = TRUE))
  
  if (table_data_temp2[["numcellmiss3"]] == 1) {
    
    warning("There is at least one row missing data in all columns. Check to make sure there is not a problem. This code is not designed to produce worksheets with more than one table in them.")
    
  }
  
  rm(table_data_temp2)
  
  colsmiss <- colSums(is.na(table_data_temp))
  colsmiss2 <- 0
  
  for (i in seq_along(colsmiss)) {
    
    if (colsmiss[i] == nrow(table_data_temp)) {colsmiss2 <- 1}
    
  }
  
  if (colsmiss2 == 1) {
    
    warning("There is at least one column missing data. Check to make sure there is not a problem. This code is not designed to produce worksheets with more than one table in them.")
    
  }
  
  rm(colsmiss, colsmiss2)
  
  if (any(is.na(table_data_temp)) == TRUE) {
    
    warning("There are some blank cells present in the data. Check to make sure these are not a problem.")
    
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
    
    stop("One or more of title, subtitle, sheetname, headrowsize, tablename and gridlines are not populated properly. They must be a single entity and not a vector.")
    
  }
  
  if (length(extraline1) > 1 | length(extraline2) > 1 | length(extraline3) > 1 | length(extraline4) > 1 |
      length(extraline5) > 1 | length(extraline6) > 1) {
    
    warning("One or more of extraline1, extraline2, extraline3, extraline4, extraline5 and extraline6 is a vector. Check that this is intentional.")
    
  }
  
  if (nchar(sheetname) > 31) {
    
    stop("The number of characters in sheetname must not exceed 31")
    
  }
   
  if (!is.null(sheetname) & is.numeric(sheetname)) {
    
    sheetname <- as.character(sheetname)
    warning("sheetname is numeric and has been changed to character class")
    
  } else if (!is.null(sheetname) & !is.character(sheetname)) {
    
    stop("sheetname must be of character class")
    
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
    
    warning("sheetname is only comprised of numbers - this can cause an issue when opening up the final spreadsheet")
    
  }
  
  if (stringr::str_detect(sheetname, " ")) {
    
    sheetname <- stringr::str_replace(sheetname, " ", "_")
    
  }
  
  if (deparse(substitute(table_data)) == "xxx_table_data2_xxx") {
    
    stop("Data frame used as table_data needs to be renamed. The code will overwrite any file named xxx_table_data2_xxx.")
    
  }
  
  if ("xxx_temp_xxx" %in% colnames(table_data) | "xxx_temp_xxx2" %in% colnames(table_data) | "xxx_temp_xxx3" %in% colnames(table_data))  { 
    
    stop("Temporary variables (xxx_temp_xxx or xxx_temp_xxx2 or xxx_temp_xxx3) already exist on the file")
    
  }
  
  if (!is.null(numdatacols) & !is.numeric(numdatacols)) {
    
    stop("numdatacols either needs to be numeric (e.g., numdatacols = 6 or numdatacols = c(2,5) or NULL")
    
  }
  
  if (!is.null(numdatacols)) {
    
    if (any(numdatacols <= 0)) {
      
      stop("Column positions for numdatacols cannot be 0 or negative")
      
    } else if (any(numdatacols > ncol(table_data))) {
      
      stop("Column positions for numdatacols should not exceed the number of columns in the data")
      
    }
    
  }
  
  if (!is.null(numdatacolsdp) & !is.numeric(numdatacolsdp)) {
    
    stop("numdatacolsdp either needs to be numeric (e.g., numdatacolsdp = 6 or numdatacolsdp = c(2,5) or NULL")
    
  }
  
  if (!is.null(numdatacolsdp)) {
    
    if (any(numdatacolsdp < 0)) {
      
      stop("numdatacolsdp cannot be negative")
      
    }
    
  }
  
  if (is.null(numdatacols) & !is.null(numdatacolsdp)) {
    
    stop("numdatacols has not been populated but numdatacolsdp has. Need to also populate numdatacols or set numdatacolsdp to NULL.")
    
  }
  
  if (!is.null(numdatacols) & !is.null(numdatacolsdp) & length(numdatacols) > 1 & length(numdatacolsdp) == 1) {
    
    numdatacolsdp <- rep(numdatacolsdp, length(numdatacols))
    warning("numdatacols specifies more than one column. numdatacolsdp has only one value and so it has been assumed that this one value represents the number of decimal places required for each column specified by numdatacols.")
  
  } else if (!is.null(numdatacols) & !is.null(numdatacolsdp) & length(numdatacols) != length(numdatacolsdp)) {
    
    stop("The number of elements in numdatacols and numdatacolsdp needs to be the same (e.g., if numdatacols = c(x,y,z) then numdatacolsdp = c(a,b,c)) or numdatacolsdp set to one value to be applied to all columns in numdatacols")
    
  }
  
  if (any(duplicated(numdatacols)) == TRUE) {
    
    stop("There is at least one column number entered multiple times in numdatacols")
    
  }
  
  if (!is.null(othdatacols) & !is.numeric(othdatacols)) {
    
    stop("othdatacols either needs to be numeric (e.g., othdatacols = 6 or othdatacols = c(2,5) or NULL")
    
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
    
    if (!is.null(numdatacols) & class(table_data[[numdatacols[i]]]) != "character" & class(table_data[[numdatacols[i]]]) != "numeric" & class(table_data[[numdatacols[i]]]) != "integer") {
      
      warning("A column identified as a number column is not of class character or numeric. Check that is intentional.")
      othdatacols <- append(othdatacols, numdatacols[i])
      
    } else if (class(table_data[[numdatacols[i]]]) == "character") {
      
      numcharcols <- append(numcharcols, numdatacols[i])
      numcharcolsdp <- append(numcharcolsdp, numdatacolsdp[i])
      
    } else if (class(table_data[[numdatacols[i]]]) == "numeric" | class(table_data[[numdatacols[i]]]) == "integer") {
      
      numericcols <- append(numericcols, numdatacols[i])
      numericcolsdp <- append(numericcolsdp, numdatacolsdp[i])
      
    }
    
  }
  
  if (!is.null(numericcols) & is.null(numericcolsdp)) {
    
    warning("There are data columns of class numeric but the number of decimal places desired has not been specified. This means that thousand commas cannot be inserted automatically. If these commas are desired then consider entering 0 in the appropriate position in numdatacolsdp.")
    
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
    warning("columnwidths should be a single word, not a vector. It will be changed back to the default of \"R_auto\".")
      
  } else if (columnwidths == "none" | columnwidths == "no" | columnwidths == "n" | columnwidths == "default") {
      
    columnwidths <- "Default"
      
  } else if (!is.null(columnwidths) & !is.character(columnwidths)) {
      
    columnwidths <- "R_auto"
    warning("columnwidths should be a character string. It will be changed back to the default of \"R_auto\".")
    
  } else if (columnwidths == "r_auto") {
      
    columnwidths <- "R_auto"
      
  } else if (columnwidths == "character") {
      
    columnwidths <- "characters"
      
  } else if (columnwidths != "r_auto" & columnwidths != "characters" & columnwidths != "specified") {
      
    columnwidths <- "R_auto"
    warning("columnwidths has not been set to \"R_auto\" or \"characters\" or \"specified\" or NULL. It will be changed back to the default of \"R_auto\".")
      
  }
  
  if (!is.null(colwid_spec)) {
    
    if (any(colwid_spec <= 0)) {
      
      stop("Column widths cannot be 0 or negative")
      
    }
      
  }
    
  if (columnwidths == "specified" & is.null(colwid_spec)) {
    
    stop("The option to specify column widths has been selected but the widths have not been provided")
    
  } else if (columnwidths != "specified" & !is.null(colwid_spec)) {
    
    stop("The option to specify column widths has not been selected but the widths have been provided")
    
  } else if (columnwidths == "specified" & length(colwid_spec) == 1 & length(colnames(table_data)) > 1) {
    
    colwid_spec <- rep(colwid_spec, length(colnames(table_data)))
    warning("There is more than one column in the table. colwid_spec has only one value and so it has been assumed that this one value represents the widths of all columns.")
    
  } else if (columnwidths == "specified" & length(colwid_spec) != length(colnames(table_data))) {
    
    stop("The number of elements in colwid_spec and the number of columns in the table need to be the same, or colwid_spec set to one value to be applied to all columns in the table")
    
  }
  
  if (!is.null(width_adj)) {
    
    if (!is.numeric(width_adj)) {
    
      stop("width_adj must be a numeric value")
      
    } else if (length(width_adj) > 1 & length(width_adj) != length(colnames(table_data)) & columnwidths == "characters") {
    
      stop("The number of elements in width_adj is not equal to the number of columns in the table data. The number of elements and columns should either be equal or width_adj should be set to only a single value.")
      
    }
    
  }
  
  # In addition to the title and subtitle, six other fields are permitted above the main data - these extra fields can be provided as vectors and so there is really no limit to the number of rows that can come before the main data
  # If a line with information on notes is wanted, this is initially created and existing rows with information are shifted down one row position
  
  extralines1 <- c(extraline1, extraline2, extraline3, extraline4, extraline5, extraline6)
  
  if ("Notes" %in% names(wb) & autonotes2 == "Yes") {
    
    for (i in seq_along(extralines1)) {
      
      if (stringr::str_detect(extralines1[i], "This worksheet contains one table|this worksheet contains one table|\\[note")) {
        
        warning("If autonotes2 is set to \"Yes\" then the information about the worksheet containing one table or the notes tab will automatically be inserted and so there is no need to have one of the extralines already stating this")
        
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
    assign(paste0(sheetname, "_startrow"), temp, envir = .GlobalEnv)
    rm(temp)
    
  } else if ("Notes" %in% names(wb) & autonotes2 == "No") {
    
    onetablenote <- 0
    notescolumn <- 0
    
    for (i in seq_along(extralines1)) {
      
      if (stringr::str_detect(extralines1[i], "This worksheet contains one table|this worksheet contains one table")) {
        
        onetablenote <- 1
        
      }
      
      if ("Notes" %in% colnames(table_data) | "Note" %in% colnames(table_data) | stringr::str_detect(extralines1[i], "\\[note")) {
        
        notescolumn <- 1
        
      }
      
    }
    
    if (onetablenote == 0) {
      
      warning("There is no recognisable reference to the worksheet containing one table. Consider whether you want to make a reference to this in one of the extra lines above the main data.")
      
    }
    
    if (notescolumn == 0) {
      
      warning("There is no recognisable notes column or reference to notes. Check whether this is OK.")
      
    }
    
    rm(onetablenote, notescolumn)
    
  }
  
  # Remove a specific data frame if it already exists as a new one will be required
  
  if (exists("xxx_table_data2_xxx", envir = .GlobalEnv)) {
    
    rm(xxx_table_data2_xxx, envir = .GlobalEnv)
    
  }
  
  # Function to deal with columns containing numbers stored as text, likely as some cells contain character values (e.g., [c] to indicate some form of statistical disclosure control)
  # The function recognises characters accepted by the GSS as symbols or shorthand applicable for use in tables (b, c, e, er, f, low, p, r, u, w, x, z)
  # Thousand commas will be inserted if necessary (e.g., 1,340)
  # Function will be called only when a specific number of decimal places is not given
  
  xxx_table_data2_xxx <<- table_data
  
  numcharvars <- function(numcharcols) {
    
    dfx <- xxx_table_data2_xxx %>%
      dplyr::mutate(xxx_temp_xxx = dplyr::case_when(.[[numcharcols]] %in% c("[b]", "[c]", "[e]", "[er]", "[f]",
                                                                            "[low]", "[p]", "[r]", "[u]", "[w]",
                                                                            "[x]", "[z]", "") ~ "0",
                                                    is.na(.[[numcharcols]]) ~ "0",
                                                    !is.na(.[[numcharcols]]) ~ gsub(",", "", .[[numcharcols]]))) %>%
      dplyr::mutate(xxx_temp_xxx = as.numeric(xxx_temp_xxx)) %>%
      dplyr::mutate(xxx_temp_xxx2 = format(xxx_temp_xxx, big.mark = ",", scientific = FALSE)) %>%
      dplyr::mutate(xxx_temp_xxx3 = dplyr::case_when(.[[numcharcols]] %in% c("[b]", "[c]", "[e]", "[er]", "[f]",
                                                                             "[low]", "[p]", "[r]", "[u]", "[w]",
                                                                             "[x]", "[z]", "") ~ as.character(.[[numcharcols]]),
                                                     is.na(.[[numcharcols]]) ~ "",
                                                     TRUE ~ as.character(xxx_temp_xxx2)))
    
    dfx[[numcharcols]] <- dfx$xxx_temp_xxx3
    
    xxx_table_data2_xxx <<- dfx %>%
      dplyr::select(-xxx_temp_xxx, -xxx_temp_xxx2, -xxx_temp_xxx3)
    
  }
  
  # Function to deal with columns containing numbers stored as text, likely as some cells contain character values (e.g., [c] to indicate some form of statistical disclosure control)
  # The function recognises characters accepted by the GSS as symbols or shorthand applicable for use in tables (b, c, e, er, f, low, p, r, u, w, x, z)
  # Thousand commas will be inserted if necessary (e.g., 1,340.54)
  # Function will be called only when a specific number of decimal places is given
  
  numcharvars2 <- function(numcharcols, numcharcolsdp) {
    
    dfx <- xxx_table_data2_xxx %>%
      dplyr::mutate(xxx_temp_xxx = dplyr::case_when(.[[numcharcols]] %in% c("[b]", "[c]", "[e]", "[er]", "[f]",
                                                                            "[low]", "[p]", "[r]", "[u]", "[w]",
                                                                            "[x]", "[z]", "") ~ "0",
                                                    is.na(.[[numcharcols]]) ~ "0",
                                                    !is.na(.[[numcharcols]]) ~ gsub(",", "", .[[numcharcols]]))) %>%
      dplyr::mutate(xxx_temp_xxx = if (numcharcolsdp >= 2) as.numeric(xxx_temp_xxx) else round(as.numeric(xxx_temp_xxx), digits = numcharcolsdp)) %>%
      dplyr::mutate(xxx_temp_xxx2 = if (numcharcolsdp >= 2) format(xxx_temp_xxx, big.mark = ",", scientific = FALSE, nsmall = numcharcolsdp) else format(xxx_temp_xxx, big.mark = ",", scientific = FALSE)) %>%
      dplyr::mutate(xxx_temp_xxx3 = dplyr::case_when(.[[numcharcols]] %in% c("[b]", "[c]", "[e]", "[er]", "[f]",
                                                                             "[low]", "[p]", "[r]", "[u]", "[w]",
                                                                             "[x]", "[z]", "") ~ as.character(.[[numcharcols]]),
                                                     is.na(.[[numcharcols]]) ~ "",
                                                     TRUE ~ as.character(xxx_temp_xxx2)))
    
    dfx[[numcharcols]] <- dfx$xxx_temp_xxx3
    
    xxx_table_data2_xxx <<- dfx %>%
      dplyr::select(-xxx_temp_xxx, -xxx_temp_xxx2, -xxx_temp_xxx3)
    
  }
  
  # If there are columns with numbers stored as text then one of the two functions above will be ran
  # Which function depends on whether the numbers stored as text should have a specific number of decimal places or not
  # If there are no columns with numbers stored as text then the data are left alone
  
  if (!is.null(numcharcols) & !is.null(numcharcolsdp)) {
  
    purrr::pmap(list(numcharcols, numcharcolsdp), numcharvars2)
    
  } else if (!is.null(numcharcols) & is.null(numcharcolsdp)) {
    
    purrr::pmap(list(numcharcols), numcharvars)
    
  } else if (is.null(numcharcols)) {
    
    xxx_table_data2_xxx <<- table_data
    
  }
  
  # Add the worksheet to the workbook and define various formatting to be used at some point
  
  openxlsx::addWorksheet(wb, sheetname)
  
  extralines2 <- c(extraline1, extraline2, extraline3, extraline4, extraline5, extraline6)
  
  tablestart <- (length(title) + length(subtitle) + length(extralines2) + 1)
  assign(paste0(sheetname, "_tablestart"), tablestart, envir = .GlobalEnv)
  
  titleformat <- openxlsx::createStyle(fontSize = fontszt, textDecoration = "bold")
  subtitleformat <- openxlsx::createStyle(fontSize = fontszst)
  normalformat <- openxlsx::createStyle(valign = "top")
  linkformat <- openxlsx::createStyle(fontColour = "blue", textDecoration = "underline")
  topformat <- openxlsx::createStyle(valign = "bottom")
  headingsformat <- openxlsx::createStyle(textDecoration = "bold", wrapText = TRUE, border = NULL, valign = "top")
  headingsformat2 <- openxlsx::createStyle(textDecoration = "bold", wrapText = TRUE, border = NULL, valign = "top", halign = "right")
  dataformat <- openxlsx::createStyle(halign = "right", valign = "top")
  
  openxlsx::addStyle(wb, sheetname, normalformat, rows = 1:(nrow(xxx_table_data2_xxx) + tablestart), cols = 1:ncol(xxx_table_data2_xxx), gridExpand = TRUE)
  openxlsx::addStyle(wb, sheetname, topformat, rows = 1:(length(title) + length(subtitle) + length(extralines2)), cols = 1, gridExpand = TRUE)
  
  openxlsx::writeData(wb, sheetname, title, startCol = 1, startRow = 1)
  
  openxlsx::addStyle(wb, sheetname, titleformat, rows = 1, cols = 1)
  
  if (!is.null(subtitle)) {
    
    openxlsx::writeData(wb, sheetname, subtitle, startCol = 1, startRow = 2)
    openxlsx::addStyle(wb, sheetname, subtitleformat, rows = 2, cols = 1)
    
  }
  
  # If a link is wanted to the contents or notes or definitions page then the code below will create the hyperlink
  
  for (i in seq_along(extralines2)) {
    
    if (tolower(extralines2[i]) == "link to notes" | tolower(extralines2[i]) == "notes") {
      
      extralines2[i] <- "Link to notes"
      
    }
    
    if (extralines2[i] == "Link to notes" & !("Notes" %in% names(wb))) {
      
      stop("Cannot put a link in to the notes tab unless notestab set to \"Yes\" in the workbook function call")
      
    }
    
    if (tolower(extralines2[i]) == "link to contents" | tolower(extralines2[i]) == "contents") {
      
      extralines2[i] <- "Link to contents"
      
    }
    
    if (extralines2[i] == "Link to contents" & !("Contents" %in% names(wb))) {
      
      stop("Cannot put a link in to the contents tab unless contentstab set to \"Yes\" in the workbook function call")
      
    }
    
    if (tolower(extralines2[i]) == "link to definitions" | tolower(extralines2[i]) == "definitions") {
      
      extralines2[i] <- "Link to definitions"
      
    }
    
    if (extralines2[i] == "Link to definitions" & !("Definitions" %in% names(wb))) {
      
      stop("Cannot put a link in to the definitions tab unless definitionstab set to \"Yes\" in the workbook function call")
      
    }
    
    if (extralines2[i] == "Link to notes") {
      
      openxlsx::writeFormula(wb, sheetname, startRow = length(title) + length(subtitle) + i, x = openxlsx::makeHyperlinkString("Notes", row = 1, col = 1, text = "Link to notes"))
      openxlsx::addStyle(wb, sheetname, linkformat, rows = length(title) + length(subtitle) + i, cols = 1, stack = TRUE)
      
    } else if (extralines2[i] == "Link to contents") {
      
      openxlsx::writeFormula(wb, sheetname, startRow = length(title) + length(subtitle) + i, x = openxlsx::makeHyperlinkString("Contents", row = 1, col = 1, text = "Link to contents"))
      openxlsx::addStyle(wb, sheetname, linkformat, rows = length(title) + length(subtitle) + i, cols = 1, stack = TRUE)
      
    } else if (extralines2[i] == "Link to definitions") {
      
      openxlsx::writeFormula(wb, sheetname, startRow = length(title) + length(subtitle) + i, x = openxlsx::makeHyperlinkString("Definitions", row = 1, col = 1, text = "Link to definitions"))
      openxlsx::addStyle(wb, sheetname, linkformat, rows = length(title) + length(subtitle) + i, cols = 1, stack = TRUE)
      
    } else if (!is.null(extralines2[i])) {
      
      hyper_rx <- "\\[(([[:graph:]]|[[:space:]])+)\\]\\([[:graph:]]+\\)"
      
      if (grepl(hyper_rx, extralines2[i]) == TRUE) {
        
        if (substr(extralines2[i], 1, 1) != "[" | substr(extralines2[i], nchar(extralines2[i]), nchar(extralines2[i])) != ")") {
          
          warning(paste0(extralines2[i], " - if this is meant to be a hyperlink, it needs to be in the format \"[xxx](xxxxxx)\""))
          
        }
        
        if ("Link to contents" %in% extralines2 & stringr::str_detect(tolower(extralines2[i]), "\\[link to contents|\\[contents")) {
          
          warning(paste0(extralines2[i], " - this appears to be duplicating a link to the contents page in another extraline parameter"))
          
        } else if ("Link to notes" %in% extralines2 & stringr::str_detect(tolower(extralines2[i]), "\\[link to notes|\\[notes")) {
          
          warning(paste0(extralines2[i], " - this appears to be duplicating a link to the notes page in another extraline parameter"))
          
        } else if ("Link to definitions" %in% extralines2 & stringr::str_detect(tolower(extralines2[i]), "\\[link to definitions|\\[definitions")) {
          
          warning(paste0(extralines2[i], " - this appears to be duplicating a link to the definitions page in another extraline parameter"))
          
        }
        
        if (stringr::str_detect(tolower(extralines2[i]), "\\[link to contents|\\[link to notes|\\[link to definitions|\\[contents|\\[notes|\\[definitions")) {
          
          warning("If you want an internal link to the contents, notes or definitions page, then set one of extraline1-6 to \"Link to contents\" or \"Link to notes\" or \"Link to definitions\"")
          
        }
        
        x <- extralines2[i]
        
        md_rx <- "\\[(([[:graph:]]|[[:space:]])+?)\\]\\([[:graph:]]+?\\)"
        md_match <- regexpr(md_rx, x, perl = TRUE)
        md_extract <- regmatches(x, md_match)[[1]]
        
        url_rx <- "(?<=\\]\\()([[:graph:]]|[[:space:]])+(?=\\))"
        url_match <- regexpr(url_rx, md_extract, perl = TRUE)
        url_extract <- regmatches(md_extract, url_match)[[1]]
        
        string_rx <- "(?<=\\[)([[:graph:]]|[[:space:]])+(?=\\])"
        string_match <- regexpr(string_rx, md_extract, perl = TRUE)
        string_extract <- regmatches(md_extract, string_match)[[1]]
        
        string_extract <- gsub(md_rx, string_extract, x)
        
        y <- stats::setNames(url_extract, string_extract)
        class(y) <- "hyperlink"
        
        rm(x, md_rx, md_match, md_extract, url_rx, url_match, url_extract, string_rx, string_match, string_extract)
        
      } else {
        
        y <- extralines2[i]
        
      }
      
      openxlsx::writeData(wb, sheetname, y, startCol = 1, startRow = length(title) + length(subtitle) + i)
      
      if (grepl(hyper_rx, extralines2[i]) == TRUE) {
        
        openxlsx::addStyle(wb, sheetname, linkformat, rows = length(title) + length(subtitle) + i, cols = 1, stack = TRUE)
        
      }
      
      rm(y)
     
    }
    
  }
  
  openxlsx::addStyle(wb, sheetname, normalformat, rows = tablestart - 1, cols = 1, stack = TRUE)
  
  openxlsx::addStyle(wb, sheetname, headingsformat, rows = tablestart, cols = 1:ncol(xxx_table_data2_xxx))
  
  # Applying specific formatting to data columns
  
  if (!is.null(numericcols)) {
    
    openxlsx::addStyle(wb, sheetname, headingsformat2, rows = tablestart, cols = numericcols)
    openxlsx::addStyle(wb, sheetname, dataformat, rows = (tablestart + 1):(nrow(xxx_table_data2_xxx) + tablestart + 1), cols = numericcols, stack = TRUE, gridExpand = TRUE)
    
  }
  
  if (!is.null(numcharcols)) {
    
    openxlsx::addStyle(wb, sheetname, headingsformat2, rows = tablestart, cols = numcharcols)
    openxlsx::addStyle(wb, sheetname, dataformat, rows = (tablestart + 1):(nrow(xxx_table_data2_xxx) + tablestart + 1), cols = numcharcols, stack = TRUE, gridExpand = TRUE)
    
  }
  
  if (!is.null(othdatacols)) {
    
    openxlsx::addStyle(wb, sheetname, headingsformat2, rows = tablestart, cols = othdatacols)
    openxlsx::addStyle(wb, sheetname, dataformat, rows = (tablestart + 1):(nrow(xxx_table_data2_xxx) + tablestart + 1), cols = othdatacols, stack = TRUE, gridExpand = TRUE)
    
  }
  
  # If a specific number of decimal places is wanted for numeric columns, the code below will do this as well as inserting thousand commas
  
  if (!is.null(numericcolsdp)) {
    
    for (i in seq_along(numericcolsdp)) {
    
      if (numericcolsdp[i] > 0) {
      
        fmta <- paste0("#,##0.", strrep("0", numericcolsdp[i]))
        fmt <- openxlsx::createStyle(numFmt = fmta)
        openxlsx::addStyle(wb, sheetname, fmt, rows = (tablestart + 1):(nrow(xxx_table_data2_xxx) + tablestart + 1), cols = numericcols[i], stack = TRUE, gridExpand = TRUE)
        rm(fmta, fmt)
        
      } else if (numericcolsdp[i] == 0) {
      
        fmt <- openxlsx::createStyle(numFmt = "#,##0")
        openxlsx::addStyle(wb, sheetname, fmt, rows = (tablestart + 1):(nrow(xxx_table_data2_xxx) + tablestart + 1), cols = numericcols[i], stack = TRUE, gridExpand = TRUE)
        rm(fmt)
        
      }
      
    }
    
  } 
  
  # Ensure table cell text is wrapped
  
  wrapformat <- openxlsx::createStyle(wrapText = TRUE)
  
  openxlsx::addStyle(wb, sheetname, wrapformat, rows = (tablestart + 1):(nrow(xxx_table_data2_xxx) + tablestart + 1), cols = 1:ncol(xxx_table_data2_xxx), stack = TRUE, gridExpand = TRUE)
  
  # tablename2 will be the name of the table accessible in Excel
  # If no specific name is given, then the name of the table will be the same as the sheetname
  
  if (!is.null(tablename) & is.character(tablename) & length(tablename) == 1) {
    
    tablename2 <- tablename
    
  } else {
    
    tablename2 <- sheetname
    
  }
  
  # Setting some specific row heights based in part on the font size
  
  openxlsx::writeDataTable(wb, sheetname, xxx_table_data2_xxx, tableName = tablename2, startRow = tablestart, startCol = 1, withFilter = FALSE, tableStyle = "none")
  
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
  
  if (sheetname != "Contents" & exists("tabcontents", envir = .GlobalEnv)) {
    
    tabcontents <<- tabcontents %>%
      dplyr::add_row("Sheet name" = sheetname, "Table description" = title)
      
  } else if (sheetname != "Contents" & !exists("tabcontents", envir = .GlobalEnv)) {
    
    tabcontents <<- data.frame() %>%
      dplyr::mutate("Sheet name" = "", "Table description" = "") %>%
      dplyr::add_row("Sheet name" = sheetname, "Table description" = title)
    
  }
  
  if (exists("tabcontents", envir = .GlobalEnv)) {
  
    tabcontents2 <- tabcontents %>%
      dplyr::rename(sheet_name = "Sheet name") %>%
      dplyr::group_by(sheet_name) %>%
      dplyr::summarise(count = n()) %>%
      dplyr::ungroup() %>%
      dplyr::summarise(check = sum(count) / n())
    
    tabcontents3 <- tabcontents %>%
      dplyr::rename(table_description = "Table description") %>%
      dplyr::group_by(table_description) %>%
      dplyr::summarise(count = n()) %>%
      dplyr::ungroup() %>%
      dplyr::summarise(check = sum(count) / n())
    
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
  # Automatic column widths can be hit and miss, so may need to sort these after running the accessible tables script
  
  if (columnwidths == "R_auto") {
    
    numchars <- max(nchar(as.character(xxx_table_data2_xxx[[1]]))) + 2
    columns <- colnames(xxx_table_data2_xxx)
    col1name <- columns[1]
    col1chars <- nchar(col1name) + 2
    
    openxlsx::setColWidths(wb, sheetname, cols = 1, widths = max(numchars, col1chars))
    openxlsx::setColWidths(wb, sheetname, cols = 2:ncol(xxx_table_data2_xxx), widths = "auto")
    
  } else if (columnwidths == "characters") {
    
    if (is.null(width_adj)) {
      
      width_adj <- 2
      
    }
    
    width_vec <- apply(xxx_table_data2_xxx, MARGIN = 2, FUN = function(x) max(nchar(as.character(x)), na.rm = TRUE))
    width_vec <- width_vec + width_adj
    
    width_vec_header <- nchar(colnames(xxx_table_data2_xxx)) + width_adj
    
    max_vec_header <- pmax(width_vec, width_vec_header)
    
    openxlsx::setColWidths(wb, sheetname, cols = 1:ncol(xxx_table_data2_xxx), widths = max_vec_header)
    
  } else if (columnwidths == "specified") {
    
    openxlsx::setColWidths(wb, sheetname, cols = 1:ncol(xxx_table_data2_xxx), widths = colwid_spec)
    
  }
  
  rm(xxx_table_data2_xxx, envir = .GlobalEnv)
  
}

###################################################################################################################
###################################################################################################################
# CONTENTS

# contentstable function creates a table of contents for the workbook
# If no contents page wanted, then do not run the contentstable function
# gridlines is by default set to "Yes", change to "No" if gridlines are not wanted
# Column widths are automatically set unless user defines specific values in colwid_spec
# Extra columns can be added, need to set extracols to "Yes" and create a dataframe extracols_contents with the desired extra columns 


contentstable <- function(gridlines = "Yes", colwid_spec = NULL, extracols = NULL) {
  
  # Check to see that a contents page is wanted, based on whether a worksheet was created in the initial workbook
  
  if (!("Contents" %in% names(wb))) {
    
    stop("contentstab cannot have been set to \"Yes\" in the workbook function call")
    
  }
  
  # Checking some of the parameters to ensure they are properly populated, if not the function will error or display a warning in the console
  
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
    
    stop("gridlines has not been populated properly. It must be a single word, either \"Yes\" or \"No\".")
    
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
    
    stop("extracols has not been populated properly. It must be a single word, either \"Yes\" or \"No\".")
    
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
  
  # To insert additional columns which are not default columns allowed by the function, a dataframe called "extracols_contents" needs to be created with the extra columns
  
  if (extracols == "Yes" & exists("extracols_contents", envir = .GlobalEnv)) {
    
    if ((nrow(tabcontents) + nrow(notesdf2a) + nrow(notesdf2b)) != nrow(extracols_contents)) {
      
      stop("The number of rows in the table of contents is not the same as in the dataframe of extra columns")
      
    }
    
    if ("Sheet name" %in% colnames(extracols_contents) | "Table description" %in% c(extracols_contents)) {
      
      warning("There is at least one duplicate column name in the contents table and the extracols_contents dataframe")
      
    }
    
    tabcontents <<- dplyr::bind_rows(notesdf2a, notesdf2b, tabcontents) %>%
      dplyr::bind_cols(extracols_contents)
    
  } else if (!(exists("extracols_contents", envir = .GlobalEnv))) {
    
    tabcontents <<- dplyr::bind_rows(notesdf2a, notesdf2b, tabcontents)
    
    if (extracols == "Yes") {
      
      warning("extracols has been set to \"Yes\" but the dataframe extracols_contents does not exist. No extra columns will be added.")
      
    }
    
  } else if (extracols == "No" & exists("extracols_contents", envir = .GlobalEnv)) {
    
    warning("extracols has been set to \"No\" but a dataframe extracols_contents exist. Check if extra columns are wanted. No extra columns have been added.")
    
  }
  
  tabcontents2 <- tabcontents %>%
    dplyr::rename(sheet_name = "Sheet name") %>%
    dplyr::group_by(sheet_name) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(count) / n())
  
  tabcontents3 <- tabcontents %>%
    dplyr::rename(table_description = "Table description") %>%
    dplyr::group_by(table_description) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(count) / n())
  
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
  linkformat <- openxlsx::createStyle(fontColour = "blue", wrapText = TRUE, valign = "top", textDecoration = "underline")
  headingsformat <- openxlsx::createStyle(textDecoration = "bold", wrapText = TRUE, border = NULL, valign = "top")
  
  openxlsx::addStyle(wb, "Contents", normalformat, rows = 1:(nrow(tabcontents) + 3), cols = 1:ncol(tabcontents), gridExpand = TRUE)
  
  openxlsx::writeData(wb, "Contents", title, startCol = 1, startRow = 1)
  
  openxlsx::addStyle(wb, "Contents", titleformat, rows = 1, cols = 1)
  
  openxlsx::writeData(wb, "Contents", extraline1, startCol = 1, startRow = 2)
  
  openxlsx::addStyle(wb, "Contents", extralineformat, rows = 2, cols = 1)
  
  openxlsx::addStyle(wb, "Contents", headingsformat, rows = 3, cols = 1:ncol(tabcontents))
  
  openxlsx::writeDataTable(wb, "Contents", tabcontents, tableName = "contents_table", startRow = 3, startCol = 1, withFilter = FALSE, tableStyle = "none")
  
  numchars <- max(nchar(tabcontents$"Sheet name"))
  
  if (is.null(colwid_spec) & ncol(tabcontents) == 2) {
    
    openxlsx::setColWidths(wb, "Contents", cols = c(1,2), widths = c(max(15, numchars + 3), 100))
    
  } else if (is.null(colwid_spec) & ncol(tabcontents) > 2) {
    
    openxlsx::setColWidths(wb, "Contents", cols = c(1,2,3:ncol(tabcontents)), widths = c(max(15, numchars + 3), 100, "auto"))
    
  } else if (!is.numeric(colwid_spec) | length(colwid_spec) != ncol(tabcontents)) {
    
    warning("colwid_spec has either been provided as non-numeric or a vector of length not equal to the number of columns in tabcontents. The default column widths have been used instead.") ######################################
    openxlsx::setColWidths(wb, "Contents", cols = c(1,2,3:max(ncol(tabcontents),3)), widths = c(max(15, numchars + 3), 100, "auto"))
    
  } else if (!is.null(colwid_spec) & is.numeric(colwid_spec) & length(colwid_spec) == ncol(tabcontents)) {  
    
    openxlsx::setColWidths(wb, "Contents", cols = 1:ncol(tabcontents), widths = colwid_spec)            
    
  }
  
  openxlsx::setRowHeights(wb, "Contents", 2, fontsz * (25/12))
  
  contentrows <- nrow(tabcontents)
  
  # Creating hyperlinks so user can quickly navigate through the spreadsheet
  
  for (i in c(4:(3 + contentrows))) {
    
    openxlsx::writeFormula(wb, "Contents", startRow = i, x = openxlsx::makeHyperlinkString(paste0(tabcontents[i-3, 1]), row = 1, col = 1, text = paste0(tabcontents[i-3, 1])))
    openxlsx::addStyle(wb, "Contents", linkformat, rows = i, cols = 1)
    
  }
  
  # Remove gridlines if they are not wanted
  
  if (gridlines == "No") {
    
    openxlsx::showGridLines(wb, "Contents", showGridLines = FALSE)
    
  }
  
}  

###################################################################################################################
###################################################################################################################
# COVER

# coverpage function will create a cover page for the front of the workbook
# If a cover page is not wanted, do not run the coverpage function
# The only compulsory parameter is title
# All other parameters are optional and preset, only populate if they are wanted
# intro: Introductory information / about: About these data / dop: Date of publication
# source: Data source(s) used / blank: Information about why some cells are blank, if necessary
# relatedlink and relatedtext - any publications associated with the data (relatedlink is the actual hyperlink, relatedtext is the text you want to appear to the user)
# names: Contact name / email: Contact email / phone: Contact telephone
# reuse: Set to "Yes" if you want the information displayed about the reuse of the data (will automatically be populated)
# govdept: Default is "ONS" but if want reuse information without reference to ONS change govdept
# extrafields: Any additional fields that the user wants present on the cover page
# extrafieldsb: The text to go in any additional fields. Only one row per field. extrafields and extrafields must be vectors of the same length.
# additlinks: Any additional hyperlinks the user wants
# addittext: The text to appear over any additional hyperlinks. additlinks and addittext must be vectors of the same length.
# order: If the user wants the cover page to be ordered in a specific way, list the fields in a vector with each field name in speech marks
# e.g., order = c("intro", "about", relatedlink", "names", "phone", "email", "extrafields")
# Change gridlines to "No" if gridlines are not wanted
# Column width automatically set unless user specifies a value in colwid_spec
# intro, about, source, dop, blank, names, phone can be set to hyperlinks - e.g., source = "[ONS](https://www.ons.gov.uk)"


coverpage <- function(title, intro = NULL, about = NULL, source = NULL, relatedlink = NULL, relatedtext = NULL,
                      dop = NULL, blank = NULL, names = NULL, email = NULL, phone = NULL, reuse = NULL,
                      gridlines = "Yes", govdept = "ONS", extrafields = NULL, extrafieldsb = NULL,
                      additlinks = NULL, addittext = NULL, colwid_spec = NULL, order = NULL) {
  
  # Check to see that a coverpage is wanted, based on whether a worksheet was created in the initial workbook
  
  if (!("Cover" %in% names(wb))) {
    
    stop("covertab has not been set to \"Yes\" in the workbook function call")
    
  }
  
  # Checking some parameters have been populated properly, the function will error if not
  
  if (is.null(title)) {
    
    stop("No title entered. Must have a title.")
    
  } else if (title == "") {
    
    stop("No title entered. Must have a title.")
    
  }
  
  if (length(title) > 1 | length(intro) > 1 | length(about) > 1 | length(source) > 1 | length(dop) > 1 |
      length(blank) > 1 | length(names) > 1 | length(email) > 1 | length(phone) > 1 | length(reuse) > 1 |
      length(govdept) > 1) {
    
    stop("One of title, intro, about, source, dop, blank, names, email, phone, reuse and govdept is more than a single entity")
    
  }
  
  if (!is.null(relatedlink) & is.null(relatedtext)) {
    
    stop("relatedlink and relatedtext either have to be both set to NULL or both set to something")
    
  } else if (is.null(relatedlink) & !is.null(relatedtext)) {
    
    stop("relatedlink and relatedtext either have to be both set to NULL or both set to something")
    
  } else if (length(relatedlink) != length(relatedtext)) {
    
    stop("relatedlink and relatedtext must be of the same length and contain the same number of elements")
    
  }
  
  if (is.null(extrafields) & !is.null(extrafieldsb)) {
    
    stop("extrafields and extrafieldsb either have to be both set to NULL or both set to something")
    
  } else if (!is.null(extrafields) & is.null(extrafieldsb)) {
    
    stop("extrafields and extrafieldsb either have to be both set to NULL or both set to something")
    
  } else if (length(extrafields) != length(extrafieldsb)) {
    
    stop("extrafields and extrafieldsb must be of the same length and contain the same number of elements")
    
  }
  
  if (is.null(additlinks) & !is.null(addittext)) {
    
    stop("additlinks and addittext either have to be both set to NULL or both set to something")
    
  } else if (!is.null(additlinks) & is.null(addittext)) {
    
    stop("additlinks and addittext either have to be both set to NULL or both set to something")
    
  } else if (length(additlinks) != length(addittext)) {
    
    stop("additlinks and addittext must be of the same length and contain the same number of elements")
    
  }
  
  if (any(duplicated(order)) == TRUE) {
    
    stop("There is at least one element entered multiple times in \"order\"")
    
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
  
  if (length(gridlines) > 1) {
    
    stop("gridlines has not been populated properly. It must be a single word, either \"Yes\" or \"No\".")
    
  }
  
  if (is.null(reuse)) {
    
    reuse <- "No"
    
  } else if (tolower(reuse) == "yes" | tolower(reuse) == "y") {
    
    reuse <- "Yes"
    
  } else if (tolower(reuse) == "no" | tolower(reuse) == "n") {
    
    reuse <- "No"
    
  }
  
  if (reuse != "No" & reuse != "Yes") {
    
    stop("reuse not set to \"Yes\" or \"No\"")
    
  }
  
  if (stringr::str_remove_all(phone, "[\" \"\\[\\]\\(\\)+[:digit:]]") != "") {
    
    warning("The phone number provided appears to contain characters which are unusual for a phone number. Check if there are any errors.")
    
  }
  
  if (grepl("\\.", email) == FALSE | grepl("@", email) == FALSE) {
    
    warning("The email address provided does not appear to contain @ and/or a dot (.). Check if there are any errors.")
    
  }
  
  # In case the function is run multiple times, removing previous row heights to ensure there will be no strange looking rows
  
  if (exists("covernumrow", envir = .GlobalEnv)) {
    
    for (i in seq_along(1:covernumrow)) {
      
      openxlsx::writeData(wb, "Cover", "", startCol = 1, startRow = i)
      
    }
    
    openxlsx::removeRowHeights(wb, "Cover", 1:covernumrow)
    
  }
  
  # Determining the number of rows that will be populated
  
  if (!is.null(intro)) {intro2 <- 1} else if (is.null(intro)) {intro2 <- 0}
  if (!is.null(about)) {about2 <- 1} else if (is.null(about)) {about2 <- 0}
  if (!is.null(source)) {source2 <- 1} else if (is.null(source)) {source2 <- 0}
  if (!is.null(relatedlink)) {related2 <- 1} else if (is.null(relatedlink)) {related2 <- 0}
  if (!is.null(dop)) {dop2 <- 1} else if (is.null(dop)) {dop2 <- 0}
  if (!is.null(blank)) {blank2 <- 1} else if (is.null(blank)) {blank2 <- 0}
  if (!is.null(additlinks)) {additlinks2 <- 1} else if (is.null(additlinks)) {additlinks2 <- 0}
  if (!is.null(names)) {names2 <- 1} else if (is.null(names)) {names2 <- 0}
  
  covernumrow <<- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + length(blank) + blank2 + length(extrafields) + length(extrafieldsb) + additlinks2 + length(additlinks) + length(names) + names2 + length(email) + length(phone) + length(reuse) + 4
  
  # Populating the cover page with the required text
  
  openxlsx::writeData(wb, "Cover", title, startCol = 1, startRow = 1)
  
  # Beginning of a long process to order the cover page as specified by the order argument
  
  introstartpos <- NULL
  aboutstartpos <- NULL
  sourcestartpos <- NULL
  relatedstartpos <- NULL
  dopstartpos <- NULL
  blankstartpos <- NULL
  namesstartpos <- NULL
  emailstartpos <- NULL
  phonestartpos <- NULL
  reusestartpos <- NULL
  extrastartpos <- NULL
  additstartpos <- NULL
  
  if (!is.null(order)) {
    
    fields <- c(title, intro, about, source, relatedlink, dop, blank, names, email, phone, reuse, extrafields, additlinks)
    
    for (i in seq_along(order)) {
      
      if (tolower(order[i]) %in% c("intro", "introduction", "introductory information")) {order[i] <- intro}
      else if (tolower(order[i]) %in% c("about", "about these data")) {order[i] <- about}
      else if (tolower(order[i]) %in% c("source", "source of data", "data source", "sources", "sources of data", "data sources")) {order[i] <- source}
      else if (tolower(order[i]) %in% c("related publications", "related publication", "related", "relatedlink", "relatedlinks", "relatedtext")) {order[i] <- "relatedlink"}
      else if (tolower(order[i]) %in% c("dop", "date of publication", "publication date")) {order[i] <- dop}
      else if (tolower(order[i]) %in% c("blank", "blank cells")) {order[i] <- blank}
      else if (tolower(order[i]) %in% c("names", "name", "contact", "contact details")) {order[i] <- names}
      else if (tolower(order[i]) %in% c("email", "email address", "e-mail", "e-mail address")) {order[i] <- email}
      else if (tolower(order[i]) %in% c("phone", "telephone", "phone number", "telephone number", "tel", "tel:")) {order[i] <- phone}
      else if (tolower(order[i]) %in% c("reuse", "reusing this publication", "reuse this publication")) {order[i] <- reuse}
      else if (tolower(order[i]) %in% c("extrafields", "extrafield", "extrafieldsb", "extrafieldb")) {order[i] <- "extrafields"}
      else if (tolower(order[i]) %in% c("additlinks", "additlink", "addittext", "additional links", "additional link")) {order[i] <- "additlinks"}
      
    }
    
    phone2 <- which(order == phone)
    email2 <- which(order == email)
    names2 <- which(order == names)
    
    if (names %in% order & phone %in% order & email %in% order) {
      
      if ((phone2 < names2) | (email2 < names2) | (phone2 > (names2 + 2)) | (email2 > (names2 + 2))) {
        
        stop("The relative positions of names, phone and email are not consistent with the expected stucture (i.e., names, phone or email, email or phone)")
        
      }
      
    } else if (names %in% order & phone %in% order) {
      
      if ((phone2 < names2) | (phone2 > (names2 + 1))) {
        
        stop("The relative positions of names and phone are not consistent with the expected structure (i.e., names, phone)")
        
      }
      
    } else if (names %in% order & email %in% order) {
      
      if ((email2 < names2) | (email2 > (names2 + 1))) {
        
        stop("The relative positions of names and email are not consistent with the expected structure (i.e., names, email)")
        
      }
      
    } else if (((email %in% order) | (phone %in% order)) & !(names %in% order)) {
      
      warning("email and/or phone have been populated but a contact name has not been provided. Check that this is intentional.")
      
    }
    
    if ("extrafields" %in% order) {
      
      x <- which(order == "extrafields")
      
      orderb <- order[1:(x-1)]
      
      if ((x + 1) <= length(order)) {
        
        orderc <- order[(x+1):length(order)]
        
      } else {orderc <- NULL}
      
      if (!is.null(orderc)) {
        
        orderd <- c(orderb, extrafields, orderc)
        
      } else if (is.null(orderc)) {
        
        orderd <- c(orderb, extrafields)
        
      }
      
      rm(x, orderb, orderc)
      
    } else if (!("extrafields" %in% order)) {
      
      orderd <- order
      
    }
    
    if ("relatedlink" %in% order) {
      
      x <- which(orderd == "relatedlink")
      
      ordere <- orderd[1:(x-1)]
      
      if ((x + 1) <= length(orderd)) {
        
        orderf <- orderd[(x+1):length(orderd)]
        
      } else {orderf <- NULL}
      
      if (!is.null(orderf)) {
        
        orderg <- c(ordere, relatedlink, orderf)
        
      } else if (is.null(orderf)) {
        
        orderg <- c(ordere, relatedlink)
        
      }
      
      rm(x, ordere, orderf)
      
    } else if (!("relatedlink" %in% order)) {
      
      orderg <- orderd
      
    }
    
    if ("additlinks" %in% order) {
      
      x <- which(orderg == "additlinks")
      
      orderh <- orderg[1:(x-1)]
      
      if ((x + 1) <= length(orderg)) {
        
        orderi <- orderg[(x+1):length(orderg)]
        
      } else {orderi <- NULL}
      
      if (!is.null(orderi)) {
        
        orderj <- c(orderh, additlinks, orderi)
        
      } else if (is.null(orderi)) {
        
        orderj <- c(orderh, additlinks)
        
      }
      
      rm(x, orderh, orderi)
      
    } else if (!("additlinks" %in% order)) {
      
      orderj <- orderg
      
    }
    
    order <- orderj
    
    rm(orderd, orderg, orderj)
    
    if (length(setdiff(order, fields)) > 0) {
      
      x <- paste(setdiff(order, fields), collapse = "  ")
      
      stop(paste0(x, "  -  these are not recognisable field names"))
      
      rm(x)
      
    }
    
    if (length(order) > length(fields)) {
      
      stop("There are more elements specified in \"order\" than are permissible")
      
    }
    
    if (any(duplicated(order)) == TRUE) {
      
      stop("There is at least one element entered multiple times in \"order\"")
      
    }
    
    order <- order[order != title]
    
    orderk <- c(1:length(order))
    orderl <- NULL
    
    for (i in seq_along(orderk)) {
      
      if (i == 1) {
        
        orderl[i] <- orderk[i]
        
      } else {
        
        orderl[i] <- orderk[i] + orderk[i-1]
        
      }
      
    }
    
    if (!is.null(reuse) & reuse %in% order) {
      
      reusepos <- which(order == reuse)
      
      if ((reusepos + 1) <= length(order)) {
        
        for (i in (reusepos + 1):length(order)) {
          
          orderl[i] <- orderl[i] + 3
          
        }
        
      }
      
    }
    
    if (!is.null(relatedlink) & any(is.element(relatedlink, order)) == TRUE) {
      
      relatedpos <- which(order %in% relatedlink)
      
      if ((utils::tail(relatedpos, 1) + 1) <= length(order)) {
        
        for (i in (utils::tail(relatedpos, 1) + 1):length(order)) {
          
          orderl[i] <- orderl[i] - length(relatedlink) + 1
          
        }
        
      }
      
    }
    
    if (!is.null(additlinks) & any(is.element(additlinks, order)) == TRUE) {
      
      additpos <- which(order %in% additlinks)
      
      if ((utils::tail(additpos, 1) + 1) <= length(order)) {
        
        for (i in (utils::tail(additpos, 1) + 1):length(order)) {
          
          orderl[i] <- orderl[i] - length(additlinks) + 1
          
        }
        
      }
      
    }
    
    if (!is.null(email)) {
      
      emailpos <- which(order == email)
      
      if ((emailpos + 1) <= length(order)) {
        
        for (i in (emailpos + 1):length(order)) {
          
          orderl[i] <- orderl[i] - 1
          
        }
        
      }
      
    }
    
    if (!is.null(phone)) {
      
      phonepos <- which(order == phone)
      
      if ((phonepos + 1) <= length(order)) {
        
        for (i in (phonepos + 1):length(order)) {
          
          orderl[i] <- orderl[i] - 1
          
        }
        
      }
      
    }
    
    for (i in seq_along(order)) {
      
      if (order[i] == intro) {introstartpos <- orderl[i] + length(title)}
      else if (order[i] == about) {aboutstartpos <- orderl[i] + length(title)}
      else if (order[i] == source) {sourcestartpos <- orderl[i] + length(title)}
      else if (order[i] == relatedlink[1]) {relatedstartpos <- orderl[i] + length(title)}
      else if (order[i] == dop) {dopstartpos <- orderl[i] + length(title)}
      else if (order[i] == blank) {blankstartpos <- orderl[i] + length(title)}
      else if (order[i] == names) {namesstartpos <- orderl[i] + length(title)}
      else if (order[i] == email) {emailstartpos <- orderl[i] + length(title)}
      else if (order[i] == phone) {phonestartpos <- orderl[i] + length(title)}
      else if (order[i] == reuse) {reusestartpos <- orderl[i] + length(title)}
      else if (order[i] %in% extrafields) {extrastartpos <- append(extrastartpos, orderl[i] + length(title))}
      else if (order[i] == additlinks[1]) {additstartpos <- orderl[i] + length(title)}
      
    }
    
    if (is.null(introstartpos) & is.null(aboutstartpos) & is.null(sourcestartpos) & is.null(relatedstartpos) &
        is.null(dopstartpos) & is.null(blankstartpos) & is.null(namesstartpos) & is.null(emailstartpos) &
        is.null(phonestartpos) & is.null(reusestartpos) & is.null(extrastartpos) & is.null(additstartpos)) {
      
      stop("No starting positions have been generated")
      
    }
    
    if (length(extrastartpos) != length(extrafields)) {
      
      stop("The lengths of the vectors for extrafields and their row starting positions are not equal. Investigate why.")
      
    }
    
  }
  
  intro_hyper <- 0
  about_hyper <- 0
  source_hyper <- 0
  dop_hyper <- 0
  blank_hyper <- 0
  names_hyper <- 0
  phone_hyper <- 0
  
  fields2 <- c(intro, about, source, dop, blank, names, phone)
  
  fields3 <- c("intro", "about", "source", "dop", "blank", "names", "phone")
  
  fields4 <- NULL
  
  for (i in seq_along(fields3)) {
    
    if (!is.null(get(fields3[i]))) {fields4 <- append(fields4, i)}
    
  }
  
  fields5 <- fields3[fields4]
  
  rm(fields3, fields4)
  
  hyper_rx <- "\\[(([[:graph:]]|[[:space:]])+)\\]\\([[:graph:]]+\\)"
  
  for (i in seq_along(fields2)) {
    
    if (grepl(hyper_rx, fields2[i]) == TRUE) {
      
      if (substr(fields2[i], 1, 1) != "[" | substr(fields2[i], nchar(fields2[i]), nchar(fields2[i])) != ")") {
        
        warning(paste0(fields2[i], " - if this is meant to be a hyperlink, it needs to be in the format \"[xxx](xxxxxx)\""))
        
      }
      
      if ("phone" %in% fields5[i]) {
        
        phone <- paste0("[Telephone: ", substr(phone, 2, nchar(phone)))
        
      }
      
      md_rx <- "\\[(([[:graph:]]|[[:space:]])+?)\\]\\([[:graph:]]+?\\)"
      md_match <- regexpr(md_rx, fields2[i], perl = TRUE)
      md_extract <- regmatches(fields2[i], md_match)[[1]]
      
      url_rx <- "(?<=\\]\\()([[:graph:]]|[[:space:]])+(?=\\))"
      url_match <- regexpr(url_rx, md_extract, perl = TRUE)
      url_extract <- regmatches(md_extract, url_match)[[1]]
      
      string_rx <- "(?<=\\[)([[:graph:]]|[[:space:]])+(?=\\])"
      string_match <- regexpr(string_rx, md_extract, perl = TRUE)
      string_extract <- regmatches(md_extract, string_match)[[1]]
      
      string_extract <- gsub(md_rx, string_extract, fields2[i])
      
      x <- stats::setNames(url_extract, string_extract)
      class(x) <- "hyperlink"
      
      if ("intro" %in% fields5[i]) {
        
        intro <- x
        intro_hyper <- 1
        
      } else if ("about" %in% fields5[i]) {
        
        about <- x
        about_hyper <- 1
        
      } else if ("source" %in% fields5[i]) {
        
        source <- x
        source_hyper <- 1
        
      } else if ("dop" %in% fields5[i]) {
        
        dop <- x
        dop_hyper <- 1
        
      } else if ("blank" %in% fields5[i]) {
        
        blank <- x
        blank_hyper <- 1
        
      } else if ("names" %in% fields5[i]) {
        
        names <- x
        names_hyper <- 1
        
      } else if ("phone" %in% fields5[i]) {
        
        phone <- x
        phone_hyper <- 1
        
      }
      
      rm(x, md_rx, md_match, md_extract, url_rx, url_match, url_extract, string_rx, string_match, string_extract)
      
    } else if (grepl(hyper_rx, fields2[i]) == FALSE & "phone" %in% fields5[i]) {
      
      phone <- paste0("Telephone: ", phone)
      
    }
    
  }
  
  if (!is.null(intro)) {
    
    if (is.null(introstartpos)) {
      
      introstart <- length(title) + 1
      
    } else if (!is.null(introstartpos)) {
      
      introstart <- introstartpos
      
    }
    
    openxlsx::writeData(wb, "Cover", "Introductory information", startCol = 1, startRow = introstart)
    openxlsx::writeData(wb, "Cover", intro, startCol = 1, startRow = introstart + 1)
    
  }
  
  if (!is.null(about)) {
    
    if (is.null(aboutstartpos)) {
      
      aboutstart <- length(title) + length(intro) + intro2 + 1
      
    } else if (!is.null(aboutstartpos)) {
      
      aboutstart <- aboutstartpos
      
    }
    
    openxlsx::writeData(wb, "Cover", "About these data", startCol = 1, startRow = aboutstart)
    openxlsx::writeData(wb, "Cover", about, startCol = 1, startRow = aboutstart + 1)
    
  }
  
  if (!is.null(source)) {
    
    if (is.null(sourcestartpos)) {
      
      sourcestart <- length(title) + length(intro) + intro2 + length(about) + about2 + 1
      
    } else if (!is.null(sourcestartpos)) {
      
      sourcestart <- sourcestartpos
      
    }
    
    openxlsx::writeData(wb, "Cover", "Source", startCol = 1, startRow = sourcestart)
    openxlsx::writeData(wb, "Cover", source, startCol = 1, startRow = sourcestart + 1)
    
  }
  
  if (!is.null(relatedlink)) {
    
    if (is.null(relatedstartpos)) {
      
      relatedstart <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + source2 + 1
      
    } else if (!is.null(relatedstartpos)) {
      
      relatedstart <- relatedstartpos
      
    }
    
    relpub <- relatedlink
    names(relpub) <- relatedtext
    class(relpub) <- "hyperlink"
    
    openxlsx::writeData(wb, "Cover", "Related publications", startCol = 1, startRow = relatedstart)
    openxlsx::writeData(wb, "Cover", relpub, startCol = 1, startRow = relatedstart + 1)
    
  }
  
  if (!is.null(dop)) {
    
    if (is.null(dopstartpos)) {
      
      dopstart <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + source2 + length(relatedlink) + related2 + 1
      
    } else if (!is.null(dopstartpos)) {
      
      dopstart <- dopstartpos
      
    }  
    
    openxlsx::writeData(wb, "Cover", "Date of publication", startCol = 1, startRow = dopstart)
    openxlsx::writeData(wb, "Cover", dop, startCol = 1, startRow = dopstart + 1)
    
    
  }
  
  if (!is.null(blank)) {
    
    if (is.null(blankstartpos)) {
      
      blankstart <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + 1
      
    } else if (!is.null(blankstartpos)) {
      
      blankstart <- blankstartpos
      
    }  
    
    openxlsx::writeData(wb, "Cover", "Blank cells", startCol = 1, startRow = blankstart)
    openxlsx::writeData(wb, "Cover", blank, startCol = 1, startRow = blankstart + 1)
    
    
  }
  
  if (!is.null(extrafields)) {
    
    for (i in seq_along(extrafields)) {
      
      if (is.null(extrastartpos)) {
        
        extrastart <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + length(blank) + blank2 + 1 + (2 * i) - 2
        
      } else if (!is.null(extrastartpos)) {
        
        extrastart <- extrastartpos[i]
        
      }
      
      if (grepl(hyper_rx, extrafieldsb[i]) == TRUE) {
        
        if (substr(extrafieldsb[i], 1, 1) != "[" | substr(extrafieldsb[i], nchar(extrafieldsb[i]), nchar(extrafieldsb[i])) != ")") {
          
          warning(paste0(extrafieldsb[i], " - if this is meant to be a hyperlink, it needs to be in the format \"[xxx](xxxxxx)\""))
          
        }
        
        x <- extrafieldsb[i]
        
        md_rx <- "\\[(([[:graph:]]|[[:space:]])+?)\\]\\([[:graph:]]+?\\)"
        md_match <- regexpr(md_rx, x, perl = TRUE)
        md_extract <- regmatches(x, md_match)[[1]]
        
        url_rx <- "(?<=\\]\\()([[:graph:]]|[[:space:]])+(?=\\))"
        url_match <- regexpr(url_rx, md_extract, perl = TRUE)
        url_extract <- regmatches(md_extract, url_match)[[1]]
        
        string_rx <- "(?<=\\[)([[:graph:]]|[[:space:]])+(?=\\])"
        string_match <- regexpr(string_rx, md_extract, perl = TRUE)
        string_extract <- regmatches(md_extract, string_match)[[1]]
        
        string_extract <- gsub(md_rx, string_extract, x)
        
        y <- stats::setNames(url_extract, string_extract)
        class(y) <- "hyperlink"
        
        rm(x, md_rx, md_match, md_extract, url_rx, url_match, url_extract, string_rx, string_match, string_extract)
        
      } else {
        
        y <- extrafieldsb[i]
        
      }
      
      openxlsx::writeData(wb, "Cover", extrafields[i], startCol = 1, startRow = extrastart)
      openxlsx::writeData(wb, "Cover", y, startCol = 1, startRow = extrastart + 1)
      
      rm(y)
      
    }
    
  }
  
  if (!is.null(additlinks)) {
    
    if (is.null(additstartpos)) {
      
      additlinkstart <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + length(blank) + blank2 + length(extrafields) + length(extrafieldsb) + 1
      
    } else if (!is.null(additstartpos)) {
      
      additlinkstart <- additstartpos
      
    }  
    
    additional <- additlinks
    names(additional) <- addittext
    class(additional) <- "hyperlink"
    
    openxlsx::writeData(wb, "Cover", "Additional links", startCol = 1, startRow = additlinkstart)
    openxlsx::writeData(wb, "Cover", additional, startCol = 1, startRow = additlinkstart + 1)
    
  }
  
  if (!is.null(names)) {
    
    if (is.null(namesstartpos)) {
      
      namesstart <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + length(blank) + blank2 + length(extrafields) + length(extrafieldsb) + additlinks2 + length(additlinks) + 1
      
    } else if (!is.null(namesstartpos)) {
      
      namesstart <- namesstartpos
      
    }
    
    openxlsx::writeData(wb, "Cover", "Contact", startCol = 1, startRow = namesstart)
    openxlsx::writeData(wb, "Cover", names, startCol = 1, startRow = namesstart + 1)
    
  }
  
  normalformat <- openxlsx::createStyle(valign = "top", wrapText = TRUE)
  subtitleformat <- openxlsx::createStyle(fontSize = fontszst, valign = "bottom", wrapText = TRUE, textDecoration = "bold")
  titleformat <- openxlsx::createStyle(fontSize = fontszt, valign = "bottom", wrapText = TRUE, textDecoration = "bold")
  linkformat <- openxlsx::createStyle(fontColour = "blue", valign = "top", wrapText = TRUE, textDecoration = "underline")
  
  if (is.null(colwid_spec) | !is.numeric(colwid_spec) | length(colwid_spec) > 1) {
    
    openxlsx::setColWidths(wb, "Cover", cols = 1, widths = 100)
    
    if (!is.null(colwid_spec) & !is.numeric(colwid_spec)) {
      
      warning("colwid_spec has not been provided as a numeric value and so the default width of 100 has been used")
      
    } else if (!is.null(colwid_spec) & length(colwid_spec) > 1) {
      
      warning("colwid_spec has been provided as a vector with more than one element and so the default width of 100 has been used")
      
    }
    
  } else if (is.numeric(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Cover", cols = 1, widths = colwid_spec)
    
  }
  
  openxlsx::addStyle(wb, "Cover", normalformat, rows = c(1:covernumrow), cols = 1)
  
  if (!is.null(intro)) {
    
    openxlsx::setRowHeights(wb, "Cover", introstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = introstart, cols = 1)
    
    if (intro_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = introstart + 1, cols = 1)}
    
  }
  
  if (!is.null(about)) {
    
    openxlsx::setRowHeights(wb, "Cover", aboutstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = aboutstart, cols = 1)
    
    if (about_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = aboutstart + 1, cols = 1)}
    
  }
  
  if (!is.null(source)) {
    
    openxlsx::setRowHeights(wb, "Cover", sourcestart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = sourcestart, cols = 1)
    
    if (source_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = sourcestart + 1, cols = 1)}
    
  }
  
  if (!is.null(relatedlink)) {
    
    openxlsx::setRowHeights(wb, "Cover", relatedstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = relatedstart, cols = 1)
    openxlsx::addStyle(wb, "Cover", linkformat, rows = (relatedstart + 1):(relatedstart + length(relatedlink)), cols = 1)
    
  }
  
  if (!is.null(dop)) {
    
    openxlsx::setRowHeights(wb, "Cover", dopstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = dopstart, cols = 1)
    
    if (dop_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = dopstart + 1, cols = 1)}
    
  }
  
  if (!is.null(blank)) {
    
    openxlsx::setRowHeights(wb, "Cover", blankstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = blankstart, cols = 1)
    
    if (blank_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = blankstart + 1, cols = 1)}
    
  }
  
  if (!is.null(extrafields)) {
    
    for (i in seq_along(extrafields)) {
      
      if (is.null(extrastartpos)) {
        
        extrastart <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + length(blank) + blank2 + 1 + (2 * i) - 2
        
      } else if (!is.null(extrastartpos)) {
        
        extrastart <- extrastartpos[i]
        
      }
      
      openxlsx::setRowHeights(wb, "Cover", extrastart, fontszst * (25/14))
      openxlsx::addStyle(wb, "Cover", subtitleformat, rows = extrastart, cols = 1)
      
      if (grepl(hyper_rx, extrafieldsb[i]) == TRUE) {
        
        openxlsx::addStyle(wb, "Cover", linkformat, rows = extrastart + 1, cols = 1)
        
      }
      
    }
    
  }
  
  if (!is.null(additlinks)) {
    
    openxlsx::setRowHeights(wb, "Cover", additlinkstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = additlinkstart, cols = 1)
    openxlsx::addStyle(wb, "Cover", linkformat, rows = (additlinkstart + 1):(additlinkstart + length(additlinks)), cols = 1)
    
  }
  
  if (!is.null(names)) {
    
    openxlsx::setRowHeights(wb, "Cover", namesstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = namesstart, cols = 1)
    
    if (names_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = namesstart + 1, cols = 1)}
    
  }
  
  openxlsx::setRowHeights(wb, "Cover", 2, fontszst * (34/14))
  
  openxlsx::addStyle(wb, "Cover", titleformat, rows = 1, cols = 1)
  
  # Create a hyperlink for any given email address
  
  if (!is.null(email)) {
    
    x <- paste0("mailto:", email)
    names(x) <- email
    class(x) <- "hyperlink"
    
    if (is.null(emailstartpos)) {
      
      emailstart <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + length(blank) + blank2 + length(extrafields) + length(extrafieldsb) + additlinks2 + length(additlinks) + length(names) + names2 + 1
      
    } else if (!is.null(emailstartpos)) {
      
      emailstart <- emailstartpos
      
    }
    
    openxlsx::writeData(wb, "Cover", x, startCol = 1, startRow = emailstart)
    
    openxlsx::addStyle(wb, "Cover", linkformat, rows = emailstart, cols = 1)
    
    emailformat <- openxlsx::createStyle(fontColour = "blue", valign = "bottom", textDecoration = "underline", wrapText = TRUE)
    
    if (emailstart == 2) {openxlsx::addStyle(wb, "Cover", emailformat, rows = 2, cols = 1)}
    
  }
  
  if (!is.null(phone)) {
    
    if (is.null(phonestartpos)) {
      
      phonestart <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + length(blank) + blank2 + length(extrafields) + length(extrafieldsb) + additlinks2 + length(additlinks) + length(names) + names2 + length(email) + 1
      
    } else if (!is.null(phonestartpos)) {
      
      phonestart <- phonestartpos
      
    }
    
    openxlsx::writeData(wb, "Cover", phone, startCol = 1, startRow = phonestart)
    
    phoneformat <- openxlsx::createStyle(valign = "bottom")
    
    if (phone_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = phonestart, cols = 1, stack = TRUE)}
    
    if (phonestart == 2) {openxlsx::addStyle(wb, "Cover", phoneformat, rows = 2, cols = 1, stack = TRUE)}
    
  }
  
  if (reuse == "Yes") {
    
    if (!is.null(govdept)) {
      
      if (stringr::str_starts(tolower(govdept), "the ")) {
        
        govdept <- substr(govdept, 5, nchar(govdept))
        
      }
      
    }
    
    if (is.null(govdept)) {
      
      orgwording <- "our organisation"
      
    } else if (tolower(govdept) != "ons" & tolower(govdept) != "office for national statistics") {
      
      orgwording <- paste0("the ", govdept, " - Source: ", govdept)
      
    } else if (!is.null(govdept) & (tolower(govdept) == "ons" | tolower(govdept) == "office for national statistics")) {
      
      orgwording <- "the Office for National Statistics - Source: Office for National Statistics"
      
    }
    
    reuse1 <- paste0("You may re-use this publication (not including logos) free of charge in any format or medium, under the terms of the Open Government Licence. Users should include a source accreditation to ", orgwording, " licensed under the Open Government Licence.")
    reuse2 <- "Alternatively you can write to: Information Policy Team, The National Archives, Kew, Richmond, Surrey, TW9 4DU; or email: psi@nationalarchives.gov.uk"
    reuse3 <- "Where we have identified any third party copyright information you will need to obtain permission from the copyright holders concerned."
    licencelink <- "https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/"
    licencetext <- "View the Open Government Licence"
    
    if (is.null(reusestartpos)) {
      
      reusestart <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + length(blank) + blank2 + length(extrafields) + length(extrafieldsb) + additlinks2 + length(additlinks) + length(names) + names2 + length(email) + length(phone) + 1
      
    } else if (!is.null(reusestartpos)) {
      
      reusestart <- reusestartpos
      
    }  
    
    openxlsx::writeData(wb, "Cover", "Reusing this publication", startRow = reusestart, startCol = 1)
    openxlsx::writeData(wb, "Cover", reuse1, startRow = reusestart + 1, startCol = 1)
    
    reuselink <- licencelink
    names(reuselink) <- licencetext
    class(reuselink) <- "hyperlink"
    
    openxlsx::writeData(wb, "Cover", reuselink, startRow = reusestart + 2, startCol = 1)
    openxlsx::writeData(wb, "Cover", reuse2, startRow = reusestart + 3, startCol = 1)
    openxlsx::writeData(wb, "Cover", reuse3, startRow = reusestart + 4, startCol = 1)
    
    openxlsx::setRowHeights(wb, "Cover", reusestart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = reusestart, cols = 1)
    openxlsx::addStyle(wb, "Cover", linkformat, rows = reusestart + 2, cols = 1)
    
  }
  
  # If gridlines are not wanted then they are removed
  
  if (gridlines == "No") {
    
    openxlsx::showGridLines(wb, "Cover", showGridLines = FALSE)
    
  }
  
}

###################################################################################################################
###################################################################################################################
# NOTES

# addnote function will add a note and its description to the workbook, specifically in the notes worksheet
# Add notes if wanted, if not then do not run the addnote function
# A link can be provided with each note as well a list of tables that the note applies to
# notenumber and notetext are the only compulsory parameters
# All other parameters are optional and preset to NULL, so only need to be defined if they are wanted
# applictabtext should be set to a vector of sheet names if a column is wanted which lists which worksheets a note is applicable to
# linktext1 and linktext2: linktext1 should be the text you want to appear and linktext2 should be the underlying link to a website, file etc


addnote <- function(notenumber, notetext, applictabtext = NULL, linktext1 = NULL, linktext2 = NULL) {
  
  # Checking that a notes page is wanted, based on whether a worksheet was created in the initial workbook
  
  if (!("Notes" %in% names(wb))) {
    
    stop("notestab has not been set to \"Yes\" in the workbook function call")
    
  }
  
  # Checking some parameters to see if they are properly populated, if not the function will error
  
  if (is.null(notenumber) | is.null(notetext)) {
    
    stop("No notenumber or notetext entered. Must have notenumber and notetext.")
    
  } else if (notenumber == "" | notetext == "") {
    
    stop("No notenumber or notetext entered. Must have notenumber and notetext.")
    
  }
  
  if (length(notenumber) > 1 | length(notetext) > 1) {
    
    stop("One or both of notenumber and notetext are not populated properly. They must be a single entity and not a vector.")
    
  }
  
  if (!is.null(applictabtext) & !is.character(applictabtext)) {
    
    stop("The parameter applictabtext is not populated properly. If it is not NULL then it has to be a string. If more than one element is needed then it should be expressed as a vector e.g., applictabtext = c(\"Table_1\", \"Table_2\")")
    
  }
  
  if (is.null(applictabtext) & autonotes2 == "Yes") {
    
    stop("Automatic listing of notes on tables has been selected but a note has no tables applicable to it")
    
  }
  
  if (length(linktext1) > 1 | length(linktext2) > 1) {
    
    stop("linktext1 and linktext2 can only be single entities and not vectors of length greater than one")
    
  }
  
  if (is.numeric(notenumber)) {
    
    notenumber <- paste0("note", as.character(notenumber))
    
  } else if (is.character(notenumber) & !grepl("\\D", notenumber, perl = TRUE) == TRUE) {
    
    notenumber <- paste0("note", notenumber)
    
  }
  
  notetemp1 <- substr(notenumber, 1, 4)
  notetemp2 <- substr(notenumber, 5, nchar(notenumber))
  
  if (tolower(notetemp1) == "note") {
    
    notenumber <- paste0("note", notetemp2)
    notetemp1 <- "note"
    
  }
  
  if (notetemp1 != "note" | !grepl("\\D", notetemp2, perl = TRUE) == FALSE) {
    
    stop("The notenumber parameter is not properly populated. It should take the form of \"note\" followed by a number.")
    
  }
  
  rm(notetemp1, notetemp2)
  
  if (!is.null(applictabtext)) {
    
    check <- 0
  
    for (i in seq_along(applictabtext)) {
      
      if (tolower(applictabtext[i]) == "all") {applictabtext[i] <- "All"}
      
      if (length(applictabtext) > 1 & applictabtext[i] == "All") {
        
        stop("The applictabtext parameter includes two or more elements but one of the elements is \"All\"")
        
      }
      
      if (stringr::str_detect(applictabtext[i], " ") | stringr::str_detect(applictabtext[i], ",")) {
        
        stop("The applictabtext contains whitespace or a comma. applictabtext should either be a single word (e.g., \"All\") or expressed as a vector (e.g., c(\"Table_1\", \"Table_2\"))")
        
      }
      
      if (tolower(applictabtext[i]) == "none") {
        
        stop("A note should be applicable to at least one of the tables. applictabtext should not be set to \"None\".")
        
      }
      
      if (!(applictabtext[i] %in% tabcontents[[1]]) & applictabtext[i] != "All") {
        
        print(paste0(applictabtext[i], " not in table of contents"))
        check <- 1
        
      }
      
    }
    
    if (check == 1) {
      
      stop("At least one of the tables mentioned in the applictabtext parameter is not in the table of contents")
      
    }
    
  }
  
  # Cleaning up the linktext1 and linktext2 parameters where no link information provided
    
  if (is.null(applictabtext)) {applictabtext <- ""}
  
  applictabtext2 <- paste(applictabtext, collapse = ", ")
  
  if (is.null(linktext1)) {
    
    linktext1 <- "No additional link"
    
  } else if (linktext1 == "") {
    
    linktext1 <- "No additional link"
    
  } else if (tolower(linktext1) == "no additional link") {
    
    linktext1 <- "No additional link"
    
  }
  
  if (is.null(linktext2)) {
    
    linktext2 <- ""
    
  } else if (tolower(linktext2) == "no additional link") {
    
    linktext2 <- ""
    
  }
  
  if (linktext1 != "" & linktext1 != "No additional link" & linktext2 == "") {
    
    stop("Invalid combination of linktext1 and linktext2")
    
  }
  
  if (linktext2 != "" & (linktext1 == "" | linktext1 == "No additional link")) {
    
    stop("Invalid combination of linktext1 and linktext2")
    
  }
  
  # Checking for any issues with duplication
  
  notesdfx <- notesdf %>%
    dplyr::add_row("Note number" = notenumber, "Note text" = notetext, "Applicable tables" = applictabtext2, "Link1" = linktext1, "Link2" = linktext2) %>%
    dplyr::mutate(Link2 = dplyr::case_when(Link1 == "No additional link" ~ "No additional link",
                                           TRUE ~ Link2)) %>%
    dplyr::mutate(Link = dplyr::case_when(Link1 == "No additional link" ~ "No additional link",
                                          TRUE ~ paste0("HYPERLINK(\"", Link2, "\", \"", Link1, "\")")))
  
  notesdf2 <- notesdfx %>%
    dplyr::rename(note_number = "Note number") %>%
    dplyr::group_by(note_number) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(count) / n())
  
  notesdf3 <- notesdfx %>%
    dplyr::rename(note_text = "Note text") %>%
    dplyr::group_by(note_text) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(count) / n())
  
  notesdf4 <- notesdfx %>%
    dplyr::group_by(Link1) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link1) | Link1 == "" | Link1 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(count)) / n()) 
  
  notesdf5 <- notesdfx %>%
    dplyr::group_by(Link2) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link2) | Link2 == "" | Link2 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(count)) / n()) 
  
  if (notesdf2$check > 1) {
    
    stop("Duplicated note number(s)")
    
  }
  
  if (notesdf3$check > 1) {
    
    warning("Duplicated note text(s). Explore to see if this is an issue.")
    
  }
  
  if (notesdf4$check > 1) {
    
    warning("Duplicated link text(s) (Link1). Explore to see if this is an issue.")
    
  }
  
  if (notesdf5$check > 1) {
    
    warning("Duplicated link text(s) (Link2). Explore to see if this is an issue.")
    
  }
  
  # Create a notes data frame in the global environment
  
  notesdf <<- notesdfx
  
  rm(notesdf2, notesdf3, notesdf4, notesdf5, notesdfx)
  
}

# notestab function will create a notes worksheet in the workbook and includes notes added using the addnote function
# If notes not wanted, then do not run the notestab function
# There are three parameters and they are optional and preset. Change contentslink to "No" if you want a contents tab but do not want a link to it in the notes tab. Change gridlines to "No" if gridlines are not wanted.
# Column widths are automatically set but the user can specify the required widths in colwid_spec
# Extra columns can be added by setting extracols to "Yes" and creating a dataframe extracols_notes with the desired extra columns


notestab <- function(contentslink = NULL, gridlines = "Yes", colwid_spec = NULL, extracols = NULL) {
  
  # Check that a notes page is wanted, based on whether a worksheet was created in the initial workbook
  
  if (!("Notes" %in% names(wb))) {
    
    stop("notestab has not been set to \"Yes\" in the workbook function call")
    
  }
  
  # Checking some parameters to see if they are properly populated, if not the function will error
  
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
    
    stop("gridlines has not been populated properly. It must be a single word, either \"Yes\" or \"No\".")
    
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
    
    stop("extracols has not been populated properly. It must be a single word, either \"Yes\" or \"No\".")
    
  }
  
  # Automatically detecting if information of which tables notes apply to has been provided or not
  
  notesdfx <- notesdf %>%
    dplyr::rename(applictab = "Applicable tables") %>%
    dplyr::filter(applictab != "")
  
  if (nrow(notesdf) > 0 & nrow(notesdf) == nrow(notesdfx)) {
    
    applictabs <- "Yes"
    
  } else if (nrow(notesdf) > 0 & nrow(notesdfx) == 0) {
    
    applictabs <- "No"
    
  } else if (nrow(notesdf) > 0 & nrow(notesdf) != nrow(notesdfx)) {
    
    stop("There may be a note without applicable tables allocated to it while other notes do have applicable tables allocated to them")
    
  } else if (nrow(notesdf) == 0) {
    
    stop("The notesdf dataframe contains no observations")
    
  }
  
  rm(notesdfx)
  
  # Automatically detecting if links have been provided or not
  
  notesdfx <- notesdf %>%
    dplyr::filter(!is.na(Link1) & Link1 != "No additional link")
  
  if (nrow(notesdfx) > 0) {
    
    links <- "Yes"
    
  } else if (nrow(notesdf) > 0 & nrow(notesdfx) == 0) {
    
    links <- "No"
    
  } else if (nrow(notesdf) == 0) {
    
    stop("The notesdf dataframe contains no observations")
    
  }
  
  rm(notesdfx)
  
  # Identifying the row number of notes which have a link associated to them
  
  notesdfy <- notesdf %>%
    dplyr::mutate(linkno = dplyr::case_when(is.na(Link1) | Link1 == "No additional link" ~ NA_real_,
                                            TRUE ~ dplyr::row_number())) %>%
    dplyr::filter(!is.na(linkno))
  
  if (nrow(notesdfy) > 0) {
    
    linkrange <- notesdfy$linkno
    
  } else {
    
    linkrange <- NULL
    
  }
  
  rm(notesdfy)
  
  # Checks associated with the automatic generation of note information for the main data tables
  
  if (applictabs == "Yes" & autonotes2 != "Yes") {
    
    warning("The applictabs parameter has been set to \"Yes\" but the automatic listing of notes on a worksheet has not been selected")
    
  } else if (applictabs == "No" & autonotes2 == "Yes") {
    
    stop("The applictabs parameter has been set to \"No\" but the automatic listing of notes on a worksheet has been selected")
    
  }
  
  # Determining whether a link to the contents page is wanted
  
  if (is.null(contentslink)) {
    
    contentslink <- "Yes"
    
  } else if (tolower(contentslink) == "no" | tolower(contentslink) == "n") {
    
    contentslink <- "No"
    
  }
  
  if (contentslink == "No" | !("Contents" %in% names(wb))) {
    
    contentstab <- "No"
    
  } else if ("Contents" %in% names(wb)) {
    
    contentstab <- "Yes"
    
  } else {
    
    contentstab <- "No"
    
  }
  
  # Checking for any note information duplication
  
  notesdf2 <- notesdf %>%
    dplyr::rename(note_number = "Note number") %>%
    dplyr::group_by(note_number) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(count) / n())
  
  notesdf3 <- notesdf %>%
    dplyr::rename(note_text = "Note text") %>%
    dplyr::group_by(note_text) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(count) / n())
  
  notesdf4 <- notesdf %>%
    dplyr::group_by(Link1) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link1) | Link1 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(count)) / n()) 
  
  notesdf5 <- notesdf %>%
    dplyr::group_by(Link2) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link2) | Link2 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(count)) / n()) 
  
  notesdf6 <- notesdf %>%
    dplyr::rename(applictab = "Applicable tables") %>%
    dplyr::filter(is.na(applictab))
  
  if (notesdf2$check > 1) {
    
    stop("Duplicated note number(s)")
    
  }
  
  if (notesdf3$check > 1) {
    
    warning("Duplicated note text(s). Explore to see if this is an issue.")
    
  }
  
  if (notesdf4$check > 1) {
    
    warning("Duplicated link text(s) (Link1). Explore to see if this is an issue.")
    
  }
  
  if (notesdf5$check > 1) {
    
    warning("Duplicated link text(s) (Link2). Explore to see if this is an issue.")
    
  }
  
  rm(notesdf2, notesdf3, notesdf4, notesdf5)
  
  if (links != "No" & links != "Yes") {
    
    stop("links not set to \"Yes\" or \"No\". There must be an issue with link information provided with the notes.")
    
  }
  
  if (applictabs != "No" & applictabs != "Yes") {
    
    stop("applictabs not set to \"Yes\" or \"No\"")
    
  } else if (applictabs == "Yes" & nrow(notesdf6) > 0) {
    
    stop("Applicable table column wanted but contains an empty cell or cells")
    
  }
  
  rm(notesdf6)
  
  # Creating a notes table with the required columns
  
  if (extracols == "Yes" & exists("extracols_notes", envir = .GlobalEnv)) {
    
    if (nrow(extracols_notes) != nrow(notesdf)) {
      
      stop("The number of rows in the notes table is not the same as in the dataframe of extra columns")
      
    }
    
  } else if (extracols == "Yes" & !(exists("extracols_notes", envir = .GlobalEnv))) {
    
    warning("extracols has been set to \"Yes\" but the extracols_notes dataframe does not exist. No extra columns will be added.")
    
  } else if (extracols == "No" & exists("extracols_notes", envir = .GlobalEnv)) {
    
    warning("extracols has been set to \"No\" but a dataframe extracols_notes exists. No extra columns have been added.")
    
  }
  
  if (links == "Yes" & applictabs == "Yes") {
  
    notesdf <<- notesdf %>%
      dplyr::select("Note number", "Note text", "Applicable tables", "Link") %>%
      {if (exists("extracols_notes", envir = .GlobalEnv)) dplyr::bind_cols(., extracols_notes) else .}
    
    class(notesdf$Link) <- "formula"
    
  } else if (links == "Yes") {
    
    notesdf <<- notesdf %>%
      dplyr::select("Note number", "Note text", "Link") %>%
      {if (exists("extracols_notes", envir = .GlobalEnv)) dplyr::bind_cols(., extracols_notes) else .}
    
    class(notesdf$Link) <- "formula"
    
  } else if (links == "No" & applictabs == "Yes") {
    
    notesdf <<- notesdf %>%
      dplyr::select("Note number", "Note text", "Applicable tables") %>%
      {if (exists("extracols_notes", envir = .GlobalEnv)) dplyr::bind_cols(., extracols_notes) else .}
    
  } else if (links == "No") {
    
    notesdf <<- notesdf %>%
      dplyr::select("Note number", "Note text") %>%
      {if (exists("extracols_notes", envir = .GlobalEnv)) dplyr::bind_cols(., extracols_notes) else .}
    
  }
  
  if ("Link" %in% colnames(notesdf)) {
    
    notesdfcols <- colnames(notesdf)
    
    for (i in seq_along(notesdfcols)) {
      
      if (notesdfcols[i] == "Link") {linkcolpos <- i}
      
    }
    
  }
  
  if (extracols == "Yes" & exists("extracols_notes", envir = .GlobalEnv)) {
    
    if (any(duplicated(colnames(notesdf))) == TRUE) {
      
      warning("There is at least one duplicate column name in the notes table and the extracols_notes dataframe")
      
    }
    
  } 
  
  # Define formatting to be used later on
  
  normalformat <- openxlsx::createStyle(valign = "top")
  topformat <- openxlsx::createStyle(valign = "bottom")
  linkformat <- openxlsx::createStyle(fontColour = "blue", valign = "top", textDecoration = "underline")
  
  openxlsx::addStyle(wb, "Notes", normalformat, rows = 1:(nrow(notesdf) + 4), cols = 1:ncol(notesdf), gridExpand = TRUE)
  
  openxlsx::writeData(wb, "Notes", "Notes", startCol = 1, startRow = 1)
  
  titleformat <- openxlsx::createStyle(fontSize = fontszt, textDecoration = "bold")
  
  openxlsx::addStyle(wb, "Notes", titleformat, rows = 1, cols = 1)
  
  openxlsx::writeData(wb, "Notes", "This worksheet contains one table.", startCol = 1, startRow = 2)
  openxlsx::addStyle(wb, "Notes", topformat, rows = 2, cols = 1)
  
  extraformat <- openxlsx::createStyle(valign = "top")
  
  if (contentstab == "Yes") {
    
    openxlsx::writeFormula(wb, "Notes", startRow = 3, x = openxlsx::makeHyperlinkString("Contents", row = 1, col = 1, text = "Link to contents"))
    
    openxlsx::addStyle(wb, "Notes", linkformat, rows = 3, cols = 1)
    
    startingrow <- 4
    
  } else if (contentstab == "No") {
    
    startingrow <- 3
    
    openxlsx::addStyle(wb, "Notes", extraformat, rows = 2, cols = 1, stack = TRUE)
    
  }
  
  if (links == "Yes" & is.null(linkrange)) {
    
    stop("Links are required in the notes tab but the row numbers where links should be have not been generated (i.e., linkrange not populated)")
    
  } else if (links == "Yes" & !is.null(linkrange)) {
    
    openxlsx::addStyle(wb, "Notes", linkformat, rows = linkrange + startingrow, cols = linkcolpos, gridExpand = TRUE)
    
  }
  
  openxlsx::writeDataTable(wb, "Notes", notesdf, tableName = "notes", startRow = startingrow, startCol = 1, withFilter = FALSE, tableStyle = "none")
  
  # The if statement below is required so that "No additional link" appears as text only in the final spreadsheet, rather than as a hyperlink
  
  if ("Link" %in% colnames(notesdf)) {
    
    noaddlinks <- notesdf[["Link"]]
    
    for (i in seq_along(noaddlinks)) {
      
      if (noaddlinks[i] == "No additional link") {
        
        openxlsx::writeData(wb, "Notes", "No additional link", startCol = linkcolpos, startRow = startingrow + i)
        
      } 
      
    }
    
  }
  
  headingsformat <- openxlsx::createStyle(textDecoration = "bold", wrapText = TRUE, border = NULL, valign = "top")
  
  openxlsx::addStyle(wb, "Notes", headingsformat, rows = startingrow, cols = 1:ncol(notesdf))
  
  extraformat2 <- openxlsx::createStyle(wrapText = TRUE)
  
  openxlsx::addStyle(wb, "Notes", extraformat2, rows = (startingrow + 1):(nrow(notesdf) + startingrow + 1), cols = 1:ncol(notesdf), stack = TRUE, gridExpand = TRUE)
  
  # Determining column widths
  
  if ((!is.null(colwid_spec) & !is.numeric(colwid_spec)) | (!is.null(colwid_spec) & length(colwid_spec) != ncol(notesdf))) {
    
    warning("colwid_spec is either a non-numeric value or a vector not of the same length as the number of columns desired in the notes tab. The widths will be determined automatically.")
    colwid_spec <- NULL
    
  }
  
  numchars <- max(nchar(notesdf$"Note text")) + 10
  
  if (links == "Yes" & applictabs == "No" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Notes", cols = c(1,2,3,4:max(ncol(notesdf),4)), widths = c(15, min(numchars, 100), 50, "auto"))
    
  } else if (links == "Yes" & applictabs == "Yes" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Notes", cols = c(1,2,3,4,5:max(ncol(notesdf),5)), widths = c(15, min(numchars, 100), 20, 50, "auto"))
    
  } else if (links == "No" & applictabs == "Yes" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Notes", cols = c(1,2,3,4:max(ncol(notesdf),4)), widths = c(15, min(numchars, 100), 20, "auto"))
    
  } else if (links == "No" & applictabs == "No" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Notes", cols = c(1,2,3:max(ncol(notesdf),3)), widths = c(15, min(numchars, 100), "auto"))
    
  } else if (!is.null(colwid_spec) & is.numeric(colwid_spec) & length(colwid_spec) == ncol(notesdf)) {
    
    openxlsx::setColWidths(wb, "Notes", cols = c(1:ncol(notesdf)), widths = colwid_spec)
    
  }
  
  openxlsx::setRowHeights(wb, "Notes", startingrow - 1, fontsz * (25/12))
  
  # Creating the text to be inserted in the main data tables regarding which notes are associated with which table
  
  if (applictabs == "Yes" & autonotes2 == "Yes") {
  
    tabcontents2 <- tabcontents %>%
      dplyr::filter(.[[1]] != "Notes" & .[[1]] != "Definitions")
    
    tablelist <- tabcontents2[[1]]
    
    for (i in seq_along(tablelist)) {
      
      notesdf7 <- notesdf %>%
        dplyr::rename(applic_tab = "Applicable tables") %>%
        dplyr::mutate(applic_tab = dplyr::case_when(applic_tab == "All" ~ paste(tablelist, collapse = ", "),
                                                    TRUE ~ applic_tab)) %>%
        dplyr::filter(stringr::str_detect(applic_tab, tablelist[i]) == TRUE)
      
      notes <- paste0("[", notesdf7[[1]], "]")
        
      if (length(notes) > 2) {
          
        notes1 <- utils::tail(notes, 2)
          
        notes2 <- utils::head(notes, -2)
          
        notes3 <- paste(notes1, collapse = " and ")
        notes4 <- paste0(", ", notes3)
          
        notes5 <- paste(notes2, collapse = ", ")
          
        notes6 <- paste0(notes5, notes4)
          
        rm(notes1, notes2, notes3, notes4, notes5)
          
      } else if (length(notes) == 2) {
          
        notes6 <- paste(notes, collapse = " and ")
          
      } else if (length(notes) == 1) {
          
        notes6 <- paste0(notes)
          
      }
      
      if (nrow(notesdf7) == 0) {
        
        notes7 <- "This worksheet contains one table."
        warning(paste0(tablelist[i], " has no notes associated with it. Check that this is intentional."))
        
      } else {
        
        notes7 <- paste0("This worksheet contains one table. For notes, see ", notes6, " on the notes worksheet.")
        
      }
        
      tempstartrow <- get(paste0(tablelist[i], "_startrow"), envir = .GlobalEnv)
        
      openxlsx::writeData(wb, tablelist[i], notes7, startCol = 1, startRow = tempstartrow)
        
      rm(notes, notes6, notes7, tempstartrow, notesdf7)
        
    }
    
    rm(tabcontents2, tablelist)
    
  }
  
  # Checking that tables associated with notes are actually in the table of contents
  
  if (applictabs == "Yes") {
    
    tabcontents2 <- tabcontents %>%
      dplyr::filter(.[[1]] != "Notes" & .[[1]] != "Definitions")
    
    tablelist <- tabcontents2[[1]]
    tablelist2 <- paste(tablelist, collapse = ", ")
    
    notesdf8 <- notesdf %>%
      dplyr::rename(applic_tab = "Applicable tables") %>%
      dplyr::mutate(applic_tab2 = tablelist2) %>%
      dplyr::mutate(applic_tab3 = dplyr::case_when(applic_tab == applic_tab2 ~ 1,
                                                   TRUE ~ 0)) %>%
      dplyr::summarise(applic_tab4 = sum(applic_tab3))
    
    if (notesdf8$applic_tab4 >= 1) {
      
      warning("There is at least one occurrence where the list of applicable tables appears to be all of the tables. The list could read \"All\" instead.")
      
    }
    
    notesdf9 <- notesdf %>%
      dplyr::rename(applic_tab = "Applicable tables") %>%
      dplyr::filter(applic_tab != "All")
    
    applictablist <- notesdf9$applic_tab
    
    applictablist2 <- paste(applictablist, collapse = ", ")
    
    applictablist3 <- unique(unlist(strsplit(applictablist2, ", ")))
    
    for (i in seq_along(applictablist3)) {
      
      if(!(applictablist3[i] %in% tabcontents2[[1]])) {
        
        stop(paste0(applictablist3), " is not in the table of contents")
        
      }
      
    }
    
    rm(tablelist, tablelist2, notesdf8, tabcontents2, notesdf9, applictablist, applictablist2, applictablist3)
  
  }
  
  rm(list = ls(pattern = "_startrow", envir = .GlobalEnv), envir = .GlobalEnv)
  
  rm(notesdf, envir = .GlobalEnv)
  
  # If gridlines are not wanted, then they are removed
  
  if (gridlines == "No") {
    
    openxlsx::showGridLines(wb, "Notes", showGridLines = FALSE)
    
  }
  
}

###################################################################################################################
###################################################################################################################
# DEFINITIONS

# adddefinition function will add a definition and its description to the workbook, specifically in the definitions worksheet
# Add definitions if wanted, if not then do run the adddefinition function
# term and definition are compulsory parameters
# A link can be added with each definition
# linktext1 and linktext2: linktext1 should be the text you want to appear and linktext2 should be the underlying link to a website, file etc


adddefinition <- function(term, definition, linktext1 = NULL, linktext2 = NULL) {
  
  # Checking that a definitions page is wanted, based on whether a worksheet was created in the initial workbook
  
  if (!("Definitions" %in% names(wb))) {
    
    stop("definitionstab has not been set to \"Yes\" in the workbook function call")
    
  }
  
  # Checking parameters have been populated properly, if not then the function will error
  
  if (is.null(term) | is.null(definition)) {
    
    stop("No term or definition entered. Must have term and definition.")
    
  } else if (term == "" | definition == "") {
    
    stop("No term or definition entered. Must have term and definition.")
    
  }
  
  if (!is.null(term) & !is.character(term)) {
    
    stop("The parameter term is not populated properly. If it is not NULL then it has to be a string.")
    
  }
  
  if (!is.null(definition) & !is.character(definition)) {
    
    stop("The parameter definition is not populated properly. If it is not NULL then it has to be a string.")
    
  }
  
  if (length(linktext1) > 1 | length(linktext2) > 1) {
    
    stop("linktext1 and linktext2 can only be single entities and not vectors of length greater than one")
    
  }
  
  if (is.null(linktext1)) {
    
    linktext1 <- "No additional link"
    
  } else if (linktext1 == "") {
    
    linktext1 <- "No additional link"
    
  } else if (tolower(linktext1) == "no additional link") {
    
    linktext1 <- "No additional link"
    
  }
  
  if (is.null(linktext2)) {
    
    linktext2 <- ""
    
  } else if (tolower(linktext2) == "no additional link") {
    
    linktext2 <- ""
    
  }
  
  if (linktext1 != "" & linktext1 != "No additional link" & linktext2 == "") {
    
    stop("Invalid combination of linktext1 and linktext2")
    
  }
  
  if (linktext2 != "" & (linktext1 == "" | linktext1 == "No additional link")) {
    
    stop("Invalid combination of linktext1 and linktext2")
    
  }
  
  # Check for any duplication of definitions
  
  definitionsdfx <- definitionsdf %>%
    dplyr::add_row("Term" = term, "Definition" = definition, "Link1" = linktext1, "Link2" = linktext2) %>%
    dplyr::mutate(Link2 = dplyr::case_when(Link1 == "No additional link" ~ "No additional link",
                                           TRUE ~ Link2)) %>%
    dplyr::mutate(Link = dplyr::case_when(Link1 == "No additional link" ~ "No additional link",
                                          TRUE ~ paste0("HYPERLINK(\"", Link2, "\", \"", Link1, "\")")))
  
  definitionsdf2 <- definitionsdfx %>%
    dplyr::group_by(Term) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(count) / n())
  
  definitionsdf3 <- definitionsdfx %>%
    dplyr::group_by(Definition) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(count) / n())
  
  definitionsdf4 <- definitionsdfx %>%
    dplyr::group_by(Link1) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link1) | Link1 == "" | Link1 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(count)) / n()) 
  
  definitionsdf5 <- definitionsdfx %>%
    dplyr::group_by(Link2) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link2) | Link2 == "" | Link2 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(count)) / n()) 
  
  if (definitionsdf2$check > 1) {
    
    stop("Duplicated term(s)")
    
  }
  
  if (definitionsdf3$check > 1) {
    
    warning("Duplicated definition(s). Explore to see if this is an issue.")
    
  }
  
  if (definitionsdf4$check > 1) {
    
    warning("Duplicated link text(s) (Link1). Explore to see if this is an issue.")
    
  }
  
  if (definitionsdf5$check > 1) {
    
    warning("Duplicated link text(s) (Link2). Explore to see if this is an issue.")
    
  }
  
  # Create definitions data frame in the global environment
  
  definitionsdf <<- definitionsdfx
  
  rm(definitionsdf2, definitionsdf3, definitionsdf4, definitionsdf5, definitionsdfx)
  
}

# definitionstab creates the worksheet with information on definitions
# If definitions not wanted, then do not run the definitionstab function
# There are three parameters and they are optional and preset. Change contentslink to "No" if you want a contents tab but do not want a link to it in the definitions tab. Change gridlines to "No" if gridlines are not wanted.
# Column widths are automatically set but the user can specify the required widths in colwid_spec
# Extra columns can be added by setting extracols to "Yes" and creating a dataframe extracols_definitions with the desired extra columns


definitionstab <- function(contentslink = NULL, gridlines = "Yes", colwid_spec = NULL, extracols = NULL) {
  
  # Checking that a definitions page is wanted, based on whether a worksheet was created in the initial workbook
  
  if (!("Definitions" %in% names(wb))) {
    
    stop("definitionstab has not been set to \"Yes\" in the workbook function call")
    
  }
  
  # Checking some parameters have been populated properly, if not the function will error
  
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
    
    stop("gridlines has not been populated properly. It must be a single word, either \"Yes\" or \"No\".")
    
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
    
    stop("extracols has not been populated properly. It must be a single word, either \"Yes\" or \"No\".")
    
  }
  
  # Automatically detecting if links have been provided or not
  
  definitionsdfx <- definitionsdf %>%
    dplyr::filter(!is.na(Link1) & Link1 != "No additional link")
  
  if (nrow(definitionsdfx) > 0) {
    
    links <- "Yes"
    
  } else if (nrow(definitionsdf) > 0 & nrow(definitionsdfx) == 0) {
    
    links <- "No"
    
  } else if (nrow(definitionsdf) == 0) {
    
    stop("The definitionsdf dataframe contains no observations")
    
  }
  
  rm(definitionsdfx)
  
  # Identifying the row number of definitions which have a link associated to them
  
  definitionsdfy <- definitionsdf %>%
    dplyr::mutate(linkno = dplyr::case_when(is.na(Link1) | Link1 == "No additional link" ~ NA_real_,
                                            TRUE ~ dplyr::row_number())) %>%
    dplyr::filter(!is.na(linkno))
  
  if (nrow(definitionsdfy) > 0) {
    
    linkrange <- definitionsdfy$linkno
    
  } else {
    
    linkrange <- NULL
    
  }
  
  rm(definitionsdfy)
  
  # Checking for any duplication of definitions
  
  if (nrow(definitionsdf) == 0) {stop("No rows in the definitions table")}
  
  definitionsdf2 <- definitionsdf %>%
    dplyr::group_by(Term) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(count) / n())
  
  definitionsdf3 <- definitionsdf %>%
    dplyr::group_by(Definition) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(count) / n())
  
  definitionsdf4 <- definitionsdf %>%
    dplyr::group_by(Link1) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link1) | Link1 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(count)) / n()) 
  
  definitionsdf5 <- definitionsdf %>%
    dplyr::group_by(Link2) %>%
    dplyr::summarise(count = n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link2) | Link2 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(count)) / n()) 
  
  if (definitionsdf2$check > 1) {
    
    stop("Duplicated term(s)")
    
  }
  
  if (definitionsdf3$check > 1) {
    
    warning("Duplicated definition(s). Explore to see if this is an issue.")
    
  }
  
  if (definitionsdf4$check > 1) {
    
    warning("Duplicated link text(s) (Link1). Explore to see if this is an issue.")
    
  }
  
  if (definitionsdf5$check > 1) {
    
    warning("Duplicated link text(s) (Link2). Explore to see if this is an issue.")
    
  }
  
  rm(definitionsdf2, definitionsdf3, definitionsdf4, definitionsdf5)
  
  if (links != "No" & links != "Yes") {
    
    stop("links not set to \"Yes\" or \"No\". There must be an issue with link information provided with the definitions.")
    
  }
  
  # Determining if a link to the contents page is wanted on the definitions worksheet
  
  if (length(contentslink) > 1) {
    
    stop("contentslink is not populated properly. It should be a single entity, either \"Yes\" or \"No\".")
    
  }
  
  if (is.null(contentslink)) {
    
    contentslink <- "Yes"
    
  } else if (tolower(contentslink) == "no" | tolower(contentslink) == "n") {
    
    contentslink <- "No"
    
  }
  
  if (contentslink == "No" | !("Contents" %in% names(wb))) {
    
    contentstab <- "No"
    
  } else if ("Contents" %in% names(wb)) {
    
    contentstab <- "Yes"
    
  } else {
    
    contentstab <- "No"
    
  }
  
  if (extracols == "Yes" & exists("extracols_definitions", envir = .GlobalEnv)) {
    
    if (nrow(extracols_definitions) != nrow(definitionsdf)) {
      
      stop("The number of rows in the definitions table and the extracols_definitions dataframe is not the same")
      
    }
    
  } else if (extracols == "No" & exists("extracols_definitions", envir = .GlobalEnv)) {
    
    warning("extracols has been set to \"No\" but a dataframe extracols_definitions exists. No extra columns have been added.")
    
  } else if (extracols == "Yes" & !(exists("extracols_definitions", envir = .GlobalEnv))) {
    
    warning("extracols has been set to \"Yes\" but a extracols_definitions dataframe does not exist. No extra columns will be added.")
    
  }
  
  if (links == "Yes") {
    
    definitionsdf <<- definitionsdf %>%
      dplyr::select("Term", "Definition", "Link") %>%
      {if (exists("extracols_definitions", envir = .GlobalEnv)) dplyr::bind_cols(., extracols_definitions) else .}
    
    class(definitionsdf$Link) <- "formula"
    
  } else if (links == "No") {
    
    definitionsdf <<- definitionsdf %>%
      dplyr::select("Term", "Definition") %>%
      {if (exists("extracols_definitions", envir = .GlobalEnv)) dplyr::bind_cols(., extracols_definitions) else .}
    
  }
  
  if ("Link" %in% colnames(definitionsdf)) {
    
    definitionsdfcols <- colnames(definitionsdf)
    
    for (i in seq_along(definitionsdfcols)) {
      
      if (definitionsdfcols[i] == "Link") {linkcolpos <- i}
      
    }
    
  }
  
  if (extracols == "Yes" & exists("extracols_definitions", envir = .GlobalEnv)) {
    
    if (any(duplicated(colnames(definitionsdf))) == TRUE) {
      
      warning("There is at least one duplicate column name in the definitions table and the extracols_definitions dataframe")
      
    }
    
  }
  
  # Define some formatting for use later on
  
  normalformat <- openxlsx::createStyle(valign = "top")
  topformat <- openxlsx::createStyle(valign = "bottom")
  linkformat <- openxlsx::createStyle(fontColour = "blue", textDecoration = "underline", valign = "top")
  
  openxlsx::addStyle(wb, "Definitions", normalformat, rows = 1:(nrow(definitionsdf) + 4), cols = 1:ncol(definitionsdf), gridExpand = TRUE)
  
  openxlsx::writeData(wb, "Definitions", "Definitions", startCol = 1, startRow = 1)
  
  titleformat <- openxlsx::createStyle(fontSize = fontszt, textDecoration = "bold")
  
  openxlsx::addStyle(wb, "Definitions", titleformat, rows = 1, cols = 1)
  
  openxlsx::writeData(wb, "Definitions", "This worksheet contains one table.", startCol = 1, startRow = 2)
  openxlsx::addStyle(wb, "Definitions", topformat, rows = 2, cols = 1)
  
  extraformat <- openxlsx::createStyle(valign = "top")
  
  if (contentstab == "Yes") {
    
    openxlsx::addStyle(wb, "Definitions", linkformat, rows = 3, cols = 1)
    
    openxlsx::writeFormula(wb, "Definitions", startRow = 3, x = openxlsx::makeHyperlinkString("Contents", row = 1, col = 1, text = "Link to contents"))
    
    startingrow <- 4
    
  } else if (contentstab == "No") {
    
    startingrow <- 3
    
    openxlsx::addStyle(wb, "Definitions", extraformat, rows = 2, cols = 1, stack = TRUE)
    
  }
  
  if (links == "Yes" & is.null(linkrange)) {
    
    stop("Links are required in the definitions tab but the row numbers where links should be have not been generated (i.e., linkrange not populated)")
    
  } else if (links == "Yes" & !is.null(linkrange)) {
    
    openxlsx::addStyle(wb, "Definitions", linkformat, rows = linkrange + startingrow, cols = linkcolpos, gridExpand = TRUE)
    
  }
  
  openxlsx::writeDataTable(wb, "Definitions", definitionsdf, tableName = "definitions", startRow = startingrow, startCol = 1, withFilter = FALSE, tableStyle = "none")
  
  if ("Link" %in% colnames(definitionsdf)) {
    
    noaddlinks <- definitionsdf[["Link"]]
    
    for (i in seq_along(noaddlinks)) {
      
      if (noaddlinks[i] == "No additional link") {
        
        openxlsx::writeData(wb, "Definitions", "No additional link", startCol = linkcolpos, startRow = startingrow + i)
        
      } 
      
    }
    
  }
  
  headingsformat <- openxlsx::createStyle(textDecoration = "bold", wrapText = TRUE, border = NULL, valign = "top")
  
  openxlsx::addStyle(wb, "Definitions", headingsformat, rows = startingrow, cols = 1:ncol(definitionsdf))
  
  extraformat2 <- openxlsx::createStyle(wrapText = TRUE)
  
  openxlsx::addStyle(wb, "Definitions", extraformat2, rows = (startingrow + 1):(nrow(definitionsdf) + startingrow + 1), cols = 1:2, stack = TRUE, gridExpand = TRUE)
  
  if ((!is.null(colwid_spec) & !is.numeric(colwid_spec)) | (!is.null(colwid_spec) & length(colwid_spec) != ncol(definitionsdf))) {
    
    warning("colwid_spec is either a non-numeric value or a vector not of the same length as the number of columns desired in the definitions tab. The widths will be determined automatically.")
    colwid_spec <- NULL
    
  }
  
  numchars1 <- max(nchar(definitionsdf$"Term")) + 10
  numchars2 <- max(nchar(definitionsdf$"Definition")) + 10
  
  if (links == "Yes" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Definitions", cols = c(1,2,3,4:max(ncol(definitionsdf),4)), widths = c(min(numchars1, 30), min(numchars2, 75), "auto", "auto"))
    
  } else if (links == "No" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Definitions", cols = c(1,2,3:max(ncol(definitionsdf),3)), widths = c(min(numchars1, 30), min(numchars2, 75), "auto"))
    
  } else if (!is.null(colwid_spec) & is.numeric(colwid_spec) & length(colwid_spec) == ncol(definitionsdf)) {
    
    openxlsx::setColWidths(wb, "Definitions", cols = c(1:ncol(definitionsdf)), widths = colwid_spec)
    
  }
  
  openxlsx::setRowHeights(wb, "Definitions", startingrow - 1, fontsz * (25/12))
  
  rm(definitionsdf, envir = .GlobalEnv)
  
  # If gridlines are not wanted, then they are removed
  
  if (gridlines == "No") {
    
    openxlsx::showGridLines(wb, "Definitions", showGridLines = FALSE)
    
  }
  
}

###################################################################################################################
###################################################################################################################
# SAVING THE FINAL SPREADSHEET

# The savingtables function only requires that the location and name of the spreadsheet be specified
# Cannot save directly to ODS, unfortunately this would need to be done manually


savingtables <- function(filename) {
  
  if (length(filename) > 1) {
    
    stop("filename is not populated properly. It should be a single entity and not a vector.")
    
  }
  
  if (stringr::str_detect(filename, " ")) {
    
    warning("GSS guidance for spreadsheets includes not using spaces in file names, instead consider using dashes")
    
  }
  
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
  
  tablestarts <- unlist(mget(paste0(sheetnames2, "_tablestart"), envir = .GlobalEnv))
  
  if (length(unique(tablestarts)) > 1) {
    
    warning("The row number of the data table headings is not the same on all worksheets. This might be frustrating for anyone reading the tables into a programming language. Consider whether it would be possible to make the tables start on the same row of each worksheet.")
    
  }
  
  rm(sheetnames, sheetnames2, tablestarts)
  
  openxlsx::saveWorkbook(wb, filename, overwrite = TRUE)
  
  print("***************** It is not currently possible to save the workbook as an ODS file via this code. For accessibility reasons, consider manually converting the workbook to an ODS file. *****************")
  
  # Remove data frames and variables from the global environment in case accessible tables needs to be run again
  
  rm(wb, envir = .GlobalEnv)
  
  if (exists("notesdf", envir = .GlobalEnv)) {
    
    rm(notesdf, envir = .GlobalEnv)
    
  }
  
  if (exists("tabcontents", envir = .GlobalEnv)) {
    
    rm(tabcontents, envir = .GlobalEnv)
    
  }
  
  if (exists("definitionsdf", envir = .GlobalEnv)) {
    
    rm(definitionsdf, envir = .GlobalEnv)
    
  }
  
  rm(list = ls(pattern = "_startrow", envir = .GlobalEnv), envir = .GlobalEnv)
  rm(list = ls(pattern = "_tablestart", envir = .GlobalEnv), envir = .GlobalEnv)
  
  if (exists("autonotes2", envir = .GlobalEnv)) {
    
    rm(autonotes2, envir = .GlobalEnv)
    
  }
  
  rm(fontsz, fontszst, fontszt, envir = .GlobalEnv)
  
  if (exists("covernumrow", envir = .GlobalEnv)) {
    
    rm(covernumrow, envir = .GlobalEnv)
    
  }
  
}

###################################################################################################################