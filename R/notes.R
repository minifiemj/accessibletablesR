###################################################################################################
# NOTES

#' @title accessibletablesR::addnote
#'
#' @description Add a note to a particular table or tables
#' 
#' @details 
#' addnote function will add a note and its description to the workbook, specifically in the notes 
#' worksheet.
#' Add notes if wanted, if not then do not run the addnote function.
#' A link can be provided with each note as well a list of tables that the note applies to.
#' notenumber and notetext are the only compulsory parameters.
#' All other parameters are optional and preset to NULL, so only need to be defined if they are 
#' wanted.
#' applictabtext should be set to a vector of sheet names if a column is wanted which lists which 
#' worksheets a note is applicable to.
#' linktext1 and linktext2: linktext1 should be the text you want to appear and linktext2 should be 
#' the underlying link to a website, file etc.
#' 
#' @param notenumber Note number
#' @param notetext Note description
#' @param applictabtext Table(s) a note is applicable to (optional)
#' @param linktext1 Text to appear in place of a link associated to a note (optional)
#' @param linktext2 Link associated to a note (optional)
#' 
#' @returns A dataframe containing all information associated to notes
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

addnote <- function(notenumber, notetext, applictabtext = NULL, linktext1 = NULL, 
                    linktext2 = NULL) {
  
  if (!("dplyr" %in% utils::installed.packages()) |
      !("stringr" %in% utils::installed.packages()) |
      !("conflicted" %in% utils::installed.packages()) |
      !("rlang" %in% utils::installed.packages())) {
    
    stop(base::strwrap("Not all required packages installed. Run the \"workbook\" function first to 
         ensure packages are installed.", prefix = " ", initial = ""))
    
  } else if (utils::packageVersion("dplyr") < "1.1.2" |
             utils::packageVersion("stringr") < "1.5.0" |
             utils::packageVersion("conflicted") < "1.2.0" |
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
  tabcontents <- acctabs$tabcontents
  notesdf <- acctabs$notesdf
  
  # Checking that a notes page is wanted, based on whether a worksheet was created in the ...
  # ... initial workbook
  
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
    
    stop(strwrap("One or both of notenumber and notetext are not populated properly. They must be a 
         single entity and not a vector.", prefix = " ", initial = ""))
    
  }
  
  if (!is.null(applictabtext) & !is.character(applictabtext)) {
    
    stop(strwrap("The parameter applictabtext is not populated properly. If it is not NULL then it 
         has to be a string. If more than one element is needed then it should be expressed as a 
         vector e.g., applictabtext = c(\"Table_1\", \"Table_2\")", prefix = " ", initial = ""))
    
  }
  
  if (is.null(applictabtext) & autonotes2 == "Yes") {
    
    stop(strwrap("Automatic listing of notes on tables has been selected but a note has no tables 
         applicable to it", prefix = " ", initial = ""))
    
  }
  
  if (length(linktext1) > 1 | length(linktext2) > 1) {
    
    stop(strwrap("linktext1 and linktext2 can only be single entities and not vectors of length 
         greater than one", prefix = " ", initial = ""))
    
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
    
    stop(strwrap("The notenumber parameter is not properly populated. It should take the form of 
         \"note\" followed by a number.", prefix = " ", initial = ""))
    
  }
  
  rm(notetemp1, notetemp2)
  
  if (!is.null(applictabtext)) {
    
    check <- 0
    
    for (i in seq_along(applictabtext)) {
      
      if (tolower(applictabtext[i]) == "all") {applictabtext[i] <- "All"}
      
      if (length(applictabtext) > 1 & applictabtext[i] == "All") {
        
        stop(strwrap("The applictabtext parameter includes two or more elements but one of the 
             elements is \"All\"", prefix = " ", initial = ""))
        
      }
      
      if (stringr::str_detect(applictabtext[i], " ") | stringr::str_detect(applictabtext[i], ",")) {
        
        stop(strwrap("The applictabtext contains whitespace or a comma. applictabtext should either 
             be a single word (e.g., \"All\") or expressed as a vector (e.g., c(\"Table_1\", 
             \"Table_2\"))", prefix = " ", initial = ""))
        
      }
      
      if (tolower(applictabtext[i]) == "none") {
        
        stop(strwrap("A note should be applicable to at least one of the tables. applictabtext 
             should not be set to \"None\".", prefix = " ", initial = ""))
        
      }
      
      if (!(applictabtext[i] %in% tabcontents[[1]]) & applictabtext[i] != "All") {
        
        print(paste0(applictabtext[i], " not in table of contents"))
        check <- 1
        
      }
      
    }
    
    if (check == 1) {
      
      stop(strwrap("At least one of the tables mentioned in the applictabtext parameter is not in 
           the table of contents", prefix = " ", initial = ""))
      
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
    dplyr::add_row("Note number" = notenumber, "Note text" = notetext, 
                   "Applicable tables" = applictabtext2, "Link1" = linktext1, 
                   "Link2" = linktext2) %>%
    dplyr::mutate(Link2 = dplyr::case_when(Link1 == "No additional link" ~ "No additional link",
                                           TRUE ~ Link2)) %>%
    dplyr::mutate(Link = 
                    dplyr::case_when(Link1 == "No additional link" ~ "No additional link",
                                     TRUE ~ paste0("HYPERLINK(\"", Link2, "\", \"", Link1, "\")")))
  
  notesdf2 <- notesdfx %>%
    dplyr::rename(note_number = "Note number") %>%
    dplyr::group_by(.data$note_number) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(.data$count) / dplyr::n())
  
  notesdf3 <- notesdfx %>%
    dplyr::rename(note_text = "Note text") %>%
    dplyr::group_by(.data$note_text) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(.data$count) / dplyr::n())
  
  notesdf4 <- notesdfx %>%
    dplyr::group_by(.data$Link1) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link1) | Link1 == "" | 
                                             Link1 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(.data$count)) / dplyr::n()) 
  
  notesdf5 <- notesdfx %>%
    dplyr::group_by(.data$Link2) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link2) | Link2 == "" | 
                                             Link2 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(.data$count)) / dplyr::n()) 
  
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
  
  notesdfx_temp <- notesdfx
  
  assign("notesdf", notesdfx_temp, envir = as.environment(acctabs))
  
  rm(notesdf2, notesdf3, notesdf4, notesdf5, notesdfx, notesdfx_temp)
  
}

#' @title accessibletablesR::notestab
#'
#' @description Create a notes page for the workbook.
#' 
#' @details 
#' notestab function will create a notes worksheet in the workbook and includes notes added using 
#' the addnote function.
#' If notes not wanted, then do not run the notestab function.
#' There are three parameters and they are optional and preset. Change contentslink to "No" if you 
#' want a contents tab but do not want a link to it in the notes tab. Change gridlines to "No" if 
#' gridlines are not wanted.
#' Column widths are automatically set but the user can specify the required widths in colwid_spec.
#' Extra columns can be added by setting extracols to "Yes" and creating a dataframe 
#' extracols_notes with the desired extra columns.
#' 
#' @param contentslink Define whether a link to the contents page is wanted (optional)
#' @param gridlines Define whether gridlines are present (optional)
#' @param colwid_spec Define widths of columns (optional)
#' @param extracols Define whether additional columns required (optional)
#' 
#' @returns A worksheet of the notes page for the workbook.
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

notestab <- function(contentslink = NULL, gridlines = "Yes", colwid_spec = NULL, extracols = NULL) {
  
  if (!("dplyr" %in% utils::installed.packages()) | 
      !("openxlsx" %in% utils::installed.packages()) |
      !("conflicted" %in% utils::installed.packages()) |
      !("stringr" %in% utils::installed.packages()) |
      !("rlang" %in% utils::installed.packages())) {
    
    stop(base::strwrap("Not all required packages installed. Run the \"workbook\" function first to 
         ensure packages are installed.", prefix = " ", initial = ""))
    
  } else if (utils::packageVersion("dplyr") < "1.1.2" | 
             utils::packageVersion("openxlsx") < "4.2.5.2" |
             utils::packageVersion("conflicted") < "1.2.0" |
             utils::packageVersion("stringr") < "1.5.0" |
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
  notesdf <- acctabs$notesdf
  autonotes2 <- acctabs$autonotes2
  fontszt <- acctabs$fontszt
  fontsz <- acctabs$fontsz
  tabcontents <- acctabs$tabcontents
  
  # Check that a notes page is wanted, based on whether a worksheet was created in the initial ...
  # ... workbook
  
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
  
  # Automatically detecting if information of which tables notes apply to has been provided or not
  
  notesdfx <- notesdf %>%
    dplyr::rename(applictab = "Applicable tables") %>%
    dplyr::filter(.data$applictab != "")
  
  if (nrow(notesdf) > 0 & nrow(notesdf) == nrow(notesdfx)) {
    
    applictabs <- "Yes"
    
  } else if (nrow(notesdf) > 0 & nrow(notesdfx) == 0) {
    
    applictabs <- "No"
    
  } else if (nrow(notesdf) > 0 & nrow(notesdf) != nrow(notesdfx)) {
    
    stop(strwrap("There may be a note without applicable tables allocated to it while other notes do 
         have applicable tables allocated to them", prefix = " ", initial = ""))
    
  } else if (nrow(notesdf) == 0) {
    
    stop("The notesdf dataframe contains no observations")
    
  }
  
  rm(notesdfx)
  
  # Automatically detecting if links have been provided or not
  
  notesdfx <- notesdf %>%
    dplyr::filter(!is.na(.data$Link1) & .data$Link1 != "No additional link")
  
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
    dplyr::filter(!is.na(.data$linkno))
  
  if (nrow(notesdfy) > 0) {
    
    linkrange <- notesdfy$linkno
    
  } else {
    
    linkrange <- NULL
    
  }
  
  rm(notesdfy)
  
  # Checks associated with the automatic generation of note information for the main data tables
  
  if (applictabs == "Yes" & autonotes2 != "Yes") {
    
    warning(strwrap("The applictabs parameter has been set to \"Yes\" but the automatic listing of 
            notes on a worksheet has not been selected", prefix = " ", initial = ""))
    
  } else if (applictabs == "No" & autonotes2 == "Yes") {
    
    stop(strwrap("The applictabs parameter has been set to \"No\" but the automatic listing of notes 
         on a worksheet has been selected", prefix = " ", initial = ""))
    
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
    dplyr::group_by(.data$note_number) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(.data$count) / dplyr::n())
  
  notesdf3 <- notesdf %>%
    dplyr::rename(note_text = "Note text") %>%
    dplyr::group_by(.data$note_text) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(.data$count) / dplyr::n())
  
  notesdf4 <- notesdf %>%
    dplyr::group_by(.data$Link1) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link1) | Link1 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(.data$count)) / dplyr::n()) 
  
  notesdf5 <- notesdf %>%
    dplyr::group_by(.data$Link2) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link2) | Link2 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(.data$count)) / dplyr::n()) 
  
  notesdf6 <- notesdf %>%
    dplyr::rename(applictab = "Applicable tables") %>%
    dplyr::filter(is.na(.data$applictab))
  
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
    
    stop(strwrap("links not set to \"Yes\" or \"No\". There must be an issue with link information 
         provided with the notes.", prefix = " ", initial = ""))
    
  }
  
  if (applictabs != "No" & applictabs != "Yes") {
    
    stop("applictabs not set to \"Yes\" or \"No\"")
    
  } else if (applictabs == "Yes" & nrow(notesdf6) > 0) {
    
    stop("Applicable table column wanted but contains an empty cell or cells")
    
  }
  
  rm(notesdf6)
  
  # Creating a notes table with the required columns
  
  if (extracols == "Yes" & exists("extracols_notes", envir = .GlobalEnv)) {
    
    extracols_notes <- get("extracols_notes", envir = .GlobalEnv)
    
    if (nrow(extracols_notes) != nrow(notesdf)) {
      
      stop(strwrap("The number of rows in the notes table is not the same as in the dataframe of 
           extra columns", prefix = " ", initial = ""))
      
    }
    
  } else if (extracols == "Yes" & !(exists("extracols_notes", envir = .GlobalEnv))) {
    
    warning(strwrap("extracols has been set to \"Yes\" but the extracols_notes dataframe does not 
            exist. No extra columns will be added.", prefix = " ", initial = ""))
    
  } else if (extracols == "No" & exists("extracols_notes", envir = .GlobalEnv)) {
    
    warning(strwrap("extracols has been set to \"No\" but a dataframe extracols_notes exists. No 
            extra columns have been added.", prefix = " ", initial = ""))
    
  }
  
  if (links == "Yes" & applictabs == "Yes") {
    
    notesdf_temp <- notesdf %>%
      dplyr::select("Note number", "Note text", "Applicable tables", "Link") %>%
      {if (exists("extracols_notes", envir = .GlobalEnv) & extracols == "Yes") 
        dplyr::bind_cols(., extracols_notes) else .}
    
    class(notesdf_temp$Link) <- "formula"
    
    assign("notesdf", notesdf_temp, envir = as.environment(acctabs))
    rm(notesdf_temp)
    
  } else if (links == "Yes") {
    
    notesdf_temp <- notesdf %>%
      dplyr::select("Note number", "Note text", "Link") %>%
      {if (exists("extracols_notes", envir = .GlobalEnv) & extracols == "Yes") 
        dplyr::bind_cols(., extracols_notes) else .}
    
    class(notesdf_temp$Link) <- "formula"
    
    assign("notesdf", notesdf_temp, envir = as.environment(acctabs))
    rm(notesdf_temp)
    
  } else if (links == "No" & applictabs == "Yes") {
    
    notesdf_temp <- notesdf %>%
      dplyr::select("Note number", "Note text", "Applicable tables") %>%
      {if (exists("extracols_notes", envir = .GlobalEnv) & extracols == "Yes") 
        dplyr::bind_cols(., extracols_notes) else .}
    
    assign("notesdf", notesdf_temp, envir = as.environment(acctabs))
    rm(notesdf_temp)
    
  } else if (links == "No") {
    
    notesdf_temp <- notesdf %>%
      dplyr::select("Note number", "Note text") %>%
      {if (exists("extracols_notes", envir = .GlobalEnv) & extracols == "Yes") 
        dplyr::bind_cols(., extracols_notes) else .}
    
    assign("notesdf", notesdf_temp, envir = as.environment(acctabs))
    rm(notesdf_temp)
    
  }
  
  notesdf <- acctabs$notesdf
  
  if ("Link" %in% colnames(notesdf)) {
    
    notesdfcols <- colnames(notesdf)
    
    for (i in seq_along(notesdfcols)) {
      
      if (notesdfcols[i] == "Link") {linkcolpos <- i}
      
    }
    
  }
  
  if (extracols == "Yes" & exists("extracols_notes", envir = .GlobalEnv)) {
    
    if (any(duplicated(colnames(notesdf))) == TRUE) {
      
      warning(strwrap("There is at least one duplicate column name in the notes table and the 
              extracols_notes dataframe", prefix = " ", initial = ""))
      
    }
    
  } 
  
  # Define formatting to be used later on
  
  normalformat <- openxlsx::createStyle(valign = "top")
  topformat <- openxlsx::createStyle(valign = "bottom")
  linkformat <- openxlsx::createStyle(fontColour = "blue", valign = "top", 
                                      textDecoration = "underline")
  
  openxlsx::addStyle(wb, "Notes", normalformat, rows = 1:(nrow(notesdf) + 4), 
                     cols = 1:ncol(notesdf), gridExpand = TRUE)
  
  openxlsx::writeData(wb, "Notes", "Notes", startCol = 1, startRow = 1)
  
  titleformat <- openxlsx::createStyle(fontSize = fontszt, textDecoration = "bold")
  
  openxlsx::addStyle(wb, "Notes", titleformat, rows = 1, cols = 1)
  
  openxlsx::writeData(wb, "Notes", "This worksheet contains one table.", startCol = 1, startRow = 2)
  
  openxlsx::addStyle(wb, "Notes", topformat, rows = 2, cols = 1)
  
  extraformat <- openxlsx::createStyle(valign = "top")
  
  if (contentstab == "Yes") {
    
    openxlsx::writeFormula(wb, "Notes", startRow = 3, 
                           x = openxlsx::makeHyperlinkString("Contents", row = 1, col = 1, 
                                                             text = "Link to contents"))
    
    openxlsx::addStyle(wb, "Notes", linkformat, rows = 3, cols = 1)
    
    startingrow <- 4
    
  } else if (contentstab == "No") {
    
    startingrow <- 3
    
    openxlsx::addStyle(wb, "Notes", extraformat, rows = 2, cols = 1, stack = TRUE)
    
  }
  
  if (links == "Yes" & is.null(linkrange)) {
    
    stop(strwrap("Links are required in the notes tab but the row numbers where links should be have 
         not been generated (i.e., linkrange not populated)", prefix = " ", initial = ""))
    
  } else if (links == "Yes" & !is.null(linkrange)) {
    
    openxlsx::addStyle(wb, "Notes", linkformat, rows = linkrange + startingrow, cols = linkcolpos, 
                       gridExpand = TRUE)
    
  }
  
  openxlsx::writeDataTable(wb, "Notes", notesdf, tableName = "notes", startRow = startingrow, 
                           startCol = 1, withFilter = FALSE, tableStyle = "none")
  
  # The if statement below is required so that "No additional link" appears as text only in the ...
  # ... final spreadsheet, rather than as a hyperlink
  
  if ("Link" %in% colnames(notesdf)) {
    
    noaddlinks <- notesdf[["Link"]]
    
    for (i in seq_along(noaddlinks)) {
      
      if (noaddlinks[i] == "No additional link") {
        
        openxlsx::writeData(wb, "Notes", "No additional link", startCol = linkcolpos, 
                            startRow = startingrow + i)
        
      } 
      
    }
    
  }
  
  headingsformat <- openxlsx::createStyle(textDecoration = "bold", wrapText = TRUE, border = NULL, 
                                          valign = "top")
  
  openxlsx::addStyle(wb, "Notes", headingsformat, rows = startingrow, cols = 1:ncol(notesdf))
  
  extraformat2 <- openxlsx::createStyle(wrapText = TRUE)
  
  openxlsx::addStyle(wb, "Notes", extraformat2, 
                     rows = (startingrow + 1):(nrow(notesdf) + startingrow + 1), 
                     cols = 1:ncol(notesdf), stack = TRUE, gridExpand = TRUE)
  
  # Determining column widths
  
  if ((!is.null(colwid_spec) & !is.numeric(colwid_spec)) | 
      (!is.null(colwid_spec) & length(colwid_spec) != ncol(notesdf))) {
    
    warning(strwrap("colwid_spec is either a non-numeric value or a vector not of the same length as 
            the number of columns desired in the notes tab. The widths will be determined 
            automatically.", prefix = " ", initial = ""))
    colwid_spec <- NULL
    
  }
  
  numchars <- max(nchar(notesdf$"Note text")) + 10
  
  if (links == "Yes" & applictabs == "No" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Notes", cols = c(1,2,3,4:max(ncol(notesdf),4)), 
                           widths = c(15, min(numchars, 100), 50, "auto"))
    
  } else if (links == "Yes" & applictabs == "Yes" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Notes", cols = c(1,2,3,4,5:max(ncol(notesdf),5)), 
                           widths = c(15, min(numchars, 100), 20, 50, "auto"))
    
  } else if (links == "No" & applictabs == "Yes" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Notes", cols = c(1,2,3,4:max(ncol(notesdf),4)), 
                           widths = c(15, min(numchars, 100), 20, "auto"))
    
  } else if (links == "No" & applictabs == "No" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Notes", cols = c(1,2,3:max(ncol(notesdf),3)), 
                           widths = c(15, min(numchars, 100), "auto"))
    
  } else if (!is.null(colwid_spec) & is.numeric(colwid_spec) & 
             length(colwid_spec) == ncol(notesdf)) {
    
    openxlsx::setColWidths(wb, "Notes", cols = c(1:ncol(notesdf)), widths = colwid_spec)
    
  }
  
  openxlsx::setRowHeights(wb, "Notes", startingrow - 1, fontsz * (25/12))
  
  # Creating the text to be inserted in the main data tables regarding which notes are ...
  # ... associated with which table
  
  if (applictabs == "Yes" & autonotes2 == "Yes") {
    
    tabcontents2 <- tabcontents %>%
      dplyr::filter(.[[1]] != "Notes" & .[[1]] != "Definitions")
    
    tablelist <- tabcontents2[[1]]
    
    for (i in seq_along(tablelist)) {
      
      notesdf7 <- notesdf %>%
        dplyr::rename(applic_tab = "Applicable tables") %>%
        dplyr::mutate(applic_tab = 
                        dplyr::case_when(applic_tab == "All" ~ paste(tablelist, collapse = ", "),
                                         TRUE ~ applic_tab)) %>%
        dplyr::filter(stringr::str_detect(.data$applic_tab, tablelist[i]) == TRUE)
      
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
        warning(strwrap(paste0(tablelist[i], " has no notes associated with it. Check that this is 
                intentional."), prefix = " ", initial = ""))
        
      } else {
        
        notes7 <- paste0("This worksheet contains one table. For notes, see ", notes6, 
                         " on the notes worksheet.")
        
      }
      
      tempstartrow <- get(paste0(tablelist[i], "_startrow"), envir = as.environment(acctabs))
      
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
      dplyr::summarise(applic_tab4 = sum(.data$applic_tab3))
    
    if (notesdf8$applic_tab4 >= 1) {
      
      warning(strwrap("There is at least one occurrence where the list of applicable tables appears 
              to be all of the tables. The list could read \"All\" instead.", prefix = " ",
                      initial = ""))
      
    }
    
    notesdf9 <- notesdf %>%
      dplyr::rename(applic_tab = "Applicable tables") %>%
      dplyr::filter(.data$applic_tab != "All")
    
    applictablist <- notesdf9$applic_tab
    
    applictablist2 <- paste(applictablist, collapse = ", ")
    
    applictablist3 <- unique(unlist(strsplit(applictablist2, ", ")))
    
    for (i in seq_along(applictablist3)) {
      
      if(!(applictablist3[i] %in% tabcontents2[[1]])) {
        
        stop(paste0(applictablist3), " is not in the table of contents")
        
      }
      
    }
    
    rm(tablelist, tablelist2, notesdf8, tabcontents2, notesdf9, applictablist, applictablist2, 
       applictablist3)
    
  }
  
  rm(list = ls(pattern = "startrow", envir = as.environment(acctabs)), 
     envir = as.environment(acctabs))
  
  rm(notesdf, envir = as.environment(acctabs))
  
  # If gridlines are not wanted, then they are removed
  
  if (gridlines == "No") {
    
    openxlsx::showGridLines(wb, "Notes", showGridLines = FALSE)
    
  }
  
  assign("wb", wb, envir = as.environment(acctabs))
  
}

###################################################################################################