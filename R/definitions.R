###################################################################################################
# DEFINITIONS

#' @title accessibletablesR::adddefinition
#'
#' @description Add a definition of a term relevant to the workbook.
#' 
#' @details 
#' adddefinition function will add a definition and its description to the workbook, specifically 
#' in the definitions worksheet.
#' Add definitions if wanted, if not then do run the adddefinition function.
#' term and definition are compulsory parameters.
#' A link can be added with each definition.
#' linktext1 and linktext2: linktext1 should be the text you want to appear and linktext2 should be 
#' the underlying link to a website, file etc.
#' 
#' @param term Term to be defined
#' @param definition Definition of term
#' @param linktext1 Text to appear in place of link associated with definition (optional)
#' @param linktext2 Link associated with definition (optional)
#' 
#' @returns A dataframe containing all information associated to definitions.
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
#'    othdatacols = c(9,10), datedatacols = 15, datedatafmt = "dd-mm-yyyy", 
#'    datenondatacols = 14, datenondatafmt = "yyyy-mm-dd", columnwidths = "specified",
#'    colwid_spec = c(18,18,18,15,17,15,12,17,12,13,23,22,12,12,12))
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

adddefinition <- function(term, definition, linktext1 = NULL, linktext2 = NULL) {
  
  if (!("dplyr" %in% utils::installed.packages()) |
      !("conflicted" %in% utils::installed.packages()) |
      !("rlang" %in% utils::installed.packages())) {
    
    stop(base::strwrap("Not all required packages installed. Run the \"workbook\" function first to 
         ensure packages are installed.", prefix = " ", initial = ""))
    
  } else if (utils::packageVersion("dplyr") < "1.1.2" |
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
  definitionsdf <- acctabs$definitionsdf
  
  # Checking that a definitions page is wanted, based on whether a worksheet was created in the ...
  # ... initial workbook
  
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
    
    stop(strwrap("The parameter term is not populated properly. If it is not NULL then it has to be 
         a string.", prefix = " ", initial = ""))
    
  }
  
  if (!is.null(definition) & !is.character(definition)) {
    
    stop(strwrap("The parameter definition is not populated properly. If it is not NULL then it has 
         to be a string.", prefix = " ", initial = ""))
    
  }
  
  if (length(linktext1) > 1 | length(linktext2) > 1) {
    
    stop(strwrap("linktext1 and linktext2 can only be single entities and not vectors of length 
         greater than one", prefix = " ", initial = ""))
    
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
    dplyr::add_row("Term" = term, "Definition" = definition, "Link1" = linktext1, 
                   "Link2" = linktext2) %>%
    dplyr::mutate(Link2 = dplyr::case_when(Link1 == "No additional link" ~ "No additional link",
                                           TRUE ~ Link2)) %>%
    dplyr::mutate(Link = 
                    dplyr::case_when(Link1 == "No additional link" ~ "No additional link",
                                     TRUE ~ paste0("HYPERLINK(\"", Link2, "\", \"", Link1, "\")")))
  
  definitionsdf2 <- definitionsdfx %>%
    dplyr::group_by(.data$Term) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(.data$count) / dplyr::n())
  
  definitionsdf3 <- definitionsdfx %>%
    dplyr::group_by(.data$Definition) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(.data$count) / dplyr::n())
  
  definitionsdf4 <- definitionsdfx %>%
    dplyr::group_by(.data$Link1) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link1) | Link1 == "" | 
                                             Link1 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(.data$count)) / dplyr::n()) 
  
  definitionsdf5 <- definitionsdfx %>%
    dplyr::group_by(.data$Link2) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link2) | Link2 == "" | 
                                             Link2 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(.data$count)) / dplyr::n()) 
  
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
  
  definitionsdfx_temp <- definitionsdfx
  
  assign("definitionsdf", definitionsdfx_temp, envir = as.environment(acctabs))
  
  rm(definitionsdf2, definitionsdf3, definitionsdf4, definitionsdf5, definitionsdfx, 
     definitionsdfx_temp)
  
}

#' @title accessibletablesR::definitionstab
#' 
#' @description Create a definitions page for the workbook.
#' 
#' @details 
#' definitionstab creates the worksheet with information on definitions.
#' If definitions not wanted, then do not run the definitionstab function.
#' There are three parameters and they are optional and preset. Change contentslink to "No" if you 
#' want a contents tab but do not want a link to it in the definitions tab. Change gridlines to 
#' "No" if gridlines are not wanted.
#' Column widths are automatically set but the user can specify the required widths in colwid_spec.
#' Extra columns can be added by setting extracols to "Yes" and creating a dataframe 
#' extracols_definitions with the desired extra columns.
#' 
#' @param contentslink Define whether a link to the contents page is wanted (optional)
#' @param gridlines Define whether gridlines are present (optional)
#' @param colwid_spec Define widths of columns (optional)
#' @param extracols Define whether additional columns required (optional)
#' 
#' @returns A worksheet of the definitions page for the workbook.
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
#'    othdatacols = c(9,10), datedatacols = 15, datedatafmt = "dd-mm-yyyy", 
#'    datenondatacols = 14, datenondatafmt = "yyyy-mm-dd", columnwidths = "specified",
#'    colwid_spec = c(18,18,18,15,17,15,12,17,12,13,23,22,12,12,12))
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

definitionstab <- function(contentslink = NULL, gridlines = "Yes", colwid_spec = NULL, 
                           extracols = NULL) {
  
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
  definitionsdf <- acctabs$definitionsdf
  fontszt <- acctabs$fontszt
  fontsz <- acctabs$fontsz
  
  # Checking that a definitions page is wanted, based on whether a worksheet was created in the ...
  # ... initial workbook
  
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
  
  # Automatically detecting if links have been provided or not
  
  definitionsdfx <- definitionsdf %>%
    dplyr::filter(!is.na(.data$Link1) & .data$Link1 != "No additional link")
  
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
    dplyr::filter(!is.na(.data$linkno))
  
  if (nrow(definitionsdfy) > 0) {
    
    linkrange <- definitionsdfy$linkno
    
  } else {
    
    linkrange <- NULL
    
  }
  
  rm(definitionsdfy)
  
  # Checking for any duplication of definitions
  
  if (nrow(definitionsdf) == 0) {stop("No rows in the definitions table")}
  
  definitionsdf2 <- definitionsdf %>%
    dplyr::group_by(.data$Term) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(.data$count) / dplyr::n())
  
  definitionsdf3 <- definitionsdf %>%
    dplyr::group_by(.data$Definition) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::summarise(check = sum(.data$count) / dplyr::n())
  
  definitionsdf4 <- definitionsdf %>%
    dplyr::group_by(.data$Link1) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link1) | Link1 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(.data$count)) / dplyr::n()) 
  
  definitionsdf5 <- definitionsdf %>%
    dplyr::group_by(.data$Link2) %>%
    dplyr::summarise(count = dplyr::n()) %>%
    dplyr::ungroup() %>%
    dplyr::mutate(count = dplyr::case_when(is.na(Link2) | Link2 == "No additional link" ~ 1,
                                           TRUE ~ count)) %>%
    dplyr::summarise(check = sum(as.numeric(.data$count)) / dplyr::n()) 
  
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
    
    stop(strwrap("links not set to \"Yes\" or \"No\". There must be an issue with link information 
         provided with the definitions.", prefix = " ", initial = ""))
    
  }
  
  # Determining if a link to the contents page is wanted on the definitions worksheet
  
  if (length(contentslink) > 1) {
    
    stop(strwrap("contentslink is not populated properly. It should be a single entity, either 
         \"Yes\" or \"No\".", prefix = " ", initial = ""))
    
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
    
    extracols_definitions <- get("extracols_definitions", envir = .GlobalEnv)
    
    if (nrow(extracols_definitions) != nrow(definitionsdf)) {
      
      stop(strwrap("The number of rows in the definitions table and the extracols_definitions 
           dataframe is not the same", prefix = " ", initial = ""))
      
    }
    
  } else if (extracols == "No" & exists("extracols_definitions", envir = .GlobalEnv)) {
    
    warning(strwrap("extracols has been set to \"No\" but a dataframe extracols_definitions exists. 
            No extra columns have been added.", prefix = " ", initial = ""))
    
  } else if (extracols == "Yes" & !(exists("extracols_definitions", envir = .GlobalEnv))) {
    
    warning(strwrap("extracols has been set to \"Yes\" but a extracols_definitions dataframe does 
            not exist. No extra columns will be added.", prefix = " ", initial = ""))
    
  }
  
  if (links == "Yes") {
    
    definitionsdf_temp <- definitionsdf %>%
      dplyr::select("Term", "Definition", "Link") %>%
      {if (exists("extracols_definitions", envir = .GlobalEnv) & extracols == "Yes") 
        dplyr::bind_cols(., extracols_definitions) else .}
    
    class(definitionsdf_temp$Link) <- "formula"
    
    assign("definitionsdf", definitionsdf_temp, envir = as.environment(acctabs))
    rm(definitionsdf_temp)
    
  } else if (links == "No") {
    
    definitionsdf_temp <- definitionsdf %>%
      dplyr::select("Term", "Definition") %>%
      {if (exists("extracols_definitions", envir = .GlobalEnv) & extracols == "Yes") 
        dplyr::bind_cols(., extracols_definitions) else .}
    
    assign("definitionsdf", definitionsdf_temp, envir = as.environment(acctabs))
    rm(definitionsdf_temp)
    
  }
  
  definitionsdf <- acctabs$definitionsdf
  
  if ("Link" %in% colnames(definitionsdf)) {
    
    definitionsdfcols <- colnames(definitionsdf)
    
    for (i in seq_along(definitionsdfcols)) {
      
      if (definitionsdfcols[i] == "Link") {linkcolpos <- i}
      
    }
    
  }
  
  if (extracols == "Yes" & exists("extracols_definitions", envir = .GlobalEnv)) {
    
    if (any(duplicated(colnames(definitionsdf))) == TRUE) {
      
      warning(strwrap("There is at least one duplicate column name in the definitions table and the 
              extracols_definitions dataframe", prefix = " ", initial = ""))
      
    }
    
  }
  
  # Define some formatting for use later on
  
  normalformat <- openxlsx::createStyle(valign = "top")
  topformat <- openxlsx::createStyle(valign = "bottom")
  linkformat <- openxlsx::createStyle(fontColour = "blue", textDecoration = "underline", 
                                      valign = "top")
  
  openxlsx::addStyle(wb, "Definitions", normalformat, rows = 1:(nrow(definitionsdf) + 4), 
                     cols = 1:ncol(definitionsdf), gridExpand = TRUE)
  
  openxlsx::writeData(wb, "Definitions", "Definitions", startCol = 1, startRow = 1)
  
  titleformat <- openxlsx::createStyle(fontSize = fontszt, textDecoration = "bold")
  
  openxlsx::addStyle(wb, "Definitions", titleformat, rows = 1, cols = 1)
  
  openxlsx::writeData(wb, "Definitions", "This worksheet contains one table.", startCol = 1, 
                      startRow = 2)
  
  openxlsx::addStyle(wb, "Definitions", topformat, rows = 2, cols = 1)
  
  extraformat <- openxlsx::createStyle(valign = "top")
  
  if (contentstab == "Yes") {
    
    openxlsx::addStyle(wb, "Definitions", linkformat, rows = 3, cols = 1)
    
    openxlsx::writeFormula(wb, "Definitions", startRow = 3, 
                           x = openxlsx::makeHyperlinkString("Contents", row = 1, 
                                                             col = 1, text = "Link to contents"))
    
    startingrow <- 4
    
  } else if (contentstab == "No") {
    
    startingrow <- 3
    
    openxlsx::addStyle(wb, "Definitions", extraformat, rows = 2, cols = 1, stack = TRUE)
    
  }
  
  if (links == "Yes" & is.null(linkrange)) {
    
    stop(strwrap("Links are required in the definitions tab but the row numbers where links should 
         be have not been generated (i.e., linkrange not populated)", prefix = " ", initial = ""))
    
  } else if (links == "Yes" & !is.null(linkrange)) {
    
    openxlsx::addStyle(wb, "Definitions", linkformat, rows = linkrange + startingrow, 
                       cols = linkcolpos, gridExpand = TRUE)
    
  }
  
  openxlsx::writeDataTable(wb, "Definitions", definitionsdf, tableName = "definitions", 
                           startRow = startingrow, startCol = 1, withFilter = FALSE, 
                           tableStyle = "none")
  
  if ("Link" %in% colnames(definitionsdf)) {
    
    noaddlinks <- definitionsdf[["Link"]]
    
    for (i in seq_along(noaddlinks)) {
      
      if (noaddlinks[i] == "No additional link") {
        
        openxlsx::writeData(wb, "Definitions", "No additional link", startCol = linkcolpos, 
                            startRow = startingrow + i)
        
      } 
      
    }
    
  }
  
  headingsformat <- openxlsx::createStyle(textDecoration = "bold", wrapText = TRUE, border = NULL, 
                                          valign = "top")
  
  openxlsx::addStyle(wb, "Definitions", headingsformat, rows = startingrow, 
                     cols = 1:ncol(definitionsdf))
  
  extraformat2 <- openxlsx::createStyle(wrapText = TRUE)
  
  openxlsx::addStyle(wb, "Definitions", extraformat2, 
                     rows = (startingrow + 1):(nrow(definitionsdf) + startingrow + 1), 
                     cols = 1:2, stack = TRUE, gridExpand = TRUE)
  
  if ((!is.null(colwid_spec) & !is.numeric(colwid_spec)) | 
      (!is.null(colwid_spec) & length(colwid_spec) != ncol(definitionsdf))) {
    
    warning(strwrap("colwid_spec is either a non-numeric value or a vector not of the same length as 
            the number of columns desired in the definitions tab. The widths will be determined 
            automatically.", prefix = " ", initial = ""))
    colwid_spec <- NULL
    
  }
  
  numchars1 <- max(nchar(definitionsdf$"Term")) + 10
  numchars2 <- max(nchar(definitionsdf$"Definition")) + 10
  
  if (links == "Yes" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Definitions", 
                           cols = c(1,2,3,4:max(ncol(definitionsdf),4)), 
                           widths = c(min(numchars1, 30), min(numchars2, 75), "auto", "auto"))
    
  } else if (links == "No" & is.null(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Definitions", 
                           cols = c(1,2,3:max(ncol(definitionsdf),3)), 
                           widths = c(min(numchars1, 30), min(numchars2, 75), "auto"))
    
  } else if (!is.null(colwid_spec) & is.numeric(colwid_spec) & 
             length(colwid_spec) == ncol(definitionsdf)) {
    
    openxlsx::setColWidths(wb, "Definitions", cols = c(1:ncol(definitionsdf)), widths = colwid_spec)
    
  }
  
  openxlsx::setRowHeights(wb, "Definitions", startingrow - 1, fontsz * (25/12))
  
  rm(definitionsdf, envir = as.environment(acctabs))
  
  # If gridlines are not wanted, then they are removed
  
  if (gridlines == "No") {
    
    openxlsx::showGridLines(wb, "Definitions", showGridLines = FALSE)
    
  }
  
  assign("wb", wb, envir = as.environment(acctabs))
  
}

###################################################################################################