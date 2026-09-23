###################################################################################################
# READING IN THE RAW DATA TABLES FROM EXCEL

#' @title accessibletablesR::readin_excel
#' 
#' @description Read in raw xlsx, xls or csv data tables
#' 
#' @details 
#' readin_excel can be used to read in raw xlsx or xls or csv files that are to be converted into
#' accessible data tables. The user can specify the desired names of the columns in the final table.
#' This can be achieved by populating the parameter colnames as a vector of the desired names in the
#' appropriate order. The raw files should have the raw column names in the first row and all other
#' rows should contain data.
#' 
#' @param filepath File path and file name of the input data table, including file type
#' @param sheetname The name of the sheet in the input data table (optional)
#' @param colnames Define the names of the columns you want in the final output (optional)
#' 
#' @returns A dataframe which can be used in the creatingtables function
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
#' @export

readin_excel <- function(filepath, sheetname = NULL, colnames = NULL) {
  
  if (!("readxl" %in% utils::installed.packages()) |
      !("dplyr" %in% utils::installed.packages()) |
      !("conflicted" %in% utils::installed.packages())) {
    
    stop(base::strwrap("Not all required packages installed. Run the \"workbook\" function first to 
         ensure packages are installed.", prefix = " ", initial = ""))
    
  } else if (utils::packageVersion("readxl") < "1.5.0" |
             utils::packageVersion("dplyr") < "1.1.2" |
             utils::packageVersion("conflicted") < "1.2.0") {
    
    stop(base::strwrap("Older versions of packages detected. Run the \"workbook\" function first to 
         ensure up to date packages are installed.", prefix = " ", initial = ""))
    
  }
  
  conflicted::conflict_prefer_all("base", quiet = TRUE)
  `%>%` <- dplyr::`%>%`
  
  # Read in xlsx or xls or csv files and name columns accordingly
  
  if (substr(filepath, nchar(filepath) - 4, nchar(filepath)) == ".xlsx") {
    
    df <- readxl::read_xlsx(path = filepath, sheet = sheetname) 
    
  } else if (substr(filepath, nchar(filepath) - 3, nchar(filepath)) == ".xls") {
    
    df <- readxl::read_xls(path = filepath, sheet = sheetname)
  
  } else if (substr(filepath, nchar(filepath) - 3, nchar(filepath)) == ".csv") {
    
    df <- utils::read.csv(file = filepath)
  
  }
  
  if (!is.null(colnames)) {
    
    df_cols <- ncol(df)
    
    if (length(colnames) != df_cols) {
      
      stop("There is a different number of columns in the dataframe than listed in colnames.")
      
    }
    
    df <- df %>%
      stats::setNames(colnames)
    
  }
  
  return(df)

}

###################################################################################################