###################################################################################################
# COVER

#' @title accessibletablesR::coverpage
#' 
#' @description Create a cover page for the workbook.
#' 
#' @details 
#' coverpage function will create a cover page for the front of the workbook.
#' If a cover page is not wanted, do not run the coverpage function.
#' The only compulsory parameter is title.
#' All other parameters are optional and preset, only populate if they are wanted.
#' intro: Introductory information / about: About these data / dop: Date of publication.
#' source: Data source(s) used / blank: Information about why some cells are blank, if necessary.
#' relatedlink and relatedtext - any publications associated with the data (relatedlink is the 
#' actual hyperlink, relatedtext is the text you want to appear to the user).
#' names: Contact name / email: Contact email / phone: Contact telephone.
#' reuse: Set to "Yes" if you want the information displayed about the reuse of the data (will 
#' automatically be populated).
#' govdept: Default is "ONS" but if want reuse information without reference to ONS change govdept.
#' extrafields: Any additional fields that the user wants present on the cover page.
#' extrafieldsb: The text to go in any additional fields. Only one row per field. extrafields and 
#' extrafields must be vectors of the same length.
#' additlinks: Any additional hyperlinks the user wants.
#' addittext: The text to appear over any additional hyperlinks. additlinks and addittext must be 
#' vectors of the same length.
#' order: If the user wants the cover page to be ordered in a specific way, list the fields in a 
#' vector with each field name in speech marks.
#' e.g., order = c("intro", "about", relatedlink", "names", "phone", "email", "extrafields").
#' Change gridlines to "No" if gridlines are not wanted.
#' Column width automatically set unless user specifies a value in colwid_spec.
#' intro, about, source, dop, blank, names, phone can be set to hyperlinks - 
#' e.g., source = "[ONS](https://www.ons.gov.uk)".
#' 
#' @param title Title for workbook
#' @param intro Introductory information (optional)
#' @param about About the data (optional)
#' @param source Data source(s) (optional)
#' @param relatedlink Link(s) to related publications (optional)
#' @param relatedtext Text to appear in place of associated link (optional)
#' @param dop Date of publication (optional)
#' @param blank Blank cell information (optional)
#' @param names Contact name (optional)
#' @param email Contact email (optional)
#' @param phone Contact phone number (optional)
#' @param reuse Define whether want to use default text on reuse of publication (optional)
#' @param gridlines Define whether gridlines are present (optional)
#' @param govdept UK Government department name (optional)
#' @param extrafields Additional fields for the cover page (optional)
#' @param extrafieldsb Text to appear in additional fields (optional)
#' @param additlinks Additional links of relevance (optional)
#' @param addittext Text to appear in place of associated additional link (optional)
#' @param colwid_spec Define widths of columns (optional)
#' @param order List of fields in order of appearance wanted on cover page (optional)
#' 
#' @returns A worksheet of the cover page for workbook
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

coverpage <- function(title, intro = NULL, about = NULL, source = NULL, relatedlink = NULL, 
                      relatedtext = NULL, dop = NULL, blank = NULL, names = NULL, email = NULL, 
                      phone = NULL, reuse = NULL, govdept = NULL, gridlines = "Yes",
                      extrafields = NULL, extrafieldsb = NULL, additlinks = NULL, addittext = NULL, 
                      colwid_spec = NULL, order = NULL) {
  
  if (!("openxlsx" %in% utils::installed.packages()) |
      !("conflicted" %in% utils::installed.packages()) |
      !("stringr" %in% utils::installed.packages())) {
    
    stop(base::strwrap("Not all required packages installed. Run the \"workbook\" function first to 
         ensure packages are installed.", prefix = " ", initial = ""))
    
  } else if (utils::packageVersion("openxlsx") < "4.2.5.2" | 
             utils::packageVersion("conflicted") < "1.2.0" | 
             utils::packageVersion("stringr") < "1.5.0") {
    
    stop(base::strwrap("Older versions of packages detected. Run the \"workbook\" function first to 
         ensure up to date packages are installed.", prefix = " ", initial = ""))
    
  }
  
  conflicted::conflict_prefer_all("base", quiet = TRUE)
  
  if (!(exists("wb", envir = as.environment(acctabs)))) {
    
    stop("Run the \"workbook\" function first to ensure that a workbook named wb exists")
    
  }
  
  wb <- acctabs$wb
  fontszst <- acctabs$fontszst
  fontszt <- acctabs$fontszt
  
  # Check to see that a coverpage is wanted, based on whether a worksheet was created in the ...
  # ... initial workbook
  
  if (!("Cover" %in% names(wb))) {
    
    stop("covertab has not been set to \"Yes\" in the workbook function call")
    
  }
  
  # Checking some parameters have been populated properly, the function will error if not
  
  if (is.null(title)) {
    
    stop("No title entered. Must have a title.")
    
  } else if (title == "") {
    
    stop("No title entered. Must have a title.")
    
  }
  
  if (length(title) > 1 | length(intro) > 1 | length(about) > 1 | length(source) > 1 | 
      length(dop) > 1 | length(blank) > 1 | length(names) > 1 | length(email) > 1 | 
      length(phone) > 1 | length(reuse) > 1 | length(govdept) > 1) {
    
    stop(strwrap("One of title, intro, about, source, dop, blank, names, email, phone, reuse and 
         govdept is more than a single entity", prefix = " ", initial = ""))
    
  }
  
  if (!is.null(relatedlink) & is.null(relatedtext)) {
    
    stop("relatedlink and relatedtext either have to be both set to NULL or both set to something")
    
  } else if (is.null(relatedlink) & !is.null(relatedtext)) {
    
    stop("relatedlink and relatedtext either have to be both set to NULL or both set to something")
    
  } else if (length(relatedlink) != length(relatedtext)) {
    
    stop(strwrap("relatedlink and relatedtext must be of the same length and contain the same number 
         of elements", prefix = " ", initial = ""))
    
  }
  
  if (is.null(extrafields) & !is.null(extrafieldsb)) {
    
    stop("extrafields and extrafieldsb either have to be both set to NULL or both set to something")
    
  } else if (!is.null(extrafields) & is.null(extrafieldsb)) {
    
    stop("extrafields and extrafieldsb either have to be both set to NULL or both set to something")
    
  } else if (length(extrafields) != length(extrafieldsb)) {
    
    stop(strwrap("extrafields and extrafieldsb must be of the same length and contain the same 
         number of elements", prefix = " ", initial = ""))
    
  }
  
  if (is.null(additlinks) & !is.null(addittext)) {
    
    stop("additlinks and addittext either have to be both set to NULL or both set to something")
    
  } else if (!is.null(additlinks) & is.null(addittext)) {
    
    stop("additlinks and addittext either have to be both set to NULL or both set to something")
    
  } else if (length(additlinks) != length(addittext)) {
    
    stop(strwrap("additlinks and addittext must be of the same length and contain the same number of 
         elements", prefix = " ", initial = ""))
    
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
    
    stop(strwrap("gridlines has not been populated properly. It must be a single word, either 
         \"Yes\" or \"No\".", prefix = " ", initial = ""))
    
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
  
  if (!is.null(phone)) {
    
    if (stringr::str_remove_all(phone, "[\" \"\\[\\]\\(\\)+[:digit:]]") != "") {
      
      warning(strwrap("The phone number provided appears to contain characters which are unusual 
              for a phone number. Check if there are any errors.", prefix = " ", initial = ""))
      
    }
    
  }
  
  if (grepl("\\.", email) == FALSE | grepl("@", email) == FALSE) {
    
    warning(strwrap("The email address provided does not appear to contain @ and/or a dot (.). Check 
            if there are any errors.", prefix = " ", initial = ""))
    
  }
  
  # In case the function is run multiple times, removing previous row heights to ensure there ...
  # ... will be no strange looking rows
  
  if (exists("covernumrow", envir = as.environment(acctabs))) {
    
    covernumrow <- acctabs$covernumrow
    
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
  
  covernumrow <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + 
    source2 + length(relatedlink) + related2 + length(dop) + dop2 + length(blank) + blank2 + 
    length(extrafields) + length(extrafieldsb) + additlinks2 + length(additlinks) + length(names) + 
    names2 + length(email) + length(phone) + length(reuse) + 4
  
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
    
    fields <- c(title, intro, about, source, relatedlink, dop, blank, names, email, phone, reuse, 
                extrafields, additlinks)
    
    for (i in seq_along(order)) {
      
      if (tolower(order[i]) %in% c("intro", "introduction", "introductory information")) {
        
        order[i] <- intro
        
      } else if (tolower(order[i]) %in% c("about", "about these data")) {
        
        order[i] <- about
        
      } else if (tolower(order[i]) %in% c("source", "source of data", "data source", "sources", 
                                          "sources of data", "data sources")) {
        
        order[i] <- source
        
      } else if (tolower(order[i]) %in% c("related publications", "related publication", "related", 
                                          "relatedlink", "relatedlinks", "relatedtext")) {
        
        order[i] <- "relatedlink"
        
      } else if (tolower(order[i]) %in% c("dop", "date of publication", "publication date")) {
        
        order[i] <- dop
        
      } else if (tolower(order[i]) %in% c("blank", "blank cells")) {
        
        order[i] <- blank
        
      } else if (tolower(order[i]) %in% c("names", "name", "contact", "contact details")) {
        
        order[i] <- names
        
      } else if (tolower(order[i]) %in% c("email", "email address", "e-mail", "e-mail address")) {
        
        order[i] <- email
        
      } else if (tolower(order[i]) %in% c("phone", "telephone", "phone number", "telephone number", 
                                          "tel", "tel:")) {
        
        order[i] <- phone
        
      } else if (tolower(order[i]) %in% c("reuse", "reusing this publication", 
                                          "reuse this publication")) {
        
        order[i] <- reuse
        
      } else if (tolower(order[i]) %in% c("extrafields", "extrafield", "extrafieldsb", 
                                          "extrafieldb")) {
        
        order[i] <- "extrafields"
        
      } else if (tolower(order[i]) %in% c("additlinks", "additlink", "addittext", 
                                          "additional links", "additional link")) 
      {order[i] <- "additlinks"}
      
    }
    
    phone2 <- which(order == phone)
    email2 <- which(order == email)
    names2 <- which(order == names)
    
    if (names %in% order & phone %in% order & email %in% order) {
      
      if ((phone2 < names2) | (email2 < names2) | (phone2 > (names2 + 2)) | 
          (email2 > (names2 + 2))) {
        
        stop(strwrap("The relative positions of names, phone and email are not consistent with the 
             expected stucture (i.e., names, phone or email, email or phone)", prefix = " ",
                     initial = ""))
        
      }
      
    } else if (names %in% order & phone %in% order) {
      
      if ((phone2 < names2) | (phone2 > (names2 + 1))) {
        
        stop(strwrap("The relative positions of names and phone are not consistent with the expected 
             structure (i.e., names, phone)", prefix = " ", initial = ""))
        
      }
      
    } else if (names %in% order & email %in% order) {
      
      if ((email2 < names2) | (email2 > (names2 + 1))) {
        
        stop(strwrap("The relative positions of names and email are not consistent with the expected 
             structure (i.e., names, email)", prefix = " ", initial = ""))
        
      }
      
    } else if (((email %in% order) | (phone %in% order)) & !(names %in% order)) {
      
      warning(strwrap("email and/or phone have been populated but a contact name has not been 
              provided. Check that this is intentional.", prefix = " ", initial = ""))
      
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
      else if (order[i] %in% extrafields) 
      {extrastartpos <- append(extrastartpos, orderl[i] + length(title))}
      else if (order[i] == additlinks[1]) {additstartpos <- orderl[i] + length(title)}
      
    }
    
    if (is.null(introstartpos) & is.null(aboutstartpos) & is.null(sourcestartpos) & 
        is.null(relatedstartpos) & is.null(dopstartpos) & is.null(blankstartpos) & 
        is.null(namesstartpos) & is.null(emailstartpos) & is.null(phonestartpos) & 
        is.null(reusestartpos) & is.null(extrastartpos) & is.null(additstartpos)) {
      
      stop("No starting positions have been generated")
      
    }
    
    if (length(extrastartpos) != length(extrafields)) {
      
      stop(strwrap("The lengths of the vectors for extrafields and their row starting positions are 
           not equal. Investigate why.", prefix = " ", initial = ""))
      
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
      
      if (substr(fields2[i], 1, 1) != "[" | 
          substr(fields2[i], nchar(fields2[i]), nchar(fields2[i])) != ")") {
        
        warning(strwrap(paste0(fields2[i], " - if this is meant to be a hyperlink, it needs to be in 
                the format \"[xxx](xxxxxx)\""), prefix = " ", initial = ""))
        
      }
      
      if ("phone" %in% fields5[i]) {
        
        phone <- paste0("[Telephone: ", substr(phone, 2, nchar(phone)))
        
      }
      
      # Hyperlink code taken from Matt Dray's a11ytables
      
      md_rx <- "\\[(([[:graph:]]|[[:space:]])+?)\\]\\([[:graph:]]+?\\)"
      md_match <- regexpr(md_rx, fields2[i], perl = TRUE)
      md_extract <- regmatches(fields2[i], md_match)[[1]]
      
      url_rx <- "(?<=\\]\\()([[:graph:]])+(?=\\))"
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
      
      rm(x, md_rx, md_match, md_extract, url_rx, url_match, url_extract, string_rx, string_match, 
         string_extract)
      
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
    
    openxlsx::writeData(wb, "Cover", "Introductory information", startCol = 1, 
                        startRow = introstart)
    openxlsx::writeData(wb, "Cover", intro, startCol = 1, startRow = introstart + 1)
    
  }
  
  if (!is.null(about)) {
    
    if (is.null(aboutstartpos)) {
      
      aboutstart <- length(title) + length(intro) + intro2 + 1
      
    } else if (!is.null(aboutstartpos)) {
      
      aboutstart <- aboutstartpos
      
    }
    
    openxlsx::writeData(wb, "Cover", "About these data", startCol = 1, 
                        startRow = aboutstart)
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
      
      relatedstart <- length(title) + length(intro) + intro2 + length(about) + about2 + 
        length(source) + source2 + 1
      
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
      
      dopstart <- length(title) + length(intro) + intro2 + length(about) + about2 + length(source) + 
        source2 + length(relatedlink) + related2 + 1
      
    } else if (!is.null(dopstartpos)) {
      
      dopstart <- dopstartpos
      
    }  
    
    openxlsx::writeData(wb, "Cover", "Date of publication", startCol = 1, startRow = dopstart)
    openxlsx::writeData(wb, "Cover", dop, startCol = 1, startRow = dopstart + 1)
    
    
  }
  
  if (!is.null(blank)) {
    
    if (is.null(blankstartpos)) {
      
      blankstart <- length(title) + length(intro) + intro2 + length(about) + about2 + 
        length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + 1
      
    } else if (!is.null(blankstartpos)) {
      
      blankstart <- blankstartpos
      
    }  
    
    openxlsx::writeData(wb, "Cover", "Blank cells", startCol = 1, startRow = blankstart)
    openxlsx::writeData(wb, "Cover", blank, startCol = 1, startRow = blankstart + 1)
    
    
  }
  
  if (!is.null(extrafields)) {
    
    for (i in seq_along(extrafields)) {
      
      if (is.null(extrastartpos)) {
        
        extrastart <- length(title) + length(intro) + intro2 + length(about) + about2 + 
          length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + 
          length(blank) + blank2 + 1 + (2 * i) - 2
        
      } else if (!is.null(extrastartpos)) {
        
        extrastart <- extrastartpos[i]
        
      }
      
      if (grepl(hyper_rx, extrafieldsb[i]) == TRUE) {
        
        if (substr(extrafieldsb[i], 1, 1) != "[" | 
            substr(extrafieldsb[i], nchar(extrafieldsb[i]), nchar(extrafieldsb[i])) != ")") {
          
          warning(strwrap(paste0(extrafieldsb[i], " - if this is meant to be a hyperlink, it needs 
                  to be in the format \"[xxx](xxxxxx)\""), prefix = " ", initial = ""))
          
        }
        
        x <- extrafieldsb[i]
        
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
        
        y <- extrafieldsb[i]
        
      }
      
      openxlsx::writeData(wb, "Cover", extrafields[i], startCol = 1, startRow = extrastart)
      openxlsx::writeData(wb, "Cover", y, startCol = 1, startRow = extrastart + 1)
      
      rm(y)
      
    }
    
  }
  
  if (!is.null(additlinks)) {
    
    if (is.null(additstartpos)) {
      
      additlinkstart <- length(title) + length(intro) + intro2 + length(about) + about2 + 
        length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + 
        length(blank) + blank2 + length(extrafields) + length(extrafieldsb) + 1
      
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
      
      namesstart <- length(title) + length(intro) + intro2 + length(about) + about2 + 
        length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + 
        length(blank) + blank2 + length(extrafields) + length(extrafieldsb) + additlinks2 + 
        length(additlinks) + 1
      
    } else if (!is.null(namesstartpos)) {
      
      namesstart <- namesstartpos
      
    }
    
    openxlsx::writeData(wb, "Cover", "Contact", startCol = 1, startRow = namesstart)
    openxlsx::writeData(wb, "Cover", names, startCol = 1, startRow = namesstart + 1)
    
  }
  
  normalformat <- openxlsx::createStyle(valign = "top", wrapText = TRUE)
  subtitleformat <- openxlsx::createStyle(fontSize = fontszst, valign = "bottom", wrapText = TRUE, 
                                          textDecoration = "bold")
  titleformat <- openxlsx::createStyle(fontSize = fontszt, valign = "bottom", wrapText = TRUE, 
                                       textDecoration = "bold")
  linkformat <- openxlsx::createStyle(fontColour = "blue", valign = "top", wrapText = TRUE, 
                                      textDecoration = "underline")
  
  if (is.null(colwid_spec) | !is.numeric(colwid_spec) | length(colwid_spec) > 1) {
    
    openxlsx::setColWidths(wb, "Cover", cols = 1, widths = 100)
    
    if (!is.null(colwid_spec) & !is.numeric(colwid_spec)) {
      
      warning(strwrap("colwid_spec has not been provided as a numeric value and so the default width 
              of 100 has been used", prefix = " ", initial = ""))
      
    } else if (!is.null(colwid_spec) & length(colwid_spec) > 1) {
      
      warning(strwrap("colwid_spec has been provided as a vector with more than one element and so 
              the default width of 100 has been used", prefix = " ", initial = ""))
      
    }
    
  } else if (is.numeric(colwid_spec)) {
    
    openxlsx::setColWidths(wb, "Cover", cols = 1, widths = colwid_spec)
    
  }
  
  openxlsx::addStyle(wb, "Cover", normalformat, rows = c(1:covernumrow), cols = 1)
  
  if (!is.null(intro)) {
    
    openxlsx::setRowHeights(wb, "Cover", introstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = introstart, cols = 1)
    
    if (intro_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = introstart + 1, 
                                              cols = 1)}
    
  }
  
  if (!is.null(about)) {
    
    openxlsx::setRowHeights(wb, "Cover", aboutstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = aboutstart, cols = 1)
    
    if (about_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = aboutstart + 1, 
                                              cols = 1)}
    
  }
  
  if (!is.null(source)) {
    
    openxlsx::setRowHeights(wb, "Cover", sourcestart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = sourcestart, cols = 1)
    
    if (source_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = sourcestart + 1, 
                                               cols = 1)}
    
  }
  
  if (!is.null(relatedlink)) {
    
    openxlsx::setRowHeights(wb, "Cover", relatedstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = relatedstart, cols = 1)
    openxlsx::addStyle(wb, "Cover", linkformat, 
                       rows = (relatedstart + 1):(relatedstart + length(relatedlink)), cols = 1)
    
  }
  
  if (!is.null(dop)) {
    
    openxlsx::setRowHeights(wb, "Cover", dopstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = dopstart, cols = 1)
    
    if (dop_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = dopstart + 1, cols = 1)}
    
  }
  
  if (!is.null(blank)) {
    
    openxlsx::setRowHeights(wb, "Cover", blankstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = blankstart, cols = 1)
    
    if (blank_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = blankstart + 1, 
                                              cols = 1)}
    
  }
  
  if (!is.null(extrafields)) {
    
    for (i in seq_along(extrafields)) {
      
      if (is.null(extrastartpos)) {
        
        extrastart <- length(title) + length(intro) + intro2 + length(about) + about2 + 
          length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + 
          length(blank) + blank2 + 1 + (2 * i) - 2
        
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
    openxlsx::addStyle(wb, "Cover", linkformat, 
                       rows = (additlinkstart + 1):(additlinkstart + length(additlinks)), cols = 1)
    
  }
  
  if (!is.null(names)) {
    
    openxlsx::setRowHeights(wb, "Cover", namesstart, fontszst * (25/14))
    openxlsx::addStyle(wb, "Cover", subtitleformat, rows = namesstart, cols = 1)
    
    if (names_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = namesstart + 1, 
                                              cols = 1)}
    
  }
  
  openxlsx::setRowHeights(wb, "Cover", 2, fontszst * (34/14))
  
  openxlsx::addStyle(wb, "Cover", titleformat, rows = 1, cols = 1)
  
  # Create a hyperlink for any given email address
  
  if (!is.null(email)) {
    
    x <- paste0("mailto:", email)
    names(x) <- email
    class(x) <- "hyperlink"
    
    if (is.null(emailstartpos)) {
      
      emailstart <- length(title) + length(intro) + intro2 + length(about) + about2 + 
        length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + 
        length(blank) + blank2 + length(extrafields) + length(extrafieldsb) + additlinks2 + 
        length(additlinks) + length(names) + names2 + 1
      
    } else if (!is.null(emailstartpos)) {
      
      emailstart <- emailstartpos
      
    }
    
    openxlsx::writeData(wb, "Cover", x, startCol = 1, startRow = emailstart)
    
    openxlsx::addStyle(wb, "Cover", linkformat, rows = emailstart, cols = 1)
    
    emailformat <- openxlsx::createStyle(fontColour = "blue", valign = "bottom", 
                                         textDecoration = "underline", wrapText = TRUE)
    
    if (emailstart == 2) {openxlsx::addStyle(wb, "Cover", emailformat, rows = 2, cols = 1)}
    
  }
  
  if (!is.null(phone)) {
    
    if (is.null(phonestartpos)) {
      
      phonestart <- length(title) + length(intro) + intro2 + length(about) + about2 + 
        length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + 
        length(blank) + blank2 + length(extrafields) + length(extrafieldsb) + additlinks2 + 
        length(additlinks) + length(names) + names2 + length(email) + 1
      
    } else if (!is.null(phonestartpos)) {
      
      phonestart <- phonestartpos
      
    }
    
    openxlsx::writeData(wb, "Cover", phone, startCol = 1, startRow = phonestart)
    
    phoneformat <- openxlsx::createStyle(valign = "bottom")
    
    if (phone_hyper == 1) {openxlsx::addStyle(wb, "Cover", linkformat, rows = phonestart, 
                                              cols = 1, stack = TRUE)}
    
    if (phonestart == 2) {openxlsx::addStyle(wb, "Cover", phoneformat, rows = 2, cols = 1, 
                                             stack = TRUE)}
    
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
      
    } else if (tolower(govdept) == "ons" | tolower(govdept) == "office for national statistics") {
      
      orgwording <- "the Office for National Statistics - Source: Office for National Statistics"
      
    }
    
    reuse1 <- paste0("You may re-use this publication (not including logos) free of charge in any ",
                     "format or medium, under the terms of the Open Government Licence. Users ",
                     "should include a source accreditation to ", orgwording, " licensed under ",
                     "the Open Government Licence.")
    reuse2 <- paste0("Alternatively you can write to: Information Policy Team, The National ",
                     "Archives, Kew, Richmond, Surrey, TW9 4DU; or ",
                     "email: psi@nationalarchives.gov.uk")
    reuse3 <- paste0("Where we have identified any third party copyright information you will ",
                     "need to obtain permission from the copyright holders concerned.")
    licencelink <- "https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/"
    licencetext <- "View the Open Government Licence"
    
    if (is.null(reusestartpos)) {
      
      reusestart <- length(title) + length(intro) + intro2 + length(about) + about2 + 
        length(source) + source2 + length(relatedlink) + related2 + length(dop) + dop2 + 
        length(blank) + blank2 + length(extrafields) + length(extrafieldsb) + additlinks2 + 
        length(additlinks) + length(names) + names2 + length(email) + length(phone) + 1
      
    } else if (!is.null(reusestartpos)) {
      
      reusestart <- reusestartpos
      
    }  
    
    openxlsx::writeData(wb, "Cover", "Reusing this publication", startRow = reusestart, 
                        startCol = 1)
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
  
  assign("covernumrow", covernumrow, envir = as.environment(acctabs))
  assign("wb", wb, envir = as.environment(acctabs))
  
}

###################################################################################################