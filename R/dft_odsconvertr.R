###################################################################################################
# COPIED FROM DEPARTMENT FOR TRANSPORT ODSCONVERTR - SAVING TO ODS

# @title Save a copy of an xlsx file as an ods file
# 
# @description Save from XLSX to ODS, from Department for Transport 
# 
# @details Uses VBA code to save a copy of an xlsx file to the accessible ODS format, 
# retaining all sheets and formatting. Files converted can be via a relative (from working 
# directory) or absolute (full) file path. In either case, the output ODS file will be returned 
# in the same folder as the XLSX file. Copied from Department for Transport in case the GitHub
# repo disappears (department-for-transport/odsconvertr)
#
# @param path path to xlsx file; can be either a relative or absolute file path

convert_to_ods <- function(path) {
  
  # Stop if file is not found
  
  if (file.exists(path) == FALSE) {
    
    stop("File not found")
    
  }
  
  # Stop if file is not an xlsx
  
  if (grepl(".xlsx", path, fixed = TRUE) == FALSE) {
    
    stop("File is not an xlsx file")
    
  }
  
  # Convert path to absolute one
  
  xlsx_all <- paste0('"',normalizePath(path), '"')
  
  ods_all <- gsub(".xlsx", ".ods", xlsx_all, fixed = TRUE)
  
  # Get path of VBS script inside package
  
  vbs_loc <- vbs_file_path('save.vbs')
  
  # Run VBS script passing it the file paths
  
  vbs_execute("save.vbs", xlsx_all, ods_all)
  
}

# @title Save a copy of an ods file as an xlsx file
# 
# @description Save from ODS to XLSX, from Department for Transport 
# 
# @details Uses VBA code to save a copy of an ods file to the easy to use xlsx format, 
# retaining all sheets and formatting. Files converted can be via a relative (from working 
# directory) or absolute (full) file path. In either case, the output ODS file will be returned 
# in the same folder as the XLSX file. Copied from Department for Transport in case the GitHub
# repo disappears (department-for-transport/odsconvertr)
#
# @param path path to ods file; can be either a relative or absolute file path

convert_to_xlsx <- function(path) {
  
  #Stop if file is not found
  
  if (file.exists(path) == FALSE) {
    
    stop("File not found")
    
  }
  
  # Stop if file is not an xlsx
  
  if (grepl(".ods", path, fixed = TRUE) == FALSE) {
    
    stop("File is not an ods file")
    
  }
  
  # Convert path to absolute one
  
  ods_all <- paste0('"',normalizePath(path), '"')
  
  xlsx_all <- gsub(".ods", ".xlsx", ods_all, fixed = TRUE)
  
  # Get path of VBS script inside package
  
  vbs_loc <- vbs_file_path('save.vbs')
  
  # Run VBS script passing it the file paths
  
  vbs_execute("save_xlsx.vbs", ods_all, xlsx_all)
  
}

# @title Formats the name of a VBS script file into the filepath within the package
# 
# @description Formats the name of a VBS script file into the filepath within the package,
#              from the Department for Transport
# 
# @details Create a filepath for a referenced VBS file. Formats the name of a VBS file to be used.
# Copied from Department for Transport in case the GitHub repo disappears 
# (department-for-transport/odsconvertr)
# 
# @param vbs_file Name of a vbs script file saved as part of the package

vbs_file_path <- function(vbs_file) {
  
  # Get path of VBS script inside package
  
  paste0('"', system.file("vbs", package = "odsconvertr"), '/', vbs_file,'"')
  
}

# @title Execute a VBS script including arguments
# 
# @description Execute a VBS script including arguments, from the Department for Transport
# 
# @details Create a filepath for a referenced VBS file. Formats the name of a VBS file to be used.
# Copied from Department for Transport in case the GitHub repo disappears 
# (department-for-transport/odsconvertr)
# 
# @param vbs_file Name of a vbs script file saved as part of the package
# @param ... Arguments to be passed to the specified VBS script

vbs_execute <- function(vbs_file, ...) {
  
  #Convert arguments into a single string
  
  arguments <- paste(c(...), collapse = " ")
  
  # Create system command for specified vbs file
  
  system_command <- paste("WScript",
                          vbs_file_path(vbs_file),
                          arguments,
                          sep = " ")
  
  # Run specified command
  
  system(command = system_command)
  
}

# @title Evaluate all formulae in an existing Excel file
# 
# @description Evaluate all formulae in an existing Excel file, from the Department for Transport
# 
# @details Uses VBA code to open an Excel file and to recalculate any formula in it. This ensures 
# that all Excel formulae have been executed in a file which may have been updated without being 
# opened (e.g. using code). Copied from Department for Transport in case the GitHub repo disappears 
# (department-for-transport/odsconvertr)
#
# @param path path to Excel file; can be either a relative or absolute file path

evaluate_formula <- function(path) {
  
  # Normalise path to an absolute one with backslashes and surrounding quotation marks
  
  path <- paste0('"', normalizePath(path), '"')
  
  # Execute specified VBS script to save xlsx files
  
  vbs_execute("recalculate.vbs", path)
  
}

###################################################################################################
