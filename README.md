# accessibletablesR

accessibletablesR is designed to produce Excel workbooks that align as closely as possible to the UK Government Analysis Function recommendations for publishing statistics in spreadsheets. 
Releasing statistics in spreadsheets - https://analysisfunction.civilservice.gov.uk/policy-store/releasing-statistics-in-spreadsheets/

accessibletablesR was developed using R 4.1.3, tidyverse version 2.0.0, openxlsx version 4.2.5.2 and conflicted version 1.2.0. It is unknown if the package will work earlier versions of R,
tidyverse, openxlsx and conflicted. accessibletablesR will install the latest versions of tidyverse, openxlsx and conflicted if these packages are not currently installed or if earlier 
versions of tidyverse (<2.0.0), openxlsx (<4.2.5.2) and conflicted (1.2.0) are currently installed.

To install accessibletablesR:

        if (!("devtools" %in% utils::installed.packages())) {utils::install.packages("devtools", dependencies = TRUE, type = "binary")}
        library("devtools")
        devtools::install_github("minifiemj/accessibletablesR")

If a firewall prevents install_github from working (a time out message may appear) then install the package manually. 
On the GitHub repo, go to the green Code icon and choose "Download ZIP". Copy the ZIP folder to a network drive.
Use devtools::install_local(link to the zipped folder) to install the package.

accessibletablesR allows for a workbook to have a cover page, a table of contents, a notes page, a definitions page and as many other tabs that the user requires (subject to the maximum number
allowed by Excel). Only one table of data can be present on each tab. accessibletablesR cannot work with multiple tables on a tab.

The final output is a xlsx file. It is not currently known how to produce an ODS workbook. For accessibility reasons, please consider manually saving an ODS workbook.

accessibletablesR contains nine main functions - workbook, creatingtables, contentstable, coverpage, addnote, notestab, adddefinition, definitionstab and savingtables.

workbook:

workbook <- function(covertab = NULL, contentstab = NULL, notestab = NULL, autonotes = NULL,
                     definitionstab = NULL, fontnm = "Arial", fontcol = "black",
                     fontsz = 12, fontszst = 14, fontszt = 16, title = NULL, creator = NULL,
                     subject = NULL, category = NULL)

This function needs to be run first. It creates the workbook within the R environment. All arguments are optional. The user can specify whether they want a cover page, table of contents, 
a notes page and a definitions page. The font name, colour and sizes can be modified if desired. Information regarding the final spreadsheet (File > Info > Properies: title, creator, subject and
category) can also be specified. There is an option for the automatic display of notes on the applicable tabs.

The default font (fontnm) is Arial, the default colour (fontcol) is black, the default normal size (fontsz) is 12, the default subtitle size (fontszst) is 14 and the default title size (fontszt)
is 16.

If a coverpage is wanted set covertab = "Yes". If a table of contents is wanted set contentstab = "Yes". If a notes page is wanted set notestab = "Yes". If a definitions page is wanted set 
definitionstab = "Yes". If the automatic display of notes on the applicable tabs is wanted set autonotes = "Yes".

To set some of the spreadsheet information, amend title, creator, subject or category.

creatingtables: 

creatingtables <- function(title, subtitle = NULL, extraline1 = NULL, extraline2 = NULL, extraline3 = NULL,
                           extraline4 = NULL, extraline5 = NULL, extraline6 = NULL, sheetname, table_data, 
                           headrowsize = NULL, numdatacols = NULL, numdatacolsdp = NULL, othdatacols = NULL, 
                           tablename = NULL, gridlines = "Yes", columnwidths = "R_auto", width_adj = NULL,
                           colwid_spec = NULL)

This function takes the raw data and transfers them into an accessible data table for the final Excel workbook. The raw data must be in a dataframe within the global environment. The
raw data need to be ordered as desired and contain the columns desired in the right position and named accordingly.

The creatingtables function needs to be run after the workbook function and run as many times as there are tables which are to be put in tabs in the final workbook.

Three of the arguments (title, sheetname, table_data) are compulsory, the others are optional. title is the title of the table that will be displayed in the Excel workbook tab.
sheetname is the name of the tab wanted for the Excel workbook. table_data is the dataframe in the R global environment to be outputted.

As well as a title, it is possible to include a subtitle and six additional lines above the table in the final Excel workbook tab. Populate the arguments subtitle, extraline1, extraline2,
extraline3, extraline4, extraline5 and extraline6 if a subtitle and/or extra lines are wanted. If a link to the contents page or notes page or definitions page is desired, then set one
of the extraline arguments to "Link to contents" or "Link to notes" or "Link to definitions". The extraline arguments can be supplied as vectors and so there is no maximum limit to the
number of rows that can come before the main data other than the limit of rows in an Excel spreadsheet.

headrowsize adjusts the height of the row containing the table column names.

If the automatic formatting of columns containing data in the form of numbers is required then the user should populate numdatacols and numdatacolsdp. numdatacols should contain the
numerical positions of the columns in the dataframe. numdatacolsdp is the required number of decimal places for each column containing data in the form of numbers. Populating
numdatacols and numdatacolsdp ensures that decimal places and thousand commas will be sorted for the final Excel workbook regardless of if the numbers are numerical or stored as text.

If there are columns containing data not in the form of numbers (e.g., text) then the user can populate othdatcols with the appropriate numerical positions of columns to ensure the
automatic formatting of the columns in the final workbook. At present it is not possible to fully format dates.

If the user wants to give a name to a table in the final Excel workbook which is different to the tab name (sheetname) then populate tablename.

If the gridlines are not desired in the final workbook, set gridlines = "No".

Automatic column widths can be a bit hit and miss. The default position of the creatingtables function is to allow openxlsx to automatically determine the column widths. The user can
instead use the maximum number of characters in a column by setting columnwidths = "characters". width_adj can be adjusted as an extra bit of width to add on to the width determined
by the number of characters. If the user knows the desired widths of all columns then they should set columnwidths = "specified" and populate colwid_spec with the width of each column 
in a numerical vector. If default column widths are wanted then set columnwidths = NULL.

contentstable:

contentstable <- function(gridlines = "Yes", colwid_spec = NULL, extracols = NULL)

If a contents page is wanted then run contentstable(). Run the function after all the data tables have been processed using the creatingtables function.

The arguments are optional. If no gridlines are wanted in the contents page in the final workbook set gridlines = "No". Column widths are determined automatically but the user can
specify the widths by populating colwid_spec. Extra columns can be provided. To do so, set extracols = "Yes" and create a dataframe called extracols_contents in the global 
environment before running the contentstable function. The extracols_contents dataframe must contain the desired extra columns and have the same number of rows as the contents table.

coverpage:

coverpage <- function(title, intro = NULL, about = NULL, source = NULL, relatedlink = NULL, relatedtext = NULL, 
                      dop = NULL, blank = NULL, names = NULL, email = NULL, phone = NULL, reuse = NULL, 
                      gridlines = "Yes", govdept = "ONS", extrafields = NULL, extrafieldsb = NULL,
                      additlinks = NULL, addittext = NULL, colwid_spec = NULL, order = NULL)

Run the function after all the data tables have been processed using the creatingtables function.

title is the only compulsory argument. It is the title to be displayed on the cover sheet. Other optional sections of a cover page that can be populated are "Introductory information"
(info), "About these data" (about), "Source" (source), "Related publications" (relatedlink, relatedtext), "Date of publication" (dop), "Blank cells" (blank), "Contact" (names, email,
phone), "Additional links" (additlinks) and "Reusing this publication" (reuse, govdept). Extra fields can be added using extrafields. One row is allowed for each extra field. The text
to populate the extra fields can be provided in extrafieldsb. If no gridlines are wanted on the cover page in the final workbook set gridlines = "No". The column width is automatically
set but can be altered by using colwid_spec.

The ordering of the fields can be amended by populating order. order can be set to a vector where the field names are provided in speech marks.

The "Reusing this publication" section has been designed for UK government departments and will not apply for other organisations. If a user is from the Office for National Statistics
(ONS) and wants a "Reusing this publication" section then set reuse = "Yes". If a user is from a UK government department but not the Office for National Statistics (ONS) set reuse = "Yes" 
and govdept = "<name of organisation>".

addnote:

addnote <- function(notenumber, notetext, applictabtext = NULL, linktext1 = NULL, linktext2 = NULL)

Run this function for as many notes as are needed.

notenumber is the number of the note and should be written as "note" followed by a number (e.g., note1).

notetext is the description associated with the note.

notenumber and notetext are the only compulsory arguments. If an additional column is wanted to specify which table (sheet name) a note applies to then populate applictabtext
(e.g., applictabtext = c("Table_1", "Table_2")

An optional column can be included that provides a link to a piece of information. To do so, populate linktext1 and linktext2. For example, set linktext1 = "General health information" and
linktext2 = "https://www.ons.gov.uk/census/census2021dictionary/variablesbytopic/healthdisabilityandunpaidcarevariablescensus2021/generalhealth". 

notestab:

notestab <- function(contentslink = NULL, gridlines = "Yes", colwid_spec = NULL, extracols = NULL)

Run this function if a notes page is wanted.

Run the function after all the data tables have been processed using the creatingtables function and after all the notes have been added using the addnote function.

The arguments are optional. If a link to the contents page is not wanted on the notes page set contentslink = "No". If no gridlines are wanted on the notes page then set gridlines = "No".
Column widths are determined automatically but can be altered to specific widths by the user in colwid_spec. Extra columns can be provided. To do so, set extracols = "Yes" and create a 
dataframe called extracols_notes in the global environment before running the notestab function. The extracols_notes dataframe must contain the desired extra columns and have the same 
number of rows as the notes table.

adddefinition:

adddefinition <- function(term, definition)

Run this function for as many definitions as are needed.

term is the item that needs defining and definition is the definition of the item.

definitionstab:

definitionstab <- function(contentslink = NULL, gridlines = "Yes", colwid_spec = NULL, extracols = NULL)

Run this function if a definitions page is wanted.

Run the function after all the definitions have been added using the adddefinition function.

The arguments are optional. If a link to the contents page is not wanted on the definitions page set contentslink = "No". If no gridlines are wanted on the definitions page then set 
gridlines = "No". Column widths are set automatically but can be altered using colwid_spec. Extra columns can be provided. To do so, set extracols = "Yes" and create a 
dataframe called extracols_definitions in the global environment before running the definitionstab function. The extracols_definitions dataframe must contain the desired extra columns 
and have the same number of rows as the definitions table.

savingtables:

savingtables <- function(filename)

This function should be run last and will output the final xlsx workbook.

filename is the location and name of the final workbook.
