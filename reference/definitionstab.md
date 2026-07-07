# accessibletablesR::definitionstab

Create a definitions page for the workbook.

## Usage

``` r
definitionstab(
  contentslink = NULL,
  gridlines = "Yes",
  colwid_spec = NULL,
  extracols = NULL
)
```

## Arguments

- contentslink:

  Define whether a link to the contents page is wanted (optional)

- gridlines:

  Define whether gridlines are present (optional)

- colwid_spec:

  Define widths of columns (optional)

- extracols:

  Define whether additional columns required (optional)

## Value

A worksheet of the definitions page for the workbook.

## Details

definitionstab creates the worksheet with information on definitions. If
definitions not wanted, then do not run the definitionstab function.
There are three parameters and they are optional and preset. Change
contentslink to "No" if you want a contents tab but do not want a link
to it in the definitions tab. Change gridlines to "No" if gridlines are
not wanted. Column widths are automatically set but the user can specify
the required widths in colwid_spec. Extra columns can be added by
setting extracols to "Yes" and creating a dataframe
extracols_definitions with the desired extra columns.

## Examples

``` r
accessibletablesR::workbook(
   covertab = "Yes", contentstab = "Yes", notestab = "Yes", definitionstab = "Yes", 
   autonotes = "Yes", 
   title = "Fuel consumption and aspects of car design and performance for various cars",
   creator = "An organisation")
                            
accessibletablesR::creatingtables(
   title = "Fuel consumption and aspects of car design and performance for various cars C",
   subtitle = "Cars",
   extraline1 = "Link to contents",
   extraline2 = "Link to notes",
   extraline3 = "Link to definitions",
   sheetname = "Table_3", table_data = dummydf, tablename = "thirdtable", headrowsize = 40,
   numdatacols = c(2:8,11:13), numdatacolsdp = c(1,0,1,0,2,1,2,0,0,3),
   othdatacols = c(9,10), datedatacols = 15, datedatafmt = "dd-mm-yyyy", 
   datenondatacols = 14, datenondatafmt = "yyyy-mm-dd", columnwidths = "specified",
   colwid_spec = c(18,18,18,15,17,15,12,17,12,13,23,22,12,12,12))
                                  
accessibletablesR::contentstable()

accessibletablesR::addnote(notenumber = "note1", 
   notetext = "Google is an internet search engine", applictabtext = "All", linktext1 = "Google",
               linktext2 = "https://www.ons.google.co.uk") 

accessibletablesR::notestab()

accessibletablesR::adddefinition(term = "Usual resident", 
   definition = "A usual resident is anyone who, on Census Day, 21 March 2021 was in the UK and 
                 had stayed or intended to stay in the UK for a period of 12 months or more, or 
                 had a permanent UK address and was outside the UK and intended to be outside the
                 UK for less than 12 months.")

accessibletablesR::definitionstab()

accessibletablesR::coverpage(
  title = "Fuel consumption and aspects of car design and performance for various cars",
  intro = "Some made up data about cars",
  about = "The output of an example of how to use accessibletablesR",
  source = "R mtcars",
  relatedlink = "https://www.rdocumentation.org/packages/datasets/versions/3.6.2/topics/mtcars)",
  relatedtext = "mtcars: Motor trend car road tests",
  dop = "26 October 2023",
  blank = "There should be no blank cells",
  names = "Your name",
  email = "yourname@emailprovider.com",
  phone = "01111 1111111111111",
  reuse = "Yes", govdept = NULL)
                             
accessibletablesR::savingtables("D:/mtcarsexample.xlsx", odsfile = "Yes", deletexlsx = "No")
#> Warning: cannot create file 'D:/mtcarsexample.xlsx', reason 'No such file or directory'
#> Error in convert_to_ods(filename): File not found
```
