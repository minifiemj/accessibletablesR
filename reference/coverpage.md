# accessibletablesR::coverpage

Create a cover page for the workbook.

## Usage

``` r
coverpage(
  title,
  intro = NULL,
  about = NULL,
  source = NULL,
  relatedlink = NULL,
  relatedtext = NULL,
  dop = NULL,
  blank = NULL,
  names = NULL,
  email = NULL,
  phone = NULL,
  reuse = NULL,
  govdept = NULL,
  cyear = NULL,
  gridlines = "Yes",
  extrafields = NULL,
  extrafieldsb = NULL,
  additlinks = NULL,
  addittext = NULL,
  colwid_spec = NULL,
  order = NULL
)
```

## Arguments

- title:

  Title for workbook

- intro:

  Introductory information (optional)

- about:

  About the data (optional)

- source:

  Data source(s) (optional)

- relatedlink:

  Link(s) to related publications (optional)

- relatedtext:

  Text to appear in place of associated link (optional)

- dop:

  Date of publication (optional)

- blank:

  Blank cell information (optional)

- names:

  Contact name (optional)

- email:

  Contact email (optional)

- phone:

  Contact phone number (optional)

- reuse:

  Define whether want to use default text on reuse of publication
  (optional)

- govdept:

  UK Government department name (optional)

- cyear:

  Crown copyright year (optional)

- gridlines:

  Define whether gridlines are present (optional)

- extrafields:

  Additional fields for the cover page (optional)

- extrafieldsb:

  Text to appear in additional fields (optional)

- additlinks:

  Additional links of relevance (optional)

- addittext:

  Text to appear in place of associated additional link (optional)

- colwid_spec:

  Define widths of columns (optional)

- order:

  List of fields in order of appearance wanted on cover page (optional)

## Value

A worksheet of the cover page for workbook

## Details

coverpage function will create a cover page for the front of the
workbook. If a cover page is not wanted, do not run the coverpage
function. The only compulsory parameter is title. All other parameters
are optional and preset, only populate if they are wanted. intro:
Introductory information / about: About these data / dop: Date of
publication. source: Data source(s) used / blank: Information about why
some cells are blank, if necessary. relatedlink and relatedtext - any
publications associated with the data (relatedlink is the actual
hyperlink, relatedtext is the text you want to appear to the user).
names: Contact name / email: Contact email / phone: Contact telephone.
reuse: Set to "Yes" if you want the information displayed about the
reuse of the data (will automatically be populated). cyear: If reuse is
set to "Yes" then set cyear to the year you want for the Crown copyright
or leave as NULL if you want the current year or set to "No" if you do
not want the year displayed govdept: Default is "ONS" but if want reuse
information without reference to ONS change govdept. extrafields: Any
additional fields that the user wants present on the cover page.
extrafieldsb: The text to go in any additional fields. Only one row per
field. extrafields and extrafields must be vectors of the same length.
additlinks: Any additional hyperlinks the user wants. addittext: The
text to appear over any additional hyperlinks. additlinks and addittext
must be vectors of the same length. order: If the user wants the cover
page to be ordered in a specific way, list the fields in a vector with
each field name in speech marks. e.g., order = c("intro", "about",
relatedlink", "names", "phone", "email", "extrafields"). Change
gridlines to "No" if gridlines are not wanted. Column width
automatically set unless user specifies a value in colwid_spec. intro,
about, source, dop, blank, names, phone can be set to hyperlinks - e.g.,
source = "\[ONS\](https://www.ons.gov.uk)".

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
  reuse = "Yes", govdept = NULL, cyear = 2023)
                             
accessibletablesR::savingtables("D:/mtcarsexample.xlsx", odsfile = "Yes", deletexlsx = "No")
#> Warning: cannot create file 'D:/mtcarsexample.xlsx', reason 'No such file or directory'
#> Error in convert_to_ods(filename): File not found
```
