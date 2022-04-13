# GoogleSheet2GoogleDocs

Google script to create multiple files (Google Docs) by using Google Spreadsheet as database.
Useful for small business automation. Can easily create to create bunch of contracts, invoices etc.

## Goal

Simple office tasks automation tool.

Say, you have Google Spreadsheet formatted as table:
First row contains column names, others rows contain data.
The data fields inserted into the document via {{field}} substitution.
The note on first row first column name should contain url/id of google document that will be used as template

### Example

#### Source data

Let assume that data is located in some spreadsheet, that have this `Code.gs` installed

| CustomerName   | ContractId | ContractDate |
|----------------|------------|--------------|
| John Doe       | CN-1228    | 12.08.2019   |
| Peter Norton   | CN-12/56   | 01.02.2020   |

#### The template document

Let have `ConfirmationLetter` Google Document as template:

| Dear {{CustomerName}}, your contract {{ContractId}} is approved at {{ContractDate}} |
|-------------------------------------------------------------------------------------|

#### Template linking

The url (or id) of template file `ConfirmationLetter` should be put in the `note` on first row first column name

#### Results 

With this tool you may create separate Google Document (or Google Spreadsheet) for each of row, using some document as template.
To execute tool, you have just to use menu "GoogleSheet2GoogleDoc" - "Fill docs with {{template}} fields"

Result file for row 1 `ConfirmationLetter_filled_1`:
| Dear John Doe, your contract CN-1228 is approved at 12.08.2019                     |
|------------------------------------------------------------------------------------|

Result file for row 2 `ConfirmationLetter_filled_2`:
| Dear Peter Norton, your contract CN-12/56 is approved at 01.02.2020                |
|------------------------------------------------------------------------------------|

## How it works

Explodes spreadsheet data to multiple documents/spreadsheets.
Uses first table on active sheet as a data source for filling multiple files by template(s)
First row on sheet should contain field names (until first empty column).
Each data row will produce own output file by template file.
Each {{field}} in document will be replaced with data row cell content.
Stops at end of sheet or first empty data row.

### How to specify template document

The template cam be Google Document or Google Spreadsheet.
That is, you may create bunch of spreadsheets from single spreadsheet.

Template document url/id:

a). Extracted from note on first column

b). Extracted from special field !template inside a row, to make this row have alternate template

### Extra customization

Processing can be customized by specifying special columns:

a). !skip - if exist, this row is skipped if value contains TRUE, Y, y, true

b). !template - if exist, will specify template file url/id for this row (used if not empty)

c). !output - if exist, specify output file name (default will be {{templateName}}_filled_{{rowNum}}

d). !output.folder - if exist, will specify sub folder name to save files to ('.' = same as template file path)

e). !pdf.folder - if exist, will specify sub folder name to save .pdf files to ('.' = same as template file path)

### Special treatments of multiline values

In multiline fields, if line started with '#).', the '#' sign is replaced by item number (starting form '1).'

## Other notes

a). template can be document or spreadsheet (that is, generation of spreadsheets supported also!)

b). if template is a spreadsheet, hidden sheets are skipped and formula cells are not changed

c). the system may generate .pdf files from output files, if !pdf.folder column exist (experimental)

d). the note of the first field may contains additional options in form OptionName=value (each on separate line)

*** Supported options are: 'SavePDF=Y'
