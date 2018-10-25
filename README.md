# Office-Tools-for-R
Basic functionality to mangle Office files in R

##### Open as an existing project in RStudio and build.

Consists of four functions (version 0.0.5):

* parseDocxTable() - for parsing simple tables from Docx files
* parseDocxTableComplex() - for parsing more complex tables from Docx files
* writeXlsxTables() - for writing parsed table to Xlsx file
* docxToXlsxWizard() - function to call all above mentioned functions

It is based on Docx XML structure, only Docx files are working, not Doc due different structure of the file.