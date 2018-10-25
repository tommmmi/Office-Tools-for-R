#' Parse formatted tables from Word DOCX files
#'
#' Parse formatted tables from Word DOCX files. Note that only DOCX files are
#' allowed (DOC is not supported).
#' Also note that the table should not contain any merged or otherwise manipulated
#' cells in it. The function will fail to parse the rows and columns correctly.
#'
#' @import xml2
#' @param word.docx Word DOCX file name (with file extensions '.docx')
#' @examples \dontrun{
#'   all.tables <- parseDocxTable("Demographic Tables.docx")
#' }
#' @export
parseDocxTable <- function(word.docx) {
  # create necessary temps
  tmpdir <- tempdir()
  tmpfile <- tempfile(tmpdir=tmpdir, fileext=".zip")

  # copy actual Word file to temp dir
  file.copy(word.docx, tmpfile)

  # Docx is standard zip file, unzip it
  unzip(tmpfile, exdir=sprintf("%s/docdata", tmpdir))

  # Read document body from certain XML file inside the Docx/Zip
  doc.base <- read_xml(sprintf("%s/docdata/word/document.xml", tmpdir))

  # remove unnecessary temps
  unlink(tmpfile)
  unlink(sprintf("%s/docdata", tmpdir), recursive=T)

  # extract namespaces for finding 'table' objects from the file
  namespace <- xml_ns(doc.base)

  # find all table objects
  doc.tables <- xml_find_all(doc.base, ".//w:tbl", ns=namespace)

  # parse table data to data.frame and return it
  lapply(doc.tables, function(tbl.extract) {

    # get cells
    cells <- xml_find_all(tbl.extract, "./w:tr/w:tc", ns=namespace)

    # get rows
    rows <- xml_find_all(tbl.extract, "./w:tr", ns=namespace)

    # get actual data
    table.data <- data.frame(matrix(xml_text(cells), ncol=(length(cells)/length(rows)), byrow=T), stringsAsFactors=F)

    # insert variable names
    colnames(table.data) <- table.data[1,]

    # remove one row from the top (which contains the variable names and are not actual data)
    table.data <- table.data[-1,]

    # remove row names
    rownames(table.data) <- NULL

    # return
    table.data
  }
  )
}



#' Parse formatted (more complex) tables from Word DOCX files
#'
#' Parse formatted tables from Word DOCX files. Note that only DOCX files are
#' allowed (DOC is not supported).
#' Tables can have merged cells but pay attention that the output matches the
#' input table.
#'
#' @import docxtractr
#' @param word.docx Word DOCX file name (with file extensions '.docx')
#' @examples \dontrun{
#'   all.tables <- parseDocxTableComplex("Demographic Tables.docx")
#' }
#' @export
parseDocxTableComplex <- function(word.docx) {
  # Read document body
  doc.base <- read_docx(word.docx)
  doc.tables <- list()

  # loop tables
  for(i in 1:docx_tbl_count(doc.base)) {
    doc.tables[[i]] <- docx_extract_tbl(doc.base, i, header=T)
  }

  # return
  doc.tables
}



#' Parse formatted (more complex) tables from Word DOCX files
#'
#' Parse formatted tables from Word DOCX files. Note that only DOCX files are
#' allowed (DOC is not supported).
#' Tables can have merged cells but pay attention that the output matches the
#' input table.
#'
#' @import xlsx
#' @param data.list Individual tables as data.frames inside a list (list elements are all data.frames)
#' @param output.file.name Word XLSX file name (without file extensions '.xlsx')
#' @examples \dontrun{
#'   writeXlsxTables(demo.tables, "Demographic tables")
#' }
#' @export
writeXlsxTables <- function(data.list, output.file.name) {
  # prepare output Excel file
  output.workbook <- createWorkbook(type="xlsx")

  # create sheets to workbook
  sheet.names <- paste0("Table ", seq(1:length(data.list)))
  output.sheets <- list()

  # fill in the tables
  for(i in 1:length(sheet.names)) {
    output.sheets[[i]] <- createSheet(output.workbook, sheetName = sheet.names[[i]])
    addDataFrame(data.list[[i]], sheet=output.sheets[[i]], startRow=1, startColumn=1, row.names=F)
  }

  # write the actual Excel file
  saveWorkbook(output.workbook, paste0(getwd(), "/", output.file.name, ".xlsx"))
}



#' Automatically transform all tables in Word file to Excel file
#'
#' Parse formatted tables from Word DOCX files and move them to an Excel file.
#'
#' @import xml2 docxtractr xlsx
#' @param word.docx Word DOCX file name (without file extensions '.docx')
#' @examples \dontrun{
#'   docxToXlsxWizard("Simple table")
#' }
#' @export
docxToXlsxWizard <- function(docx.file) {
  parsed.tables.list <- parseDocxTableComplex(paste0(docx.file, ".docx"))
  writeXlsxTables(parsed.tables.list, paste0("Extracted Tables from '", docx.file, ".docx' File"))
}
