#' Collection of R functions for different Office files
#'
#' This package contains a few functions to extract table-formatted tables
#' from Microsoft Office Word documents (not older DOC files). There is no
#' need for manual manipulation of the Word files before using this R package.
#'
#' The functions included in \pkg{OfficeTools} R package are:
#'
#' \code{\link{parseDocxTable}} is used for parsing table-formatted tables from Word
#' file to a list object containing data.frames; works only for simple tables
#' without merged cells or otherwise altered table structure
#'
#' \code{\link{parseDocxTableComplex}} is used for parsing table-formatted tables
#' from Word file to a list object containing data.frames; works also with complex
#' tables with merged cells or otherwise altered table structure
#'
#' \code{\link{writeXlsxTables}} is used for automatically generating Excel file
#' from a 'list' type R object containing data.frames
#'
#' \code{\link{docxToXlsxWizard}} is used for automatically extracting all table-
#' formatted tables from a Word file and saving them to similarly named Excel
#' file
"_PACKAGE"
