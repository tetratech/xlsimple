#' Load a full MS Excel File
#'
#' Given a MS Excel File it loads the full workbook creating as many variables as
#' sheets exists in the workbook.
#'
#' @param filename The path and the name to the MS Excel File.
#' @param environment Environment were the new variables will be located. By
#' default is parent environment.
#' @param verbose If set to \code{TRUE} useful messages are shown.
#' @param free.warnings If set to \code{TRUE} it shows any warnings
#' result of loading the content of the book's sheets.
#' @details This function was developed by Carles Hernandez-Ferrer
#'
#' https://github.com/carleshf/loadxls/blob/master/R/read_all.R
#'
#' @export
readXLworkbook <- function(filename, environment=parent.frame(), verbose=TRUE,
                     free.warnings=TRUE) {
  if(!file.exists(filename)) {
    stop("Given file '", filename, "' does not exists.")
  }
  wb <- XLConnect::loadWorkbook(filename, create=FALSE)
  trash <- lapply(XLConnect::getSheets(wb), function(sheet) {
    if(verbose) {
      message("Loading sheet '", sheet, "'.")
    }
    if(!free.warnings) {
      suppressWarnings({
        var <- XLConnect::readWorksheet(wb, sheet)
      })
    } else {
      var <- XLConnect::readWorksheet(wb, sheet)
    }
    assign(sheet, var, envir=environment)
  })
  rm(trash)
}
