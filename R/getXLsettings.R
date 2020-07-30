#' Get Default 'Excel' cell styles
#'
#' Get Default 'Excel' cell styles
#'
#' @param fname 'Excel' file name
#'
#' @details If no 'Excel' file is specified then the built in cell styles are used.
#'   User-supplied 'Excel' file would need to have the following cell styles
#'   specified: xl.descrip, xl.normal, xl.header, xl.header.wrap, xl.hyperlink.
#'
#' @examples
#' XL.wb <- getXLsettings()
#' XL.wb <- addXLsheetStd(XL.wb, mtcars)
#' XL.wb <- addXLsheetStd(XL.wb, mtcars, "mtcars1")
#' XL.wb <- addXLsheetStd(XL.wb, mtcars, "mtcars2", "Standard mtcars data frame")
#' XL.wb$pName <- "ProjName" # optional, blank if not included
#' XL.wb$pDesc <- "ProjDesc" # optional, blank if not included
#' saveXLworkbook(XL.wb, file.path(tempdir(), 'myXLfile.xlsx'), timeStamp=FALSE, clean=FALSE)
#' saveXLworkbook(XL.wb, file.path(tempdir(), 'myXLfile.xlsx'), timeStamp=TRUE,  clean=FALSE)
#' saveXLworkbook(XL.wb, file.path(tempdir(), 'myXLfile.xlsx'), timeStamp=TRUE,  clean=TRUE)
#' saveXLworkbook(XL.wb, file.path(tempdir(), 'myXLfile.xlsx'))
#'
#' @return list with workbook and default cell styles
#'
#' @export
getXLsettings <-function(fname=NA) {

  # Pick up file name
  if(is.na(fname)) {
    # use built in settings
    myPackage <- 'xlsimple'
    fname <- system.file("extdata", paste0(myPackage, ".xlsx"), package = myPackage)
  } else {
    # use settings based on user supplied file
    fname <- .findFile(folder='.', file=fname )
    if(is.na(fname)) stop('No file not found.')
  }

  # Load worksheet
  wb <- XLConnect::loadWorkbook(fname)
  getXLsettings <- list(wb                = wb,
                        sheet_names       = XLConnect::getSheets(wb),
                        sheet_template    = XLConnect::getSheets(wb)[1],
                        style.descrip     = XLConnect::getCellStyle (wb , "xl.descrip"),
                        style.normal      = XLConnect::getCellStyle (wb , "xl.normal"),
                        style.header      = XLConnect::getCellStyle (wb , "xl.header"),
                        style.hyperlink   = XLConnect::getCellStyle (wb , "xl.hyperlink"),
                        style.header.wrap = XLConnect::getCellStyle (wb , "xl.header.wrap")
                        )

  # Add readme tab
  getXLsettings <- .addXLreadme(getXLsettings)

  return(getXLsettings)

}
