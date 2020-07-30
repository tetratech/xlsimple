#' Add a sheet to the 'Excel' workbook
#'
#' Add a sheet to the 'Excel' workbook
#'
#' @param wbList list with workbook and default cell styles (i.e., output from getXLsettings)
#' @param df data frame to output to sheet
#' @param sheetName sheet name
#' @param descrip description of sheet
#'
#' @details The \code{sheetName} is the name that will be used for the sheet. When
#' \code{sheetName} is not specified, the name of the \code{df} will be used. The
#' \code{descrip} is a character string that will be placed into cell A1 of the sheet
#' and is best used to provide a brief description of the sheet; and the dataframe
#' will begin at cell A3. If \code{descrip} is not specified, then the dataframe will
#' begin at cell A1. The outputted table will have filters turned on and table headers
#' and first column frozen.
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
addXLsheetStd <- function(wbList=XL.wb, df=NA, sheetName=NA, descrip=NA) {
# wbList=XL.wb; df=mtcars; sheetName=NA; descrip="mtcars data"

  # use name of df as sheetName if sheetName not provided
  if(is.na(sheetName)) sheetName = deparse(substitute(df))
  XLConnect::cloneSheet(wbList[['wb']], wbList[['sheet_template']], sheetName)

  # base cell locations
  myStartRow <- 1
  myStartCol <- 1
  myRef      <- paste0(LETTERS[myStartCol],myStartRow) #"A1"

  # write out description if provided
  if(!is.na(descrip)) {
    XLConnect::writeWorksheet  (wbList[['wb']], data=descrip, sheet=sheetName, startRow=myStartRow,
                     startCol=myStartCol, header=FALSE)
    XLConnect::setCellStyle    (wbList[['wb']], sheet=sheetName, row=myStartRow, col=myStartCol,
                     cellstyle=wbList[['style.descrip']])
    myStartRow <- 3
    myStartCol <- 1
    myRef      <- paste0(LETTERS[myStartCol],myStartRow) #"A3"
  }

  # write out table
  XLConnect::writeWorksheet  (wbList[['wb']], data=df, sheet=sheetName, startRow=myStartRow,
                   startCol=myStartCol, header=TRUE)
  XLConnect::setCellStyle    (wbList[['wb']], sheet=sheetName, row=myStartRow, col=1:dim(df)[2],
                   cellstyle=wbList[['style.header.wrap']])
  suppressWarnings(
    XLConnect::setCellStyle    (wbList[['wb']], sheet=sheetName, row =(myStartRow+1):dim(df)[1] , col=1:dim(df)[2],
                   cellstyle=wbList[['style.normal']]))
  XLConnect::setAutoFilter   (wbList[['wb']], sheet=sheetName, reference = XLConnect::aref(myRef, dim(df)))
  XLConnect::createFreezePane(wbList[['wb']], sheet = sheetName, myStartCol+1, myStartRow+1)
  XLConnect::setColumnWidth  (wbList[['wb']], sheet = sheetName, column = 1:dim(df)[2], width = 4050)
  XLConnect::setMissingValue (wbList[['wb']], value = "")

  # write readme entry
  N <- 11
  N <- N + length(setdiff(XLConnect::getSheets(wbList[['wb']]),
                          c(wbList[['myRead']] , wbList[['sheet_names']])))
  XLConnect::writeWorksheet  (wbList[['wb']], data=sheetName, sheet=wbList[['myRead']], startRow=N,
                              startCol=1, header=FALSE)
  XLConnect::writeWorksheet  (wbList[['wb']], data=descrip, sheet=wbList[['myRead']], startRow=N,
                              startCol=2, header=FALSE)
  XLConnect::setCellStyle    (wbList[['wb']], sheet=wbList[['myRead']], row=N, col=2,
                              cellstyle=wbList[['style.normal']])
  XLConnect::setCellStyle    (wbList[['wb']], sheet=wbList[['myRead']], row=N, col=1,
                              cellstyle=wbList[['style.header']])
  ## formula for link to worksheets
  formula.Link <- paste0('HYPERLINK("["&$B$6&"]"&A',N,'&"!A1",A',N,')')
  XLConnect::setCellFormula   (wbList[['wb']], sheet=wbList[['myRead']],row=N,col=3,formula.Link)
  XLConnect::setCellStyle     (wbList[['wb']], sheet=wbList[['myRead']],row=N, col=3,
                              cellstyle=wbList[['style.hyperlink']])


  return(wbList)
}






