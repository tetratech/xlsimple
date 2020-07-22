#' Save Excel Workbook to disk
#'
#' Save Excel Workbook to disk
#'
#' @param wbList list with workbook and default cell styles (i.e., output from getXLsettings)
#' @param fname Excel file name
#' @param timeStamp Logical field to include date/time stamp in file name (TRUE [default]).
#' @param clean Logical field indicating whether to remove original sheets in workbook
#' @examples
#'\dontrun{
#' XL.wb <- getXLsettings()
#' XL.wb <- addXLsheetStd(XL.wb, mtcars)
#' XL.wb <- addXLsheetStd(XL.wb, mtcars, "mtcars1")
#' XL.wb <- addXLsheetStd(XL.wb, mtcars, "mtcars2", "Standard mtcars data frame")
#' XL.wb$pName <- "ProjName" # optional, blank if not included
#' XL.wb$pDesc <- "ProjDesc" # optional, blank if not included
#' saveXLworkbook(XL.wb, 'myXLfile.xlsx', timeStamp=FALSE, clean=FALSE)
#' saveXLworkbook(XL.wb, 'myXLfile.xlsx', timeStamp=TRUE,  clean=FALSE)
#' saveXLworkbook(XL.wb, 'myXLfile.xlsx', timeStamp=TRUE,  clean=TRUE)
#' saveXLworkbook(XL.wb, 'myXLfile.xlsx')
#' }
#' @return n/a
#' @export
saveXLworkbook <- function(wbList, fname="xl.Out.xlsx", timeStamp=FALSE, clean=TRUE) {

  n <- format(Sys.time(),"%Y_%m_%d_%H%M%S")

  # insert time stamp into base file name if requested
  if (timeStamp) {
    #n <- format(Sys.time(),"%Y_%m_%d_%H%M%S")
    fExt  <- tools::file_ext(fname)
    fname <- tools::file_path_sans_ext(fname)
    fname <- paste0(fname,"_",n,".",fExt)
  }

  # write readme file info
  ## FileName and Created Date
  XLConnect::writeWorksheet  (wbList[['wb']], data=fname, sheet=wbList[['myRead']], startRow=6,
                              startCol=2, header=FALSE)
  XLConnect::writeWorksheet  (wbList[['wb']], data=n, sheet=wbList[['myRead']], startRow=7,
                              startCol=2, header=FALSE)
  XLConnect::setCellStyle    (wbList[['wb']], sheet=wbList[['myRead']], row=6:7, col=2,
                              cellstyle=wbList[['style.header']])
  ## ProjectName and Description
  XLConnect::writeWorksheet  (wbList[['wb']], data=wbList[['pName']], sheet=wbList[['myRead']], startRow=3,
                              startCol=2, header=FALSE)
  XLConnect::writeWorksheet  (wbList[['wb']], data=wbList[['pDesc']], sheet=wbList[['myRead']], startRow=4,
                              startCol=2, header=FALSE)
  ## FullPath and Worksheet (Tab)
  XLConnect::writeWorksheet  (wbList[['wb']], data="xFullPath", sheet=wbList[['myRead']], startRow=8,
                              startCol=2, header=FALSE)
  XLConnect::writeWorksheet  (wbList[['wb']], data="xWS", sheet=wbList[['myRead']], startRow=9,
                              startCol=2, header=FALSE)
  formula.FullPath <- 'LEFT(CELL("filename",B8),FIND("]",CELL("filename",B8)))'
  XLConnect::setCellFormula  (wbList[['wb']], sheet=wbList[['myRead']],row=8,col=2,formula.FullPath)
  formula.WS <- 'MID(CELL("filename",B9),FIND("]",CELL("filename",B9))+1,LEN(CELL("filename",B9))-FIND("]",CELL("filename",B9)))'
  XLConnect::setCellFormula  (wbList[['wb']], sheet=wbList[['myRead']],row=9,col=2,formula.WS)

  if(clean) {
    XLConnect::removeSheet(XL.wb[['wb']], wbList[['sheet_names']])
  }

  # save
    XLConnect::saveWorkbook (XL.wb[['wb']], fname)
}
