#' Add a readme sheet to the 'Excel' workbook
#'
#' Add a readme sheet to the 'Excel' workbook
#'
#' @param wbList list with workbook and default cell styles (i.e., output from getXLsettings)
#'
#' @return list with workbook and default cell styles
#'
#' @export
.addXLreadme <- function(wbList=XL.wb) {
# df<-mtcars; sheetName=NA; descrip = "My table descrip"

  wbList[['myRead']] <- '_readme'

  if(!(wbList[['myRead']] %in%  XLConnect::getSheets(wbList[['wb']]))) {
    XLConnect::cloneSheet(wbList[['wb']], wbList[['sheet_template']], wbList[['myRead']])
    strReadMe.Names <- c("Documentation", "","Project Name:","Project Description:","","File:","Created:","Full Path:","Worksheet:","","Sheet")
    XLConnect::writeWorksheet(wbList[['wb']],
                              data=strReadMe.Names,
                              sheet=wbList[['myRead']], startRow=1, startCol=1, header=FALSE)
    XLConnect::writeWorksheet(wbList[['wb']],
                              data=c("Description"),
                              sheet=wbList[['myRead']], startRow=length(strReadMe.Names), startCol=2, header=FALSE)
    XLConnect::writeWorksheet(wbList[['wb']],
                              data=c("Link"),
                              sheet=wbList[['myRead']], startRow=length(strReadMe.Names), startCol=3, header=FALSE)

    XLConnect::setCellStyle  (wbList[['wb']], sheet=wbList[['myRead']], row=1:length(strReadMe.Names), col=1,
                                cellstyle=wbList[['style.descrip']])
    XLConnect::setCellStyle  (wbList[['wb']], sheet=wbList[['myRead']], row=length(strReadMe.Names), col=2:3,
                              cellstyle=wbList[['style.descrip']])
    XLConnect::setColumnWidth  (wbList[['wb']], sheet = wbList[['myRead']], column = 1:2, width = 6500)
    # Freeze Panes for "sheets" table
    XLConnect::createFreezePane(wbList[['wb']], sheet = wbList[['myRead']], 1, 12)

  }

  return(wbList)
}






