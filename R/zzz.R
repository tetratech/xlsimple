.onLoad <- function(libname = find.package("xlsimple"), pkgname = "xlsimple"){

  # declaration of global variables (http://stackoverflow.com/questions/9439256)
  if(getRversion() >= "2.15.1")
    utils::globalVariables(c("begin", "XL.wb"))
  invisible()

}

.onAttach <-  function(libname = find.package("xlsimple"), pkgname = "xlsimple"){
  packageStartupMessage("Notice: This package is a simple wrapper for selected XLConnect functions and is provided 'AS IS.'")
}
