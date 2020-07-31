
<!-- README.md is generated from README.Rmd. Please edit that file -->

# ‘xlsimple’

    #> Last Update: 2020-07-31 09:07:59

Provides a simple wrapper for some ‘XLConnect’ functions. Individual
sheets can include a description on the first row to remind user what is
in the data set. Auto filters and freeze rows are turned on. A brief
readme file is created that provides a summary listing of the created
sheets and where provided the description.

Included one function, read\_all(), from the package ‘loadxls’ by Carles
Hernandez-Ferrer (carleshf). This was in a GitHub repository, but it
does not exist as of 2020-07-20. The original blog post is below.
<https://carleshf87.wordpress.com/2016/01/24/loadxls-r-package/>

## Purpose

To quickly read/write ‘Excel’ worksheet names to/from R. Includes some
default options that make it easier on the user.

## Installation

You can install the released version of ‘baytrends’ from
[CRAN](https://CRAN.R-project.org) with:

``` r
install.packages("xlsimple")
```

And the development version from [GitHub](https://github.com) with:

    # install.packages("devtools")
    devtools::install_github("tetratech/xlsimple")
