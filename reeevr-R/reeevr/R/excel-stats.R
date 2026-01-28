#' Convert AVERAGE(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' meanA <- excel_average(c(1,2,3,4))
#'
#' @export
excel_average <- function(...) {

  return(mean(...))
}

#' Convert AVERAGEA(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' meanA <- excel_averagea(c(1,2,3,4,"TEST",NA))
#'
#' @export
excel_averagea <- function(...) {

  temp  <- vector()
  for (val in ...){
    if( is.numeric(val) ){
      temp <- c(temp, val)
    }
  }
  return(mean(temp))
}

#' Convert AVERAGEIF(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' meanif <- excel_averageif(c(1,2,3,4))
#'
#' @export
excel_averageif <- function(...) {

  return(mean(...))
}

#' Convert AVERAGEIFs(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' meanifs <- excel_averageifs(c(1,2,3,4))
#'
#' @export
excel_averageifs <- function(...) {

  return(mean(...))
}

#' Convert COUNT(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' count <- excel_count(c(1,2,3,4))
#'
#' @export
excel_count <- function(...) {

  return()
}

#' Convert COUNTA(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' counta <- excel_counta(c(1,2,3,4))
#'
#' @export
excel_counta <- function(...) {

  return()
}

#' Convert COUNTBLANK(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' countblank <- excel_countblank(c(1,2,3,4))
#'
#' @export
excel_countblank <- function(...) {

  return()
}

#' Convert COUNTIF(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' countif <- excel_countif(c(1,2,3,4))
#'
#' @export
excel_countif <- function(...) {

  return()
}

#' Convert COUNTIFS(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' countifs <- excel_countifs(c(1,2,3,4))
#'
#' @export
excel_countifs <- function(...) {

  return()
}

#' Convert STDEV.S(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' stdev.s <- excel__xlfn.stdev.s(c(1,2,3,4))
#'
#' @export
excel__xlfn.stdev.s <- function(...) {

  return()
}

#' Convert STDEV.P(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' stdev.p <- excel__xlfn.stdev.p(c(1,2,3,4))
#'
#' @export
excel__xlfn.stdev.p <- function(...) {

  return()
}

#' Convert STDEVA(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' stdeva <- excel_stdeva(c(1,2,3,4))
#'
#' @export
excel_stdeva <- function(...) {

  return()
}

#' Convert STDEVP(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' stdevp <- excel_stdevp(c(1,2,3,4))
#'
#' @export
excel_stdevp <- function(...) {

  return()
}

#' Convert STDEVPA(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' stdevpa <- excel_stdevpa(c(1,2,3,4))
#'
#' @export
excel_stdevpa <- function(...) {

  return()
}

#' Convert VAR(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' var <- excel_var(c(1,2,3,4))
#'
#' @export
excel_var <- function(...) {

  return()
}

#' Convert VAR.P(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' var.p <- excel__xlfn.var.p(c(1,2,3,4))
#'
#' @export
excel__xlfn.var.p <- function(...) {

  return()
}

#' Convert VAR.s(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' var.s <- excel__xlfn.var.s(c(1,2,3,4))
#'
#' @export
excel__xlfn.var.s <- function(...) {

  return()
}

#' Convert VARA(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' vara <- excel_vara(c(1,2,3,4))
#'
#' @export
excel_vara <- function(...) {

  return()
}

#' Convert VARP(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' varp <- excel_varp(c(1,2,3,4))
#'
#' @export
excel_varp <- function(...) {

  return()
}

#' Convert VARPA(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return
#'
#' @examples
#' varpa <- excel_varpa(c(1,2,3,4))
#'
#' @export
excel_varpa <- function(...) {

  return()
}
