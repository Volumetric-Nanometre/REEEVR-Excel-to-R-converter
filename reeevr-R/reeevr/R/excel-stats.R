#' Generic sd
#'
#' @param x vector of values to calculate standard deviation
#' @param na.rm TRUE/FALSE - ignore NA in list
#' @param sample TRUE/FALSE - when true returns sample standard deviation. When false returns population.
#' @return Returns standard deviation of input vector
#'
#' @examples
#' popsd = generic_sd(c(1,2,3,4,5,6), sample = FALSE)
#' samplesd = generic_sd(c(1,2,3,4,5,6))
#' @export
generic_sd <- function(x, na.rm = FALSE, sample  = TRUE) {

  if(sample)
    return (sd(x,na.rm))
  else{
    n = length(x)
    return (sqrt((n-1)/n) * sd(x,na.rm))
  }

}

#' Generic var
#'
#' @param x vector of values to calculate variance
#' @param na.rm TRUE/FALSE - ignore NA in list
#' @param sample TRUE/FALSE - when true returns sample variance. When false returns population.
#' @return Returns variance of input vector
#'
#' @examples
#' popvar = generic_var(c(1,2,3,4,5,6), sample = FALSE)
#' samplevar = generic_var(c(1,2,3,4,5,6))
#' @export
generic_var <- function(x, na.rm = FALSE, sample  = TRUE) {

  if(sample)
    return (var(x,na.rm=na.rm))
  else{
    n = length(x)
    return ((n-1)/n * var(x,na.rm=na.rm))
  }

}


#' Convert MEDIAN(...)
#'
#' @param ... list
#' @return median of values
#'
#' @examples
#' medianSingleVars <- excel_median(0, 3, 5, 2.5, -10)
#' medianListVars <- excel_median(list(0, 3, 5, 2.5, -10))
#' medianMultiListVars <- excel_median(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#' @export
excel_median <- function(...) {

  args = list(...)
  if(length(args) == 1)
    return (apply(do.call(rbind, ...), 2, median))

  else
   return (excel_median(args))

}


#' Convert AVERAGE(...)
#'
#' @param ... list of values
#' @return mean of values
#'
#' @examples
#' meanSingleVars <- excel_average(0, 3, 5, 2.5, -10)
#' meanListVars <- excel_average(list(2, 3, 5, 2.5, -10))
#' meanMultiListVars <- excel_average(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#' @export
excel_average <- function(...) {

  args = list(...)
  if(length(args) == 1)
    return (apply(do.call(rbind, ...), 2, mean))

  else
    return (excel_average(args))

}

#' Convert STDEV(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return sample standard deviation
#'
#' @examples
#' stdevSingleVars <- excel_stdev(0, 3, 5, 2.5, -10)
#' stdevListVars <- excel_stdev(list(2, 3, 5, 2.5, -10))
#' stdevMultiListVars <- excel_stdev(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#'
#' @export
excel_stdev <- function(...) {

  args = list(...)
  if(length(args) == 1)
    return (apply(do.call(rbind, ...), 2, generic_sd))

  else
    return (excel_stdev(args))
}

#' Convert STDEVP(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return population standard deviation
#'
#' @examples
#' stdevpSingleVars <- excel_stdevp(0, 3, 5, 2.5, -10)
#' stdevpListVars <- excel_stdevp(list(2, 3, 5, 2.5, -10))
#' stdevpMultiListVars <- excel_stdevp(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#'
#' @export
excel_stdevp <- function(...) {

  args = list(...)
  if(length(args) == 1){
    return (apply(do.call(rbind, ...), 2, generic_sd, sample = FALSE))

  }
  else
    return (excel_stdevp(args))
}

#' Convert STDEV.S(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return sample standard deviation
#'
#' @examples
#' stdev.sSingleVars <- excel__xlfn.stdev.s(0, 3, 5, 2.5, -10)
#' stdev.sListVars <- excel__xlfn.stdev.s(list(2, 3, 5, 2.5, -10))
#' stdev.sMultiListVars <- excel__xlfn.stdev.s(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#'
#' @export
excel__xlfn.stdev.s <- function(...) {

  args = list(...)
  if(length(args) == 1)
    return (apply(do.call(rbind, ...), 2, generic_sd))

  else
    return (excel__xlfn.stdev.s(args))
}

#' Convert STDEV.P(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return populations tandard deviation
#'
#' @examples
#' stdev.pSingleVars <- excel__xlfn.stdev.p(0, 3, 5, 2.5, -10)
#' stdev.pListVars <- excel__xlfn.stdev.p(list(2, 3, 5, 2.5, -10))
#' stdev.pMultiListVars <- excel__xlfn.stdev.p(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#'
#' @export
excel__xlfn.stdev.p <- function(...) {

  args = list(...)
  if(length(args) == 1)
    return (apply(do.call(rbind, ...), 2, generic_sd, sample = FALSE))

  else
    return (excel__xlfn.stdev.p(args))
}

#' Convert VAR(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return sample variance of list
#'
#' @examples
#' varSingleVars <- excel_var(0, 3, 5, 2.5, -10)
#' varListVars <- excel_var(list(2, 3, 5, 2.5, -10))
#' varMultiListVars <- excel_var(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#'
#' @export
excel_var <- function(...) {

  args = list(...)
  if(length(args) == 1)
    return (apply(do.call(rbind, ...), 2, generic_var))

  else
    return (excel_var(args))

}

#' Convert VAR.S(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return sample variance of list
#'
#' @examples
#' var.sSingleVars <- excel__xlfn.var.s(0, 3, 5, 2.5, -10)
#' var.sListVars <- excel__xlfn.var.s(list(2, 3, 5, 2.5, -10))
#' var.sMultiListVars <- excel__xlfn.var.s(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#'
#' @export
excel__xlfn.var.s <- function(...) {

  args = list(...)
  if(length(args) == 1)
    return (apply(do.call(rbind, ...), 2, generic_var))

  else
    return (excel__xlfn.var.s(args))

}

#' Convert VARP(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return population variance of list
#'
#' @examples
#' varpSingleVars <- excel_varp(0, 3, 5, 2.5, -10)
#' varpListVars <- excel_varp(list(2, 3, 5, 2.5, -10))
#' varpMultiListVars <- excel_varp(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#'
#' @export
excel_varp <- function(...) {

  args = list(...)
  if(length(args) == 1)
    return (apply(do.call(rbind, ...), 2, generic_var, sample = FALSE))

  else
    return (excel_varp(args))

}

#' Convert VAR.P(...)
#'
#' @param ... list of values - allows non numeric characters
#' @return population variance of list
#'
#' @examples
#' var.pSingleVars <- excel__xlfn.var.p(0, 3, 5, 2.5, -10)
#' var.pListVars <- excel__xlfn.var.p(list(2, 3, 5, 2.5, -10))
#' var.pMultiListVars <- excel__xlfn.var.p(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#'
#' @export
excel__xlfn.var.p <- function(...) {

  args = list(...)
  if(length(args) == 1)
    return (apply(do.call(rbind, ...), 2, generic_var, sample = FALSE))

  else
    return (excel__xlfn.var.p(args))

}

