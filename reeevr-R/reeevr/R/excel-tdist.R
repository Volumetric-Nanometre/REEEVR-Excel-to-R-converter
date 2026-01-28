#' Excel T.DIST function
#'
#' return correct norm distribution
#' @param x value
#' @param df degrees of freedom
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return pt or dt
#'
#' @examples
#' t.dist <- excel__xlfn.t.dist(0.2,10,True)
#'
#' @export
excel__xlfn.t.dist <- function(x, df, cumulative) {

  if(cumulative){
    return(pt(x,df))

  }
  else{
    return(dt(x,df))
  }

}

#' Excel TDIST function
#'
#' return correct norm distribution
#' @param x value
#' @param df degrees of freedom
#' @param tails TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return pt or dt
#'
#' @examples
#' tdist <- excel_tdist(0.2,10,1)
#' tdist <- excel_tdist(0.2,10,2)
#'
#' @export
excel_tdist <- function(x, df, tails) {

  if(tails == 1){
    return(pt(x,df, lower.tail = FALSE))

  }
  if(tails == 2){
    return(2*pt(-abs(x),df))
  }

}

#' Excel T.DIST.2T function
#'
#' return correct norm distribution
#' @param x value
#' @param df degrees of freedom
#' @return pt
#'
#' @examples
#' t.dist.2t <- excel__xlfn.t.dist.2t(0.2,10)
#'
#' @export
excel__xlfn.t.dist.2t <- function(x, df) {

  return(2*pt(-abs(x),df))

}

#' Excel T.DIST.RT function
#'
#' return correct norm distribution
#' @param x value
#' @param df degrees of freedom
#' @return pt
#'
#' @examples
#' t.dist.rt <- excel__xlfn.t.dist.rt(0.2,10)
#'
#' @export
excel__xlfn.t.dist.rt <- function(x, df) {

  return(pt(x,df, lower.tail = FALSE))

}

#' Excel T.INV function
#'
#' return correct norm distribution
#' @param x value
#' @param df degrees of freedom
#' @return qt
#'
#' @examples
#' t.inv <- excel__xlfn.t.inv(0.2,10)
#'
#' @export
excel__xlfn.t.inv <- function(x, df) {

  return(-qt(x,df, lower.tail = FALSE))

}

#' Excel T.INV.2T function
#'
#' return correct norm distribution
#' @param x value
#' @param df degrees of freedom
#' @return qt
#'
#' @examples
#' t.inv.2t <- excel__xlfn.t.inv.2t(0.2,10)
#'
#' @export
excel__xlfn.t.inv.2t <- function(x, df) {

  return(-qt(x/2, df))


}
