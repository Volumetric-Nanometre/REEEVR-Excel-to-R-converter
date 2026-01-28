#' Excel F.DIST function
#'
#' return correct norm distribution
#' @param x value
#' @param df1 degree of freedom 1
#' @param df2 degree of freedom 2
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return pf or df
#'
#' @examples
#' f.dist <- excel__xlfn.f.dist(0.2,10,10,True)
#'
#' @export
excel__xlfn.f.dist <- function(x, df1, df2, cumulative) {

  if(cumulative){
    return(pf(x,df1,df2))

  }
  else{
    return(df(x,df1,df2))
  }

}

#' Excel F.DIST.RT function
#'
#' return correct norm distribution
#' @param x value
#' @param df1 degree of freedom 1
#' @param df2 degree of freedom 2
#' @return pf
#'
#' @examples
#' f.dist.rt <- excel__xlfn.f.dist.rt(0.2,10,10)
#'
#' @export
excel__xlfn.f.dist.rt <- function(x, df1, df2) {

  return(pf(x,df1,df2, lower.tail = FALSE))

}

#' Excel F.INV function
#'
#' return correct norm distribution
#' @param x probability
#' @param df1 degree of freedom 1
#' @param df2 degree of freedom 2
#' @return qf
#'
#' @examples
#' f.inv <- excel__xlfn.f.inv(0.2,10,10)
#'
#' @export
excel__xlfn.f.inv <- function(x, df1, df2) {

  return(qf(x,df1,df2))

}

#' Excel F.INV.RT function
#'
#' return correct norm distribution
#' @param x probability
#' @param df1 degree of freedom 1
#' @param df2 degree of freedom 2
#' @return qf
#'
#' @examples
#' f.inv.rt <- excel__xlfn.f.inv.rt(0.2,10,10)
#'
#' @export
excel__xlfn.f.inv.rt <- function(x, df1, df2) {

  return(qf(x,df1,df2, lower.tail = FALSE))

}
