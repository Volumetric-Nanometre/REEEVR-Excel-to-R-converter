#' Excel EXPON.DIST function
#'
#' return correct norm distribution
#' @param x value
#' @param lambda scaling
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return pexp or dexp
#'
#' @examples
#' expon.dist <- excel__xlfn.expon.dist(0.2,100,True)
#'
#' @export
excel__xlfn.expon.dist <- function(x, lambda, cumulative) {

  if(cumulative){
    return(pexp(x,lambda))

  }
  else{
    return(dexp(x,lambda))
  }

}

#' Excel EXPONDIST function
#'
#' return correct norm distribution
#' @param x value
#' @param lambda scaling
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return pexp or dexp
#'
#' @examples
#' expondist <- excel_expondist(0.2,100,True)
#'
#' @export
excel_expondist <- function(x, lambda, cumulative) {

  if(cumulative){
    return(pexp(x,lambda))

  }
  else{
    return(dexp(x,lambda))
  }

}
