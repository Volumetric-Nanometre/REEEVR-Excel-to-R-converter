#' Excel WEIBULL.DIST function
#'
#' @param x value to evaluate
#' @param alpha shape
#' @param beta scale
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return qbinom or dbinom
#'
#' @examples
#' weibull.dist <- excel__xlfn.weibull.dist(0.2,100,6,True)
#'
#' @export
excel__xlfn.weibull.dist <- function(x, alpha, beta, cumulative) {

  if(cumulative){
    return(pweibull(x, shape = alpha, scale = beta))

  }
  else{
    return(dweibull(x, shape = alpha, scale = beta))
  }

}

#' Excel WEIBULL function
#'
#' @param x value to evaluate
#' @param alpha shape
#' @param beta scale
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return qbinom or dbinom
#'
#' @examples
#' weibull <- excel_weibull(0.2,100,6,True)
#'
#' @export
excel_weibull <- function(x, alpha, beta, cumulative) {

  if(cumulative){
    return(pweibull(x, shape = alpha, scale = beta))

  }
  else{
    return(dweibull(x, shape = alpha, scale = beta))
  }

}

