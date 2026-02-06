#' Excel POISSON.DIST function
#'
#' return correct norm distribution
#' @param x Number of events
#' @param mean expectation value
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return ppois or dpois
#'
#' @examples
#' poisson.dist <- excel__xlfn.poisson.dist(10,10,True)
#'
#' @export
excel__xlfn.poisson.dist <- function(x, mean, cumulative) {

  if(cumulative){
    return(ppois(x, mean))

  }
  else{
    return(dpois(x, mean))
  }

}

#' Excel POISSON function
#'
#' return correct norm distribution
#' @param x Number of events
#' @param mean expectation value
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return ppois or dpois
#'
#' @examples
#' poisson <- excel_poisson(10,10,True)
#'
#' @export
excel_poisson <- function(x, mean, cumulative) {

  if(cumulative){
    return(ppois(x, mean))

  }
  else{
    return(dpois(x, mean))
  }

}
