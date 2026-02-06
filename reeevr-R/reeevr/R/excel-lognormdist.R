#' Excel LOGNORM.DIST function
#'
#' return correct norm distribution
#' @param x list of random numbers
#' @param mean mean of distribution
#' @param sd standard deviation of distribution
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return qlnorm or dlnorm
#'
#' @examples
#' lnorm.dist <- excel__xlfn.lnorm.dist(0.2,100,6,True)
#'
#' @export
excel__xlfn.lognorm.dist <- function(x, mean, sd,cumulative) {

  if(cumulative){
    return(plnorm(x,mean, sd))

  }
  else{
    return(dlnorm(x,mean, sd))
  }

}

#' Excel LOGNORMDIST function
#'
#' return correct norm distribution
#' @param x list of random numbers
#' @param mean mean of distribution
#' @param sd standard deviation of distribution
#' @return qlnorm or dlnorm
#'
#' @examples
#' lnormdist <- excel_lnormdist(0.2,100,6)
#'
#' @export
excel_lognormdist <- function(x, mean, sd) {

  return(plnorm(x,mean, sd))
}

