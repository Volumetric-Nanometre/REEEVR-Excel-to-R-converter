#' Excel NORM.DIST function
#'
#' return correct norm distribution
#' @param x list of random numbers
#' @param mean mean of distribution
#' @param sd standard deviation of distribution
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return qnorm or dnorm
#'
#' @examples
#' norm.dist <- excel__xlfn.norm.dist(0.2,100,6,True)
#'
#' @export
excel__xlfn.norm.dist <- function(x, mean, sd,cumulative) {

  if(cumulative){
    return(pnorm(x,mean,sd))

  }
  else{
    return(dnorm(x,mean,sd))
  }

}

#' Excel NORMDIST function
#'
#' return correct norm distribution
#' @param x list of random numbers
#' @param mean mean of distribution
#' @param sd standard deviation of distribution
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return qnorm or dnorm
#'
#' @examples
#' normdist <- excel_normdist(0.2,100,6,True)
#'
#' @export
excel_normdist <- function(x, mean, sd,cumulative) {

  if(cumulative){
    return(pnorm(x,mean,sd))

  }
  else{
    return(dnorm(x,mean,sd))
  }

}

#' Excel NORM.S.DIST function
#'
#' return correct norm distribution
#' @param x distribution value
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return qnorm or dnorm
#'
#' @examples
#' norm.s.dist <- excel__xlfn.norm.s.dist(0.2,100,6,True)
#'
#' @export
excel__xlfn.norm.s.dist <- function(x,cumulative) {

  if(cumulative){
    return(pnorm(x,mean=0,sd=1))

  }
  else{
    return(dnorm(x,mean=0,sd=1))
  }

}

#' Excel NORMSDIST function
#'
#' return correct norm distribution
#' @param x distribution value
#' @return qnorm
#'
#' @examples
#' normsdist <- excel_normsdist(0.2)
#'
#' @export
excel_normsdist <- function(x) {

    return(pnorm(x,mean=0,sd=1))
}
