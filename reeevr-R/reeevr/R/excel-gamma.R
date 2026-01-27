#' Excel GAMMA.DIST function
#'
#' @param x list of random numbers
#' @param alpha shape in R
#' @param beta scale in R
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return qgamma or dgamma
#'
#' @examples
#' gamma.dist <- excel__xlfn.gamma.dist(0.2,100,6,True)
#'
#' @export
excel__xlfn.gamma.dist <- function(x, alpha, beta,cumulative) {

  if(cumulative){
    return(pgamma(x,shape=alpha, scale=beta))

  }
  else{
    return(dgamma(x,shape=alpha, scale=beta))
  }

}

#' Excel GAMMADIST function
#'
#' @param x list of random numbers
#' @param alpha shape in R
#' @param beta scale in R
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return qgamma or dgamma
#'
#' @examples
#' gammadist <- excel_gammadist(0.2,100,6,True)
#'
#' @export
excel_gammadist <- function(x, alpha, beta,cumulative) {

  if(cumulative){
    return(pgamma(x,shape=alpha, scale=beta))

  }
  else{
    return(dgamma(x,shape=alpha, scale=beta))
  }

}

#' Excel GAMMA.INV function
#'
#' @param x list of random numbers
#' @param alpha shape in R
#' @param beta scale in R
#' @return qgamma
#'
#' @examples
#' gamma.inv <- excel__xlfn.gamma.inv(0.2,100,6)
#'
#' @export
excel__xlfn.gamma.inv <- function(x, alpha, beta) {

    return(qgamma(x,shape=alpha, scale=beta))

}

#' Excel GAMMAINV function
#'
#' @param x list of random numbers
#' @param alpha shape in R
#' @param beta scale in R
#' @return qgamma
#'
#' @examples
#' gammainv <- excel_gammainv(0.2,100,6)
#'
#' @export
excel_gammainv <- function(x, alpha, beta) {

  return(qgamma(x,shape=alpha, scale=beta))

}


