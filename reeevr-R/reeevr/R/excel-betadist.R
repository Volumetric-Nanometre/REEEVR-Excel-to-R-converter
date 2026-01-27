#' Excel BETA.DIST function
#'
#' return correct norm distribution
#' @param x value to evaluate
#' @param alpha
#' @param beta
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @param A Optional parameter - lower bound
#' @param B Optional parameter - upper bound

#' @return qbinom or dbinom
#'
#' @examples
#' beta.dist <- excel__xlfn.beta.dist(0.2,100,6,True)
#'
#' @export
excel__xlfn.beta.dist <- function(x, alpha, beta, cumulative, A=0, B=1) {

  x01 <- (x - A) / (B - A)

  if(cumulative){
    return(pbeta(x01, shape1 = alpha, shape2 = beta))

  }
  else{
    return(dbeta(x01, shape1 = alpha, shape2 = beta)/(B-A))
  }

}

#' Excel BETADIST function
#'
#' return correct norm distribution
#' @param x value to evaluate
#' @param alpha
#' @param beta
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @param A Optional parameter - lower bound
#' @param B Optional parameter - upper bound

#' @return qbinom or dbinom
#'
#' @examples
#' betadist <- excel_betadist(0.2,100,6,True)
#'
#' @export
excel_betadist <- function(x, alpha, beta, A=0, B=1) {


    x01 <- (x - A) / (B - A)

    return(pbeta(x01, shape1 = alpha, shape2 = beta))

}

