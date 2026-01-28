#' Excel BINOM.DIST function
#'
#' return correct norm distribution
#' @param x Number of successful trials
#' @param N TOtal number of trials
#' @param prob Probability
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return pbinom or dbinom
#'
#' @examples
#' binom.dist <- excel_xlfn.binom.dist(0.2,100,6,True)
#'
#' @export
excel__xlfn.binom.dist <- function(x, N, prob, cumulative) {

  if(cumulative){
    return(pbinom(x,size=N, prob=prob))

  }
  else{
    return(dbinom(x,size=N, prob=prob))
  }

}

#' Excel BINOMDIST function
#'
#' return correct norm distribution
#' @param x Number of successful trials
#' @param N TOtal number of trials
#' @param prob Probability
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return pbinom or dbinom
#'
#' @examples
#' binomdist <- excel_binomdist(0.2,100,6,True)
#'
#' @export
excel_binomdist <- function(x, N, prob, cumulative) {

  if(cumulative){
    return(pbinom(x,size=N, prob=prob))

  }
  else{
    return(dbinom(x,size=N, prob=prob))
  }

}

#' Excel NEGBINOM.DIST function
#'
#' return correct norm distribution
#' @param num_f Number of failed trials
#' @param num_s Number of successful trials
#' @param prob Probability
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return pnbinom or dnbinom
#'
#' @examples
#' negbinom.dist <- excel__xlfn.negbinom.dist(3,4,0.25,True)
#'
#' @export
excel__xlfn.negbinom.dist <- function(num_f, num_s, prob, cumulative) {

  if(cumulative){
    return(pnbinom(num_f,size=num_s, prob=prob))

  }
  else{
    return(dnbinom(num_f,size=num_s, prob=prob))
  }

}

#' Excel NEGBINOMDIST function
#'
#' return correct norm distribution
#' @param num_f Number of failed trials
#' @param num_s Number of successful trials
#' @param prob Probability
#' @return dnbinom
#'
#' @examples
#' negbinomdist <- excel_negbinomdist(3,4,0.25)
#'
#' @export
excel_negbinomdist <- function(num_f, num_s, prob) {

    return(dnbinom(num_f,size=num_s, prob=prob))

}

#' Excel BINOM.INV function
#'
#' return correct norm distribution
#' @param N Total number of trials
#' @param prob Probability of success
#' @param alpha citerion value
#' @return qbinom or dbinom
#'
#' @examples
#' binomdist <- excel__xlfn.binom.inv(100,0.2,6)
#'
#' @export
excel__xlfn.binom.inv <- function(N, prob, alpha) {

    return(qbinom(p=alpha, size=N, prob=prob ))

}
