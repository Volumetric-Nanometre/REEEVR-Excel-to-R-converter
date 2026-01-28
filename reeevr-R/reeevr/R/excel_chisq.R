#' Excel CHIDIST function
#'
#' return correct norm distribution
#' @param x values
#' @param df degrees of freedom
#' @return pchisq
#'
#' @examples
#' chidist <- excel_chidist(2,100)
#'
#' @export
excel_chidist <- function(x, df) {

    return(pchisq(x,df,lower.tail = FALSE))

}

#' Excel CHIINV function
#'
#' return correct norm distribution
#' @param x probability
#' @param df degrees of freedom
#' @return
#'
#' @examples
#' chidist <- excel_chidist(2,100)
#'
#' @export
excel_chiinv <- function(x, df) {

  return(qchisq(x,df,lower.tail = FALSE))

}

#' Excel CHISQ.DIST function
#'
#' return correct norm distribution
#' @param x values
#' @param df degrees of freedom
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return pchisq or dchisq
#'
#' @examples
#' chisq.dist <- excel__xlfn.chisq.dist(2,100,True)
#'
#' @export
excel__xlfn.chisq.dist <- function(x, df, cumulative) {

  if(cumulative){
    return(pchisq(x,df))

  }
  else{
    return(dchisq(x,df))
  }

}

#' Excel CHISQ.DIST.RT function
#'
#' return correct norm distribution
#' @param x values
#' @param df degrees of freedom
#' @param cumulative TRUE/FALSE flag switching between cumulative distribution function, and probability mass function.
#' @return dchisq
#'
#' @examples
#' chisq.dist.rt <- excel__xlfn.chisq.dist.rt(2,100)
#'
#' @export
excel__xlfn.chisq.dist.rt <- function(x, df) {

  return(pchisq(x,df,lower.tail = FALSE))

}

#' Excel CHISQ.INV function
#'
#' return correct norm distribution
#' @param x val
#' @param df degrees of freedom
#' @return qchisq
#'
#' @examples
#' chisq.inv <- excel__xlfn.chisq.inv(2,100)
#'
#' @export
excel__xlfn.chisq.inv<- function(x, df) {

  return(qchisq(x,df))

}

#' Excel CHISQ.INV.RT function
#'
#' return correct norm distribution
#' @param x val
#' @param df degrees of freedom
#' @return qchisq
#'
#' @examples
#' chisq.inv.rt <- excel__xlfn.chisq.inv.rt(2,100)
#'
#' @export
excel__xlfn.chisq.inv.rt<- function(x, df) {

  return(qchisq(x,df,lower.tail=FALSE))

}

#' Excel CHISQ.TEST function
#'
#' return correct norm distribution
#' @param x actual range
#' @param y expected range
#' @return chisq.test
#'
#' @examples
#' chisq.test<- excel__xlfn.chisq.test(2,100)
#'
#' @export
excel__xlfn.chisq.test<- function(x, y) {

  return(chisq.test(x,y))

}

#' Excel CHITEST function
#'
#' return correct norm distribution
#' @param x actual range
#' @param y expected range
#' @return chisq.test
#'
#' @examples
#' chitest<- excel__xlfn.chisq.inv.rt(2,100)
#'
#' @export
excel_chitest<- function(x, y) {

  return(chisq.test(x,y))

}
