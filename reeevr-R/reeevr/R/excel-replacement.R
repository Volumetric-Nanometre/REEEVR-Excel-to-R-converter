#' Excel CHOOSE function
#'
#' Take N inputs and returnt he nth result.
#' @param n the nth input in the list of N inputs
#' @param ... the N inputs. N<n is forbidden
#' @return The nth input
#'
#' @examples
#' sixth <- excel_choose(6, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
#' tenth <- excel_choose(10, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
#'
#' @export

excel_choose <- function(n, ...) {
  inputs <- list(...)
  return(inputs[[n]])
}


#' Ignored an Excel function
#'
#' Take N inputs and return NA.
#' Inform the user that the expected function is not
#' implimented.
#' @param ... the N inputs.
#' @return NA
#'
#' @examples
#' ignored_function <- excel_ignore(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
#'
#' @export

excel_ignore <- function(...) {
  cat(paste("Excel function has been ignored. An NA value has been returned.\n",
    "If this is detrimental to your work, please contact the maintainer about\n",
    "about implimenting the function that has been ignored.\n",
    sep = ""
  ))

  return(NA)
}

#' Excel IFERROR function
#'
#' Check if value is blank and return non-blank value.
#' @param value expected non-blank value
#' @param value_if_error value to return if blank
#' @return return non-blank value
#'
#' @examples
#' num <- excel_iferror(6)
#' blank <- excel_iferror(NA)
#'
#' @export

excel_iferror <- function(value, value_if_error) {

  returnVar = value
  if(is.blank(value))
    returnVar = value_if_error
  return(returnVar)
}

#' Excel GAMMAINV function
#'
#' return corret qgamma distribution
#' @param x list of random numbers
#' @param alpha
#' @param beta
#' @return qgamma output
#'
#' @examples
#' gaminv <- excel_gammainv(0.2,100,6)
#'
#' @export
excel_gammainv <- function(x, alpha, beta) {

  return(qgamma(x,shape=alpha,scale=beta))
}

#' Select which SUM should be used
#'
#' Select the correct SUM type based on inputs.
#' If the inputs are a singular list, then call sum()
#' If inputs are a list of lists, call Reduce('+',list())
#' @param ... list of inputs
#' @return either sum(c(...)) or Reduce('+',list(...))
#'
#' @examples
#' sumvals <- excel_sum_select(list(1,0,1))
#' sumvals <- excel_sum_select(list(c(1,0,1),c(1,0,1),c(1,0,1)))
#' @export
excel_sum_select <- function(...) {
  vars = list(...)
  for(val in vars)
    if(length(val)>1)
      return (Reduce('+',...))

  return(sum(...))
}

