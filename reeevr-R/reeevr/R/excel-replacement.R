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
