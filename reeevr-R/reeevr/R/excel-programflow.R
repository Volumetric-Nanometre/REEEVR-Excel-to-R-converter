#' Excel INDEX function
#'
#' Take three inputs and return the value at location
#' @param inputarray array to be accessed
#' @param row  row of data
#' @param column column of data
#' @return data located at (row,column)
#'
#' @examples
#' test <- excel_index(array(c(1,2,3,4),dim=c(2,2)),1,1)
#'
#' @export
#'
excel_index <- function(inputarray,row, column) {
  return(inputarray[row,column])
}

#' Excel CHOOSE function
#'
#' Take N inputs and return the nth result.
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
