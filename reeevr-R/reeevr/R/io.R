#' Screen Dump
#'
#' Take N inputs and dumps them to screen
#' @param ... the N inputs
#'
#' @examples
#' screen_dump(6, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
#'
#' @export

screen_dump <- function(...) {

  outputs <- list(...)

  for (output in outputs){

    print(output)
  }
}

#' Write PSA to output
#'
#' Take N inputs and dumps them to screen
#' @param file file path
#' @param dec  allows user to determine their own radix symbol (default = ".")
#' @param ... the N inputs
#'
#' @examples
#' PSA_output(c(1,0,1), c(1,2,3), c(3,4,5))
#'
#' @export
PSA_output <-function(file, dec = ".", ...){

  df <- data.frame(...)
  write.table(df, file = file, dec = dec, row.names = FALSE)

}
