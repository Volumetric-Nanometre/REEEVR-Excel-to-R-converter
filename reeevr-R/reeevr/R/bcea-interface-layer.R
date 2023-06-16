#' Form BCEA Dataframe
#'
#' Take N inputs and dumps them to screen
#' @param ... the N inputs
#' @return dataframe containing all inputs
#'
#' @examples
#' dataframe <- BCEA_dataframe(a=c(1,2,3),b=c(4,5,6))
#'
#' @export

BCEA_dataframe <- function(...) {

  df <- data.frame(...)

  return(df)
}
