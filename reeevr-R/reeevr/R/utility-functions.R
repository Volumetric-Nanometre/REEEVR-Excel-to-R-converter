#' is.blank function
#'
#' Check if value is a "blank" value
#' @param x the value to check
#' @param false.triggers values to trigger when false
#' @return true/false
#'
#' @examples
#' sixth <- is.blank(6)
#' tenth <- is.blank(NA)
#'
#' @export
#'
is.blank <- function(x, false.triggers=FALSE){
  if(is.function(x)) return(FALSE) # Some of the tests below trigger
  # warnings when used on functions
  return(
    is.null(x) ||                # Actually this line is unnecessary since
      length(x) == 0 ||            # length(NULL) = 0, but I like to be clear
      all(is.na(x)) ||
      all(x=="") ||
      (false.triggers && all(!x))
  )
}
