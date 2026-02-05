#' Excel RAND function
#'
#' Take number of random numbers, and the validation flag
#' @param numberofruns integer number of PSA runs
#' @param validate  TRUE/FALSE
#' @return vector of random numbers, or a vector of validation samples
#'
#' @examples
#' test <- excel_rand(numberofruns,validate)
#'
#' @export
#'
excel_rand <- function(numberofruns,validate) {

  # If validate, provide array of static values that can be input into the EXCEL sheet.
  # Values must exist as [0-1] as this is the output of runif
  # Range of values are provided to provide validation.
  if(validate){
    return(c(0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9))
  }
  else{
    return(runif(numberofruns))
  }

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


#' Construct the correct form of array
#'
#' @param ... input vars
#' @return array
#'
#' @examples
#'
#' @export
excel_array <- function(dim,...) {
  rows = dim[0]
  columns = dim[1]
  vars = list(...)
  for(var in vars){
    print(dim(var))
  }
}
