#' Convert MID(str, n)
#'
#' @param str text string
#' @param x starting location of first char
#' @param n number of characters to be captured
#' @return
#'
#' @examples
#'  midstringOne<- excel_mid("string",1,1)
#'  midstringTwo<- excel_mid("string",1,5)
#'  midstringThree<- excel_mid("string",1,6)
#'  midstringFour<- excel_mid("string",3,1)
#'
#' @export
excel_mid <- function(str, x, n) {

  return (substr(str, x, x + n - 1))
}

#' Convert LEFT(str, n)
#'
#' @param str text string
#' @param n number of characters to be captured
#' @return
#'
#' @examples
#'  leftstringOne<- excel_left("string",1)
#'  leftstringTwo<- excel_left("string",5)
#'  leftstringThree<- excel_left("string",3)
#'
#' @export
excel_left <- function(str, n) {

  return (substr(str, 1, n))
}


#' Convert RIGHT(str, n)
#'
#' @param str text string
#' @param n number of characters to be captured
#' @return
#'
#' @examples
#'  rightstringOne<- excel_right("string",1)
#'  rightstringTwo<- excel_right("string",5)
#'  rightstringThree<- excel_right("string",3)
#'
#' @export
excel_right <- function(str, n) {

  return (substr(str, nchar(str) - n - 1, nchar(str)))
}


#' Convert CONCATENATE
#'
#' @param ... up to 256 indiviual cells to be concatenated
#' @return
#'
#' @examples
#'
#'
#' @export
excel_concatenate <- function(...) {

  return (paste(...,sep="",collapse = NULL))
}

#' Convert CONCAT
#'
#' @param ... up to 256 indiviual cells to be concatenated
#' @return
#'
#' @examples
#'
#'
#' @export
excel__xlfn.concat <- function(...) {

  return (paste(...,sep="",collapse = NULL))
}
