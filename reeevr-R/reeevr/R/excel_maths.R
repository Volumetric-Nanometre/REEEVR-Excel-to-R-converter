#' Select which SUM should be used
#'
#' Select the correct SUM type based on inputs.
#' If the inputs are a singular list, then call sum()
#' If inputs are a list of lists, call Reduce('+',list())
#' @param ... list of inputs
#' @return either sum(c(...)) or Reduce('+',list(...))
#'
#' @examples
#' sumvals <- excel_sum(list(1,0,1))
#' sumvals <- excel_sum(list(c(1,0,1),c(1,0,1),c(1,0,1)))
#' @export
excel_sum <- function(arrayinput) {

  if(length(dim(arrayinput)) < 1)
    return (Reduce('+',arrayinput))

  return (colSums(arrayinput))
}

#' Convert POWER(a,b) to a x b x c x ...
#'
#' @param a
#' @param b
#' @return
#'
#' @examples
#' powerSingleVars <- excel_power(2, 3)
#' powerListVars <- excel_product(c(2, 3, 5), c(2.5 ,-10, 2))
#' @export
excel_power <- function(a, b) {

  temp = a^b

  return (temp)
}


#' Convert PRODUCT(...) to a x b x c x ...
#'
#' @param ... list
#' @return
#'
#' @examples
#' productSingleVars <- excel_product(2, 3, 5, 2.5, -10)
#' productListVars <- excel_product(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9))
#' @export
excel_product <- function(...) {

  vars = list(...)
  temp = vars[1]/vars[1] # Allows us to create either a single value or list of the correct length
  #  and sets all values to 1 (obviously small error introduction due to float division)
  for(var in vars){
    temp <- temp * var
  }

  return (temp)
}

#' Convert FLOOR(x, n)
#' See https://stackoverflow.com/questions/47177246/floor-and-ceiling-with-2-or-more-significant-digits
#'
#' @param x list or single value
#' @param n number of significant figures - integer
#' @return
#'
#' @examples
#' floorSingleVarOne <- excel_floor(2.3456, 3)
#' floorSingleVarTwo <- excel_floor(2345.6, 2)
#' floorSingleVarOneNeg <- excel_floor(-2.3456, 3)
#' floorSingleVarTwoNeg <- excel_floor(-2345.6, 2)
#'
#' floorListleVars <- excel_floor(c(2.3456, 2345, -2.345,-2093.987,-223,0), 3)
#' @export
excel_floor <- function(x, n) {

  pow <- floor( log10( abs(x) ) ) + 1 - n
  y <- floor(x / 10 ^ pow) * 10^pow
  # handle the x = 0 case
  y[x==0] <- 0
  return (y)
}

#' Convert CEILING(x, n)
#' See https://stackoverflow.com/questions/47177246/floor-and-ceiling-with-2-or-more-significant-digits
#'
#' @param x list or single value
#' @param n number of significant figures - integer
#' @return
#'
#' @examples
#' ceilingSingleVarOne <- excel_ceiling(2.3456, 3)
#' ceilingSingleVarTwo <- excel_ceiling(2345.6, 2)
#' ceilingSingleVarOneNeg <- excel_ceiling(-2.3456, 3)
#' ceilingSingleVarTwoNeg <- excel_ceiling(-2345.6, 2)
#'
#' ceilingListleVars <- excel_ceiling(c(2.3456, 2345, -2.345,-2093.987,-223,0), 3)
#' @export
excel_ceiling <- function(x, n) {

  pow <- floor( log10( abs(x) ) ) + 1 - n
  y <- ceiling(x / 10 ^ pow) * 10^pow
  # handle the x = 0 case
  y[x==0] <- 0
  return (y)
}

#' Convert MMULT(A, B)
#' Performs AB = X and X are all matrices.
#'
#' @param A Matrix of size n x m
#' @param B MAtrix of size m x l
#' @return
#'
#' @examples
#'
#' @export
excel_mmult <- function(A, B) {

  return (A %*% B)

}

#' Convert LOG(x, base)
#'
#' @param x value
#' @param base base of logarithm
#' @return
#'
#' @examples
#' log1 <- excel_log(10)
#' log2 <- excel_log(16,2)
#'
#' @export
excel_log <- function(x, base=10) {

  return(log(x,base))
}



