#' Select which SUM should be used
#'
#' Select the correct SUM type based on inputs.
#' If the inputs are a singular list, then call sum()
#' If inputs are a list of lists, call Reduce('+',list())
#' @param ... list of inputs
#' @return either sum(c(...)) or Reduce('+',list(...))
#'
#' @examples
#' sumSingleVars <- excel_sum(0, 3, 5, 2.5, -10)
#' sumListVars <- excel_sum(list(2, 3, 5, 2.5, -10))
#' sumMultiListVars <- excel_sum(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#' @export
excel_sum <- function(...) {

  args = list(...)
  if(length(args) == 1)
    return (apply(do.call(rbind, ...), 2, sum))

  else
    return (excel_sum(args))
}

#' Convert POWER(a,b)
#'
#' @param a value
#' @param b exponent
#' @return a^b
#'
#' @examples
#' powerSingleVars <- excel_power(2, 3)
#' @export
excel_power <- function(a, b) {

  temp = a^b

  return (temp)
}


#' Convert PRODUCT
#'
#' @param ... list of inputs
#' @return either single value product, or vector of products
#'
#' @examples
#' productSingleVars <- excel_product(0, 3, 5, 2.5, -10)
#' productListVars <- excel_product(list(2, 3, 5, 2.5, -10))
#' productMultiListVars <- excel_product(list(c(2, 3, 5), c(2.5 ,-10, 2), c(4, 3.3, 9)))
#' @export
excel_product <- function(...) {

  args = list(...)
  if(length(args) == 1)
    return (apply(do.call(rbind, ...), 2, prod))

  else
    return (excel_product(args))
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
#' @return log(x,b)
#'
#' @examples
#' log1 <- excel_log(10)
#' log2 <- excel_log(16,2)
#'
#' @export
excel_log <- function(x, base=10) {

  return(log(x,base))
}

#' Convert COUNT
#'
#' Count the number of numerical values in  the selected cells
#'
#' @param ... list of inputs
#' @return number of cells that are numeric
#'
#' @examples
#' countSingleVars <- excel_count(0, 3, 5, 'A', -10)
#' countListVars <- excel_count(list(2, 'V', 5, 2.5, -10))
#' @export
excel_count <- function(...) {
  args = list(...)
  count = 0

  for (val in args) {
    if (is.list(val)) {
      count = count + sum(sapply(val, is.numeric))
    }
    else if (is.numeric(val)) {
      count = count + 1
    }
  }

  return(count)
}

#' Convert COUNTA
#'
#' Count the number of numerical values in  the selected cells
#'
#' @param ... list of inputs
#' @return number of cells that are not NA
#'
#' @examples
#' countaSingleVars <- excel_counta(0, 3, 5, 'A', -10)
#' countaListVars <- excel_counta(list(2, 'V', 5, 2.5, -10))
#' @export
excel_counta <- function(...) {
  args = list(...)
  count = 0

  for (val in args) {
    if (is.list(val)) {
      count = count + sum(sapply(val, function(x) !is.na(x)))
    }
    else if (!is.na(val)) {
      count = count + 1
    }
  }

  return(count)
}

#' Convert COUNTIF
#'
#' Count the number of numerical values in  the selected cells that satisfy the criterion
#'
#' @param range range of which we will compare
#' @return number of cells that satisfy the criteria
#'
#' @examples
#' excel_countif(list(2, 'V', 5, 2.5, -10), "V")  # Returns: 1
#' excel_countif(list(2, 'V', 5, 2.5, -10), 2)    # Returns: 1
#' excel_countif(list(2, 'V', 5, 2.5, -10), ">2") # Returns: 2
#' excel_countif(list(2, 'V', 5, 2.5, -10), "<=2.5") # Returns: 3
#' excel_countif(list(2, 3, 5, "Test", "val", NA), "<>3") # Returns: 5
#' excel_countif(list(2, 3, 5, "Test", "val", NA), "=Test") # Returns: 1
#' @export
excel_countif <- function(range, criteria) {
  # Check if criteria is a comparison operator string
  if (grepl("^(>=|<=|<>|>|<|=)", criteria)) {
    # Extract operator and value
    operator <- sub("^(>=|<=|<>|>|<|=).*", "\\1", criteria)
    value_str <- sub("^(>=|<=|<>|>|<|=)", "", criteria)

    # Try to convert to numeric
    value <- suppressWarnings(as.numeric(value_str))

    # If value is numeric, do numeric comparison
    if (!is.na(value)) {
      numeric_range <- suppressWarnings(as.numeric(range))

      result <- switch(operator,
                       ">" = numeric_range > value,
                       "<" = numeric_range < value,
                       ">=" = numeric_range >= value,
                       "<=" = numeric_range <= value,
                       "<>" = numeric_range != value | is.na(numeric_range),
                       "=" = numeric_range == value)
    } else {
      # Text comparison
      result <- switch(operator,
                       "<>" = range != value_str,
                       "=" = range == value_str)
    }

    return(sum(result, na.rm = TRUE))
  } else {
    # Direct equality match
    return(sum(suppressWarnings(range == criteria), na.rm = TRUE))
  }
}
