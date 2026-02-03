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

#' Excel FLOOR Function
#'
#' Rounds a number down to the nearest multiple of significance.
#' This function replicates the behavior of Excel's FLOOR function.
#'
#' @param number The numeric value to round down
#' @param significance The multiple to which you want to round
#' @return The rounded number
#' @examples
#' excel_floor(3.7, 2)      # Returns 2
#' excel_floor(26.75, 0.1)  # Returns 26.7
#' excel_floor(-2.5, -2)    # Returns -2
#' excel_floor(1.5, 1)      # Returns 1
#'
excel_floor <- function(number, significance) {
  # Handle NA values
  if (is.na(number) || is.na(significance)) {
    return(NA)
  }

  # If significance is 0, return 0
  if (significance == 0) {
    return (0)
  }

  return (floor(number / significance) * significance)
}

#' Excel FLOOR.MATH Function
#'
#' Rounds a number down to the nearest integer or to the nearest multiple of significance.
#' Unlike FLOOR, FLOOR.MATH allows you to control the rounding direction for negative numbers.
#'
#' @param number The numeric value to round down
#' @param significance The multiple to which you want to round (default = 1)
#' @param mode For negative numbers, controls rounding direction:
#'             0 or omitted = round toward zero (toward positive infinity)
#'             non-zero = round away from zero (toward negative infinity)
#' @return The rounded number
#' @examples
#' excel__xlfn.floor.math(4.3)           # Returns 4
#' excel__xlfn.floor.math(4.3, 2)        # Returns 4
#' excel__xlfn.floor.math(-4.3)          # Returns -5
#' excel__xlfn.floor.math(-4.3, 1, 1)    # Returns -4
#' excel__xlfn.floor.math(-4.3, 2)       # Returns -6
#' excel__xlfn.floor.math(-4.3, 2, 1)    # Returns -4
#'
excel__xlfn.floor.math <- function(number, significance = 1, mode = 0) {
  # Handle NA values
  if (is.na(number))
    return(NA)

  if (is.na(significance))
    significance <- 1

  if (is.na(mode))
    mode <- 0

  # If significance is 0, return 0
  if (!significance)
    return(0)

  # Use absolute value of significance (unlike FLOOR)
  significance <- abs(significance)

  # Handle positive numbers - always round down
  if (number >= 0) {
    result <- floor(number / significance) * significance
  } else {
    # Handle negative numbers based on mode
    if (!mode) {
      # Mode 0: Round toward zero (up for negative numbers)
      # This means we use ceiling on the absolute value
      result <- -ceiling(abs(number) / significance) * significance
    }
    else {
      # Mode non-zero: Round away from zero (down for negative numbers)
      # This means we use floor on the absolute value
      result <- -floor(abs(number) / significance) * significance
    }
  }

  return(result)
}

#' Excel CEILING Function
#'
#' Rounds a number up to the nearest multiple of significance.
#' This function replicates the behavior of Excel's CEILING function.
#'
#' @param number The numeric value to round up
#' @param significance The multiple to which you want to round
#' @return The rounded number
#' @examples
#' excel_ceiling(2.5, 1)      # Returns 3
#' excel_ceiling(26.75, 0.1)  # Returns 26.8
#' excel_ceiling(-2.5, -2)    # Returns -4
#' excel_ceiling(1.2, 1)      # Returns 2
#'
excel_ceiling <- function(number, significance) {
  # Handle NA values
  if (is.na(number) || is.na(significance)) {
    return(NA)
  }

  # If significance is 0, return 0
  if (significance == 0) {
    return(0)
  }

  return(ceiling(number / significance) * significance)
}

#' Excel CEILING.MATH Function
#'
#' Rounds a number up to the nearest integer or to the nearest multiple of significance.
#' Unlike CEILING, CEILING.MATH allows you to control the rounding direction for negative numbers.
#'
#' @param number The numeric value to round up
#' @param significance The multiple to which you want to round (default = 1)
#' @param mode For negative numbers, controls rounding direction:
#'             0 or omitted = round toward zero (away from infinity)
#'             non-zero = round away from zero (toward negative infinity)
#' @return The rounded number
#' @examples
#' excel__xlfn.ceiling.math(4.3)           # Returns 5
#' excel__xlfn.ceiling.math(4.3, 2)        # Returns 6
#' excel__xlfn.ceiling.math(-4.3)          # Returns -4
#' excel__xlfn.ceiling.math(-4.3, 1, 1)    # Returns -5
#' excel__xlfn.ceiling.math(-4.3, 2)       # Returns -4
#' excel__xlfn.ceiling.math(-4.3, 2, 1)    # Returns -6
#'
excel__xlfn.ceiling.math <- function(number, significance = 1, mode = 0) {
  # Handle NA values
  if (is.na(number))
    return(NA)

  if (is.na(significance))
    significance <- 1

  if (is.na(mode))
    mode <- 0

  # If significance is 0, return 0
  if (!significance)
    return(0)

  # Use absolute value of significance (unlike CEILING)
  significance <- abs(significance)

  # Handle positive numbers - always round up
  if (number >= 0) {
    result <- ceiling(number / significance) * significance
  }
  else {
    # Handle negative numbers based on mode
    if (!mode) {
      # Mode 0: Round toward zero (up for negative numbers)
      # This means we use floor on the absolute value
      result <- -floor(abs(number) / significance) * significance
    }
    else {
      # Mode non-zero: Round away from zero (down for negative numbers)
      # This means we use ceiling on the absolute value
      result <- -ceiling(abs(number) / significance) * significance
    }
  }

  return(result)
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
