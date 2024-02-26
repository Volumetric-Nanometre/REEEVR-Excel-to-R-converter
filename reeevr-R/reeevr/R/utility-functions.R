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


#' array_construct function
#'
#' Correctly interprets and constructs arrays
#' @param ... vars to form the array
#' @param dims dimensions of the array we are constructing
#' @return the constructed array
#'
#' @examples
#' array_construct(array(c(40,40,40,40,20,20),c(2,3)),array(c(20,20,20,20,20,20),c(2,3)))
#' @export
#'
array_construct <- function(...){

  arraycomponents = lapply(list(...),dim)
  print(arraycomponents)

  if(any(unlist(Map(is.null, arraycomponents)))){

    for(index in 1:length(arraycomponents)){
      if(is.null(arraycomponents[[index]]))
        arraycomponents[[index]] = length(list(...)[[index]])
    }

  }
  print(arraycomponents)


  if(length(unique(arraycomponents)) == 1)
    return(array_equal(...))

  else
    return(array_unequal(...))

}

array_equal <- function(...){
  print("Equal")

  equalarrays = list(...)
  innerdims = dim(equalarrays[[1]])
  numofarrays = length(equalarrays)
  print(numofarrays)
  newdim =c(innerdims,numofarrays)
  print(newdim)
  return(array(c(...),dim = newdim))
}

array_unequal <- function(...){
  print("Unequal")
  return()
}
