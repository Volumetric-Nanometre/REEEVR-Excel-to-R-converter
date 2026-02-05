library(BCEA)
library(ggplot2)
library(purrr)

#' Form BCEA matrix
#'
#' Take N inputs and dumps them to screen
#' @param ... the N inputs
#' @return dataframe containing all inputs
#'
#' @examples
#' costs <- BCEA_matrix(a=c(1,2,3),b=c(4,5,6))
#' effectiveness <- BCEA_matrix(a=c(1,2,3),b=c(4,5,6))
#'
#' @export

BCEA_matrix <- function(...) {

  m <- cbind(...)

  return(m)
}


#' Call BCEA
#'
#' Take N inputs and dumps them to screen
#' @param ... the N inputs
#' @return dataframe containing all inputs
#'
#' @examples
#' BCEA_all_output_loop(costs,effectiveness,treatments,Kmax)
#'
#' @export

BCEA_all_output_loop <- function(costs,effectiveness,treatments,Kmax) {

  totalRef <- length(treatments)


  for(ref in 1:totalRef){

    outfolder <- treatments[ref]
    if(!dir.exists(outfolder))
      dir.create(outfolder)

    bcea_output <- bcea(effectiveness,costs, ref = ref,interventions = treatments, Kmax = Kmax)
    png(file.path(outfolder,"ce-plane-plot.png"))
    ceplane.plot(bcea_output)
    dev.off()
    png(file.path(outfolder,"multi-plot.png"))
    multice = multi.ce(bcea_output)
    ceac.plot(multice)
    dev.off()

  }

}




