# Ths script should allways be called from Rstudio, 
# or passed as an argument to Rscript.exe on windows, 
# or the results of 'which Rscript' on linux...
if (!require(pacman)) {install.packages("pacman")}
pacman::p_load(dplyr, knitr, ini) 
# For this to work each sourced script shall not perform a call to the function setwd() 
# this script is doing that...

# ------------- Set working directory -------------
# getCurrentFileLocation only works when being called as a script to Rscript.exe or in RStudio
# not usind here::here as it doesn't tell us where this script actually is, 
# it might tell us where this project is...


#' @export
getCurrentFileLocation <- function() {
  if (requireNamespace("knitr")) {
    # get absolute path using dir = FALSE
    this_file <- knitr::current_input(dir = FALSE) 
  }
  if(is.null(this_file)) {
    this_file <- try(rstudioapi::getSourceEditorContext()$path)
    if(class(this_file) == "try-error") {
      this_file <- try(
        commandArgs() %>%
          tibble::enframe(name = NULL) %>%
          tidyr::separate(col=value, into=c("key", "value"), sep="=", fill='right') %>%
          dplyr::filter(key == "--file") %>%
          dplyr::pull(value)
      )
      if(class(this_file) == "try-error") {
        this_file <- NULL
      }
    }    
  }
  return(this_file)
}	

getCurrentFileLocation()
newWorkingDirectory <- try(setwd(dirname(getCurrentFileLocation())))
if(class(newWorkingDirectory) == "try-error") {
  warning("  Fatal ERROR - RunAll.R: \n", newWorkingDirectory[1])
} else {
  message("Note: Working directory is now: \n\t'",getwd(),"'")
  
  # read ini values from (assets-RunAll/RunAll.ini)



  # ------------------------------------------------------------
  # ------------- Run all production ready scripts -------------
  
  
  # ------------------------------------------------------------
  # Read our list of scripts to run from the values in the runAll.ini file
  
  
  # ------------------------------------------------------------
  # If a ./scripts directory exists run each .R script found that is named 001 through 999 -------------
  
  source("assets-RunAll/logSysInfo.R")
  # [TODO] read the log path from the RunAll.ini configuration file.
  logSysInfo()  
}
