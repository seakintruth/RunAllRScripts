if (!require(pacman)) {install.packages("pacman")}
pacman::p_load(dplyr, knitr, tcltk, fs)

# Ths script should be called from either:
# Rstudio, sourced in Rgui.exe
# or passed as an argument to Rscript.exe on windows,
# or the results of 'which Rscript' in bash on linux...

# For this to work each sourced script shall not perform a call to the
function setwd()
# this script is doing that...

# ------------- Set working directory -------------
# getCurrentFileLocation only works when being called as a script to
Rscript.exe or in RStudio
# not usind here::here as it doesn't tell us where this script actually is,
# it might tell us where this project is...


#' @export
getCurrentFileLocation <- function() {
  if (requireNamespace("knitr")) {
    # get absolute path using dir = FALSE
    this_file <- knitr::current_input(dir = FALSE)
  }
  if(is.null(this_file)) {
    this_file <- try(
      rstudioapi::getSourceEditorContext()$path,
      silent=TRUE
    )
    if(class(this_file) == "try-error") {
      this_file <- try(
        commandArgs() %>%
          tibble::enframe(name = NULL) %>%
          tidyr::separate(col=value, into=c("key", "value"), sep="=",
fill='right') %>%
          dplyr::filter(key == "--file") %>%
          dplyr::pull(value),
          silent=TRUE
      )
      if(
                (class(this_file) == "try-error") || (length(this_file)==0)
        ){
        this_file <- NULL
      }
    }
  }
  return(this_file)
}       
pathCurrentRunAllFile <- getCurrentFileLocation()
if (is.null(pathCurrentRunAllFile)){
        pathCurrentRunAllFile <- tcltk::tk_choose.files(
                "RunAll.R",
                caption="Select the RunAll.R file for this project",
                multi=FALSE,
                filters=matrix(c("Only the RunAll.R file","RunAll.R"),ncol=2)
        )
}
newWorkingDirectory <- try(setwd(dirname(pathCurrentRunAllFile)))
if(class(newWorkingDirectory) == "try-error") {
  warning("  Fatal ERROR - RunAll.R: \n", newWorkingDirectory[1])
} else {
  # read ini values from (assets-RunAll/RunAll.ini)
  iniRunAll <- read.ini("assets-RunAll/RunAll.ini")
  if ((iniRunAll$Config$ScriptsDirectoryUseRelativePath ==  "TRUE")){
        iniRunAll$Config$ScriptsDirectoryName <-
fs::path_join(c(getwd(),iniRunAll$Config$ScriptsDirectoryName))
  }
  if ((iniRunAll$Config$ReportsDirectoryUseRelativePath ==  "TRUE")){
        iniRunAll$Config$ReportsDirectoryName <-
fs::path_join(c(getwd(),iniRunAll$Config$ReportsDirectoryName))
  }
  if ((iniRunAll$Log$LocationUseRelativePath ==  "TRUE")){
        iniRunAll$Log$Location <- fs::path_join(c(getwd(),iniRunAll$Log$Location))
  }

  message("Note: Working directory for ",
iniRunAll$Config$ProjectName, " is now: \n\t'",getwd(),"'")

  # ------------- Source local packages and functions -------------
  source("assets-RunAll/logSysInfo.R")
  source("assets-RunAll/ini.R")
  # ------------------------------------------------------------
  # Generate our list of scripts to run from the values in the runAll.ini file
  scriptsToRun <- unlist(iniRunAll$ScriptFiles)
  if ((iniRunAll$Config$ScriptFilesUseRelativePath == "TRUE")){

  }
  # ------------------------------------------------------------
  # Generate our list of scripts to run from the values in the 'scripts' folder


  # ------------- Run all production ready scripts (ordered) -------------

  # [TODO] read the log path from the RunAll.ini configuration file.
  logSysInfo()
}
