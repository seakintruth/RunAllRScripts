# only base packages requried...
logSysInfo <- function(logFilePath = "assets-RunAll/Log/sessionInfo.csv", UserMessage="") {
  sessionNames <- c("platform","running","R.version")
  sessionInfoValues <- suppressWarnings(sessionInfo())
  projectName <-        if (exists("iniRunAll")) {
          iniRunAll$Config$ProjectName
  } else {
          "Project Not Specified"
  }
  sessionInfoValues <- matrix(
    c(
        projectName,
      sessionInfoValues$R.version$version.string,
      sessionInfoValues$platform,
      sessionInfoValues$running,
      Sys.time(),
        paste(UserMessage,sep = "; ")
    ),ncol=6
  )
  colnames(sessionInfoValues) <- c("project",sessionNames,"TimeStamp","UserMessage")
  if (file.exists(logFilePath)) {
    write.table(
      sessionInfoValues,
      logFilePath,
      append = TRUE,
      row.names = FALSE,
      col.names = FALSE,
      sep = ","
    )
  } else {
    if (!dir.exists(dirname(logFilePath))) {
      # Build log file path
      dir.create(dirname(logFilePath))
    }
    write.csv(
      sessionInfoValues,
      logFilePath,
      row.names = FALSE
    )
  }
}
