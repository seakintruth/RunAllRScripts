# only base packages requried...
logSysInfo <- function(logFilePath = "assets-RunAll/Log/sessionInfo.csv") {
  sessionNames <- c("platform","running","R.version")
  sessionInfoValues <- sessionInfo()[sessionNames]
  sessionInfoValues <- matrix(
    c(
      sessionInfoValues$R.version$version.string,
      sessionInfoValues$platform,
      sessionInfoValues$running,
      Sys.time()
    ),ncol=4
  )
  colnames(sessionInfoValues) <- c(sessionNames,"TimeStamp")
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
