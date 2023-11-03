library("openxlsx")
library("data.table")
library("here")
library("usethis")
library("lubridate")
library("dplyr")



fixLogical <- function(frame, column = "") {
  frame <- mutate(frame, "{column}" := if_else(get(column) == "TRUE" | get(column) == 1, as.logical(TRUE), as.logical(FALSE)))
  return(frame)
}

Load_ExcelTable <- function(myFileName, mySubfolderName, mySheetName, myHeaderRowCount) {
    Header_myDataFrame <- NA
    Header_myDataFrame <- read.xlsx (myFileName, sheet = mySheetName, 
                                     colNames = TRUE)[(1:max(1, myHeaderRowCount - 1)), ]

    n_Col_Header <- ncol(Header_myDataFrame)

    ## Read values of the table
    myDataFrame <- NA
    myDataFrame <- read.xlsx (myFileName, sheet = mySheetName, 
                             rowNames = FALSE, colNames = FALSE, 
                             startRow = (myHeaderRowCount + 1), 
                             cols = (1:n_Col_Header), 
                             skipEmptyCols = FALSE, na.strings = c("NA", ""))

    n_Col_Data <- ncol(myDataFrame)

    myDataFrame <- if (n_Col_Data < n_Col_Header) {
        cbind(myDataFrame, matrix(data = NA, 
                                  nrow = nrow(myDataFrame), 
                                  ncol = (n_Col_Header - n_Col_Data)))
    } else {
        myDataFrame
    }

    colnames(myDataFrame) <- colnames(Header_myDataFrame)
    rownames(myDataFrame) <- myDataFrame[, 1]
    return(myDataFrame)
}

# -----------------------------------------------------------------------------
#' @title Convert string to R-Date
#' @description Converts a given string to a R Date (taken from the openxlsx2 package)
# -----------------------------------------------------------------------------
convert_date <- function(x, origin = "1900-01-01", ...) {
    # use as.integer to only get the integer part of a number.  in openxml dates are integers only.
    x <- as.integer(x)
    notNa <- !is.na(x)
    earlyDate <- x < 60

    if (origin == "1900-01-01") {
        x[notNa] <- x[notNa] - 2
        x[earlyDate & notNa] <- x[earlyDate & notNa] + 1
    }

    as.Date(x, origin = origin, ...)
}

# ----------------------------------------------------------------------------- Removes the Tabula naming prefix
# -----------------------------------------------------------------------------
convertSheetName = function(sheetName) {
    sheetName <- sub("Tab\\.?", "", sheetName) 
    # 2023-11-03 / TL: Changed, the naming of these sheets should include the "Const." 
    # These datasets are not to be changed or supplemented by users.    
#    sheetName <- sub("Tab\\.(Const\\.)?", "", sheetName)
    sheetName
}

# ----------------------------------------------------------------------------- Returns a vector of worksheet names
# ------------------list-----------------------------------------------------------
listsheets <- function() {
  fh <- "data-raw/tabula-values.xlsx" # 2023-11-03 / TL: Changed
    # fh <- system.file("extdata", "copy_tabula-values.xlsx", package = "tabulaconv")
    # fh <- system.file('extdata', 'test.xlsx', package = 'tabulaconv')
    names <- openxlsx::getSheetNames(fh)
    names
}

# ----------------------------------------------------------------------------- Process sheets
# -----------------------------------------------------------------------------
process <- function(sheets) {
  skip <- c("BlankSheet", "Info", "Tab.AuxCalc.Climate") # 2023-11-03 TL / changed
  # skip <- c("BlankSheet", "Info", "Tab.Info.Sheet", "Tab.Const.DataFormat",
    # "Tab.Const.Orientation","Tab.AuxCalc.Climate" # temp
    # )

    sheets <- sheets[!sheets %in% skip]
    sheetlist <- c()
    for (sheet in sheets) {
        sheetlist <- c(sheetlist, sheet)
    }
    sheetlist
}

# ----------------------------------------------------------------------------- Convert2 sheets
# -----------------------------------------------------------------------------
convert <- function(sheetname) {
     fh <-  "data-raw/tabula-values.xlsx" # 2023-11-03 / TL: changed
     #    fh <- system.file("extdata", "copy_tabula-values.xlsx", package = "tabulaconv")

    data <- openxlsx::read.xlsx(xlsxFile = fh, sheet = sheetname, detectDates = FALSE, rowNames = FALSE, colNames = TRUE,
        skipEmptyRows = FALSE, skipEmptyCols = FALSE, na.strings = c("NA", "#", ""))

    dt <- as.data.table(data[10:nrow(data), ], na.rm = FALSE, sorted = FALSE, keep.rownames = FALSE)

    for (i in 1:ncol(data)) {
        #
        colname <- colnames(data)[i]
        coltype <- toupper(data[5, i])

        # Sad ...
        if (is.na(coltype) || coltype == "") {
            msg <- paste("Missing Coltype for Column:", colname)
            print(msg)
            next
        }

        #
        print("---------------------------------------------------")
        print(paste("COL:", i, "Name:", colname, "Type:", coltype))
        print("---------------------------------------------------")

        # Conversion of specific types
        if (coltype == toupper("VARCHAR")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.character), .SDcols = colname]
        } else if (coltype == toupper("INTEGER")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.numeric), .SDcols = colname]
        } else if (coltype == toupper("REAL")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.numeric), .SDcols = colname]
        } else if (coltype == toupper("DATE")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, convert_date), .SDcols = colname]
        } else if (coltype == toupper("TEXT")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.character), .SDcols = colname]
        } else if (coltype == toupper("TEXT")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.character), .SDcols = colname]
        } else if (coltype == toupper("PICTURE")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.character), .SDcols = colname]
        } else {
            # Bail out
            error <- paste("Unmatched column type:", colname, "- Type:", coltype, "- Sheet:", sheetname)
            stop(error)
        }
    }

    df <- as.data.frame(dt)

    print(paste0("Setting rowsnames for: ", sheetname))

    row.names(df) <- df[,1]

    df
}


convertBuildingData <- function(sheetname = "Data.Building") {
    filename <- system.file("extdata", "Building-Data.xlsx", package = "tabulaconv")

    data <- openxlsx::read.xlsx(xlsxFile = filename, sheet = sheetname, detectDates = FALSE, rowNames = FALSE, colNames = TRUE,
        skipEmptyRows = FALSE, skipEmptyCols = FALSE, na.strings = c("NA", "#", ""))

    lastRow <- nrow(data)
    lastRow <- 103
    dt <- as.data.table(data[0:lastRow, ], na.rm = FALSE, sorted = FALSE, keep.rownames = FALSE)

    for (i in 1:ncol(data)) {
        #
        colname <- colnames(data)[i]
        coltype <- toupper(data[98, i])

        colVariant <-  toupper(data[90, i])

        if (colVariant == "Output") {
            print("Output")
        }
        if (colVariant == "Input") {
            print("Input")
        }

        # Sad ...
        if (is.na(coltype) || coltype == "") {
            msg <- paste("Missing Coltype for Column:", colname)
            print(msg)
            next
        }

        #
        print("---------------------------------------------------")
        print(paste("COL:", i, "Name:", colname, "Type:", coltype))
        print("---------------------------------------------------")

        # Conversion of specific types
        if (coltype == toupper("VARCHAR")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.character), .SDcols = colname]
        } else if (coltype == toupper("INTEGER")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.numeric), .SDcols = colname]
        } else if (coltype == toupper("REAL")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.numeric), .SDcols = colname]
        } else if (coltype == toupper("DATE")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, convert_date), .SDcols = colname]
        } else if (coltype == toupper("DATETIME")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, convert_date), .SDcols = colname]
        } else if (coltype == toupper("TEXT")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.character), .SDcols = colname]
        } else if (coltype == toupper("BOOLEAN")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.logical), .SDcols = colname]
        } else if (coltype == toupper("PICTURE")) {
            dt[, colname] <- dt[1:.N, lapply(.SD, as.character), .SDcols = colname]
        } else {
            # Bail out
            error <- paste("Unmatched column type:", colname, "- Type:", coltype, "- Sheet:", sheetname)
            stop(error)
        }
    }

    df <- as.data.frame(dt)

    print(paste0("Setting rowsnames for: ", sheetname))

    row.names(df) <- df[,1]

    buildingData <- df

    compress <- "bzip2"
    compression_level <- 9
    save(buildingData, file = file.path(system.file("data", ".", package = "tabulaconv"), "building.data.rda"), compression_level = compression_level, compress = compress)
}


convertTypes <- function(input) {

    rawInput <- input[103:nrow(input),]

     # Remove metadata
     for (i in 1:ncol(input)) {
        #
        colname <- colnames(input)[i]
        coltype <- toupper(input[98:98, ..i])

        # Sad ...
        if (is.na(coltype) || coltype == "") {
            msg <- paste("Missing Coltype for Column:", colname)
            print(msg)
            next
        }


        # print("---------------------------------------------------")
        # print(paste("Name:", colname, "Type:", coltype))
        # print("---------------------------------------------------")

        # Conversion of specific types
        if (coltype == toupper("VARCHAR")) {
            rawInput[, colname] <- rawInput[1:.N, lapply(.SD, as.character), .SDcols = colname]
        } else if (coltype == toupper("INTEGER")) {
            rawInput[, colname] <- rawInput[1:.N, lapply(.SD, as.numeric), .SDcols = colname]
        } else if (coltype == toupper("REAL")) {
            rawInput[, colname] <- rawInput[1:.N, lapply(.SD, as.numeric), .SDcols = colname]
        } else if (coltype == toupper("DATE")) {
            rawInput[, colname] <- rawInput[1:.N, lapply(.SD, convert_date), .SDcols = colname]
        } else if (coltype == toupper("DATETIME")) {
            rawInput[, colname] <- rawInput[1:.N, lapply(.SD, convert_date), .SDcols = colname]
        } else if (coltype == toupper("TEXT")) {
            rawInput[, colname] <- rawInput[1:.N, lapply(.SD, as.character), .SDcols = colname]
        } else if (coltype == toupper("BOOLEAN")) {
            #rawInput[, colname] <- rawInput[1:.N, sapply(.SD, conv), .SDcols = colname]
            rawInput <- fixLogical(rawInput, colname)
        } else if (coltype == toupper("PICTURE")) {
            rawInput[, colname] <- rawInput[1:.N, lapply(.SD, as.character), .SDcols = colname]
        } else {
            # Bail out
            error <- paste("Unmatched column type:", colname, "- Type:", coltype, "- Sheet:", sheetname)
            stop(error)
        }
    }

     rawInput
}


prepareBuildingData <- function() {
    # Input file
    filename <- system.file("extdata", "Building-Data.xlsx", package = "tabulaconv")

    print("Starting ...")

    # Export options
    compress <- "bzip2"
    compression_level <- 9

    # Open XLSX
    data <- openxlsx::read.xlsx(
        xlsxFile = filename,
        sheet = "Data.Building",
        detectDates = FALSE,
        rowNames = FALSE,
        colNames = TRUE,
        skipEmptyRows = FALSE,
        skipEmptyCols = FALSE,
        na.strings = c("NA", "#", "")
    )

    dataTable <- as.data.table(data[1:nrow(data), ], na.rm = FALSE, sorted = FALSE, keep.rownames = FALSE)


    # Split sheet into 2 distinct dataframes based on the "Input" / "Output" annotation
    input <- dataTable[
        ,
        .SD,
        .SDcols = colnames(data[which( data[90,] == "Input" | data[90,] == "Type_Datafield_WebTool" & !is.na(data[90,]))])
    ]


    output <- dataTable[
        ,
        .SD,
        .SDcols = colnames(data[which( data[90,] == "Output" | data[90,] == "Type_Datafield_WebTool" & !is.na(data[90,]))])
    ]

    # setup proper types
    buildingDataInput <- convertTypes(input)
    buildingDataOutput <- convertTypes(output)

    # Save Outputs
    save(buildingDataInput, file = file.path(system.file("data", ".", package = "tabulaconv"), "building.input.data.rda"), compression_level = compression_level, compress = compress)
    save(buildingDataOutput, file = file.path(system.file("data", ".", package = "tabulaconv"), "building.output.data.rda"), compression_level = compression_level, compress = compress)

    print("Done.")

}


#' @export
WriteScript_DataPackage_tabuladata <- function () {
  # 2023-11-03 / TL 
  
  load ("data/info.sheet.rda") 
  
  n_Sheet <- nrow (info.sheet) 
  Sheets_Skip <- c ("Info", "Tab.AuxCalc.Climate") 

  DataPackageScript <- ""
  TableDescription <- ""
  
  info.sheet$Name_Sheet_R <-   
      gsub ("Tab.", "", info.sheet$Name_Sheet)
  
  i_Sheet <- 2
  for (i_Sheet in (1 : n_Sheet)) {
    if (! info.sheet$Name_Sheet [i_Sheet] %in% Sheets_Skip) {
      
      ## Data package script
      DataPackageScript <- 
        append (
          DataPackageScript, 
          paste0 ("#\' ",
            info.sheet$Name_Sheet_R [i_Sheet]
          )
      )
      DataPackageScript <- 
        append (DataPackageScript, "#\' ",
      )
      DataPackageScript <- 
        append (
          DataPackageScript, 
          paste0 ("#\' ",
                  info.sheet$Content [i_Sheet]
          )
        )
      DataPackageScript <- 
        append (
          DataPackageScript, 
          paste0 ("\'",
                  tolower (info.sheet$Name_Sheet_R [i_Sheet]),
                  "\'"
          )
        )
      DataPackageScript <- 
        append (DataPackageScript, "",
        )
      
      ## Description of the tables for README.md
      TableDescription <- 
        append (
          TableDescription, 
          paste0 ("- **\'",
                  tolower (info.sheet$Name_Sheet_R [i_Sheet]),
                  "\':**"
          )
        )
      TableDescription <- 
        append (
          TableDescription, 
          paste0 ("    ",
                  info.sheet$Content [i_Sheet]
          )
        )
      TableDescription <- 
        append (
          TableDescription, 
          paste0 ("    Excel sheet name: \'Tab.",
                  info.sheet$Name_Sheet_R [i_Sheet],
                  "\'"
          )
        )
      TableDescription <- 
        append (TableDescription, "",
        )
      
    } # End if 
  } # End loop by i_Sheet
  
  write.csv (DataPackageScript, file = "data-raw/R-DataPackage-Script.csv",
             row.names = FALSE)
  write.csv (TableDescription, file = "data-raw/TableDescription.csv",
             row.names = FALSE)
  
} 


#' @export
runConversion <- function(path = "") {
    if (path == "") {
        path <- here::here()
    }

    if (path == "") {
        stop("Empty path")
    }

    # Collect all sheets
    sheets <- listsheets()
    sheets <- process(sheets)

    # Process single sheet
    # sheet <- 'Tab.System.HD'
    # sheet <- 'Tab.AuxCalc.Climate'
    # result <- convert(sheetname = sheet)
    # varName <- tolower(convertSheetName(sheet))
    # assign(varName,result)
    # save(list = varName, file = file.path(path, paste0(sheet,'.rda')))

    # stop("Done ...")

    # Process all sheets
    for (sheet in sheets) {
        print(paste0("Processing sheet: ", sheet))
        result <- convert(sheetname = sheet)
        varName <- tolower(convertSheetName(sheet))
        assign(varName, result)
        filename <- paste0 ("data/", varName, ".rda") # 2023-11-03 / TL: changed
        # filename <- file.path(system.file
        #                       ("data", ".",
        #                         package = "tabulaconv"),
        #                       paste0(varName, ".rda"))
        save(list = varName, file = filename)
    }
}


ccDateTime <- function(d) {
    lubridate::ymd_hms(d, locale = "de_DE", tz = "Europe/Berlin")
}

ccDate <- function(d) {
    lubridate::ymd_hms(d, locale = "de_DE", tz = "Europe/Berlin")
}

#' @export
runClimateConversion <- function() {

    filename <- system.file("extdata", "Gradtagzahlen-Deutschland.xlsx", package = "tabulaconv")

    compress <- "bzip2"
    compression_level <- 9

    #############################################
    # list.station.ta
    #############################################
    list.station.ta <- Load_ExcelTable(myFileName = filename, mySubfolderName = "Input/Climate", mySheetName = "List.Station.TA",
        myHeaderRowCount = 1)

    # 2011-02-01
    list.station.ta$Date_Start_DataBase <- ccDate(list.station.ta$Date_Start_DataBase)
    # 2022-01-23
    list.station.ta$Date_End_DataBase <- ccDate(list.station.ta$Date_End_DataBase)

    save(list.station.ta, file = file.path(system.file("data", ".", package = "tabulaconv"), "list.stationta.rda"),
         compression_level = compression_level, compress = compress)

    #############################################
    # stationmapping
    #############################################
    tab.stationmapping <- Load_ExcelTable(myFileName = filename, mySubfolderName = "Input/Climate", mySheetName = "Tab.StationMapping",
        myHeaderRowCount = 1)

    save(tab.stationmapping, file = file.path(system.file("data", ".", package = "tabulaconv"), "tab.stationmapping.rda"),
         compression_level = compression_level, compress = compress)

    #############################################
    # data.ta.hd
    #############################################
    data.ta.hd <- Load_ExcelTable(myFileName = filename, mySubfolderName = "Input/Climate", mySheetName = "Data.TA.HD",
        myHeaderRowCount = 1)

    # 1891-01-01
    data.ta.hd$Date_Start_DataBase <- ccDate(data.ta.hd$Date_Start_DataBase)
    # 2011-03-31
    data.ta.hd$Date_End_DataBase <- ccDate(data.ta.hd$Date_End_DataBase)

    # Date_Compiled (2015-02-17 00:00:00)
    data.ta.hd$Date_Compiled <- ccDateTime(data.ta.hd$Date_Compiled)

    save(data.ta.hd, file = file.path(system.file("data", ".", package = "tabulaconv"), "data.ta.hd.rda"),
         compression_level = compression_level, compress = compress)

    #############################################
    # data.sol
    #############################################
    data.sol <- Load_ExcelTable(myFileName = filename, mySubfolderName = "Input/Climate", mySheetName = "Data.Sol", myHeaderRowCount = 1)

    # 1891-01-01
    data.sol$Date_Start_DataBase <- ccDate(data.sol$Date_Start_DataBase)

    # 2011-03-31
    data.sol$Date_End_DataBase <- ccDate(data.sol$Date_End_DataBase)

    # Time!
    data.sol$Date_Compiled <- ccDateTime(data.sol$Date_Compiled)

    save(data.sol, file = file.path(system.file("data", ".", package = "tabulaconv"), "data.sol.rda"), compression_level = compression_level, compress = compress)

    #############################################
    # Tab.Estim.Sol.Orient
    #############################################
    tab.estim.sol.orient <- Load_ExcelTable(myFileName = filename, mySubfolderName = "Input/Climate", mySheetName = "Tab.Estim.Sol.Orient", myHeaderRowCount = 1)

    # Time!
    tab.estim.sol.orient$Date_Change <- ccDateTime(tab.estim.sol.orient$Date_Change)

    save(tab.estim.sol.orient, file = file.path(system.file("data", ".", package = "tabulaconv"), "tab.estim.sol.orient.rda"), compression_level = compression_level, compress = compress)

}

