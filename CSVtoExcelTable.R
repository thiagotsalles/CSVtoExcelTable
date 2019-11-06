"
Script used to export column-organized data of csv files to frequency
tables in Excel.
"

library(xlsx)            # to export to Excel
# oldOpt <- options()      # backup of default options 
# options(OutDec= ",")     # set decimal separator


#### ========================= FUNCTIONS ========================== ####
# Function to specify the parameters of read.csv function
readData <- function(data_csv) {
  data = read.csv(
    data_csv, h=T, sep=";", na.strings="N/D"
  )
  return(data)
}

# Function to mount dataframes with tables
mountTables <- function(data_csv){
  # Reading csv file and mounting empty database for tables
  data = readData("data_example.csv")
  db = vector("list", length(data))
  nams = gsub("\\.", " ", names(data)) # column names
  
  # Mounting the tables and populating the database with them
  for (i in 1 : length(nams)) {
    freq_tab = data.frame(table(data[i])) # table with frequencies
    # Excluding a specific item ('Not applicable')
    #  freq_tab = freq_tab[freq_tab[, 1] != "Not applicable", ]
    
    # Inclusion of percentages
    percent = matrix(nrow=nrow(freq_tab))
    for (j in 1 : length(percent)) {
      percent[j] <- round(freq_tab[j, 2] / sum(freq_tab[, 2]), 4)
    }
    percent_tab <- cbind(freq_tab, percent) # tabela com freq. e %
    percent_tab[, 1] <- as.character(percent_tab[, 1])
    names(percent_tab) <- c(nams[i], "Frequency", "Percentage (%)")
    
    # Inclusion of row with totals
    totals = data.frame("Total", NA, NA)
    for (k in 2 : length(totals)) {
      totals[k] <- round(sum(percent_tab[, k]), 0)
    }
    colnames(totals) <- names(percent_tab)
    # Table with frequency, % and totals
    totals_tab <- rbind(percent_tab, totals) 
    
    # Inclusion of the mean when data is numeric
    is_num = suppressWarnings(!is.na(as.numeric(totals_tab[1, 1])))
    if (is_num) {
      sX = sum(as.numeric(as.character(freq_tab[, 1])) * freq_tab[, 2])
      m = round(sX / totals[, 2], 2)
      m_frame = data.frame("Mean", m, NA)
      colnames(m_frame) <- names(totals_tab)
      m_tab <- rbind(totals_tab, m_frame)
      db[[i]] <- m_tab  # Inclusion of final table in database
    } else {
      db[[i]] <- totals_tab # Inclusion of final table in database
    }
  }
  
  # Names to dataframes
  names(db) <- nams
  
  return(db)
}

# Function to export to Excel. excel_name = filename
exportToExcel <- function(tables, excel_name) {
  wb = createWorkbook()
  tab_name = gsub(".xlsx", "", excel_name)
  sheet = createSheet(wb, sheetName=tab_name)
  cs = CellStyle(wb) + Alignment(h="ALIGN_CENTER")
  
  for (i in 1 : length(tables)) {
    cab = CellStyle(wb) +
      Font(wb, isBold=TRUE) + Alignment(h="ALIGN_CENTER", wrapText=T) +
      Border(position="BOTTOM") +
      setColumnWidth(sheet, colWidth=19,
                     colIndex=(2 + 4 * (i - 1)):(4 + 4 * (i - 1)))
    data = tables[[i]]
    addDataFrame(data, sheet, row.names=F, colnamesStyle=cab,
                 colStyle=list(`1`=cs, `2`=cs, `3`=cs),
                 startRow=2, startColumn=(2 + 4 * (i - 1))
    )
  }
  saveWorkbook(wb, excel_name)
}

#### ========================== EXAMPLE =========================== ####
results <- mountTables("data_example.csv")
exportToExcel(results, "Results.xlsx")


