library(tidyverse)

library(readxl)

library(dplyr)

if(!require('openxlsx')) install.packages('openxlsx')

library('openxlsx')

 

options(dplyr.summarise.inform = FALSE)

 

s <- createStyle(numFmt = "0.0000")

###########################The input output folders and files################

folder <- "output"

excel_output_name <- "final_results_r.xlsx"

default_dir <- "W:/IE-ANALYSIS/NZIP 1.0/2022 update/Analysis/EEP/*.xlsx"

 

#############################################

# Include some useful general functions######

#############################################

 

# Check if file is already open

file.closed <- function(path) {

  suppressWarnings(

    "try-error" %in% class(

    try(file.rename(excel_output_name, excel_output_name),

          silent = TRUE

      )

    )

  )

}

 

# Reading data from files to a dataframe

make_df<-function(filename){

  df = tryCatch(

    { # try block

      if (str_sub(filename,-3,-1) =="csv"){

        df<-read.csv(filename,fileEncoding = "latin1")

      }

      else{

        df<-read_excel(filename)

      }

    },

    error=function(cond){return(filename)} #grab the filename if there is an error

    )

 

    if (class(df)[1] == 'character') {

      print(df)

    } else{return(df)}

  }

 

'%!in%' <- Negate('%in%') #useful for checking output

 

#################### Import scenario specific data ##########################

cat("select input file and working directory")

infile <- file.choose()

wdir <- dirname(infile)

 

#setwd("W:/IE-ANALYSIS/NZIP 1.0 Spring 2022 Pathway update/Final Reference Scenarios/Git/r data")

setwd(wdir)

 

#Creating a folder if it doesn't exist

if (file.exists(folder)) {

  cat("The folder already exists")

} else {

  dir.create(folder)

}

 

# This is an attempt to check if excel file is open but it is not reliable as it can cause system hangs in Rstudio

#if(file.closed(excel_output_name)){

#  xlsxTemplate <- system.file(excel_output_name, package = "openxlsx")

#} else {

#  message(paste("Please close the file ", excel_output_name))

#  cat("select input file and working directory")

#  infile <- file.choose()

#  wdir <- dirname(infile)

#}

start_time <- Sys.time() # starts timer after file opening procedure as this is variable depending on user

 

xlsxTemplate <- system.file(excel_output_name, package = "openxlsx")

 

slabels <- make_df(infile)

 

#Create variables for scenario comparison based on the excel lookup table

 

years <- as.vector(na.omit(slabels[["Years"]])) # The years to be compared

 

scenarios <- as.vector(na.omit(slabels[["Scenarios"]])) # The scenarios to be compared

scenario_names <- as.vector(na.omit(slabels[["Scenario_names"]])) # The scenarios to be compared

files <- as.vector(na.omit(slabels[["Files"]])) # The scenarios to be compared

 

columns <- as.vector(na.omit(slabels[["Column_labels"]])) # column labels to be used

variables <- as.vector(na.omit(slabels[["Summary_variables"]])) # variables to be used

 

derived_calcs <- data.frame(differences = as.vector(na.omit(slabels[["Difference_variables"]])),

                                minuends = as.vector(na.omit(slabels[["Minuends"]])),

                                subtrahends = as.vector(na.omit(slabels[["Subtrahends"]])))

 

derived_cols <- match(derived_calcs[,1], outputs, nomatch = NA)

 

outputs <- as.vector(na.omit(slabels[["Outputs"]])) # output labels to be used

outputs_labels <- as.vector(na.omit(slabels[["Outputs_labels"]])) # output labels to be used

yearoi <- as.vector(na.omit(slabels[["Yearoi"]])) # output labels to be used

split_tech <- as.vector(na.omit(slabels[["Split_by_tech"]])) # output labels to be used

verbose <- as.vector(na.omit(slabels[["Verbose"]])) # output labels to be used to output all intermediate files

#names(calculations_list) <- variables #This is not correct if the variables are not all valid

 

sec_grp <- as.vector(na.omit(slabels[["Sector_lookup_type"]]))

sec_lookup <- as.vector(na.omit(slabels[["Sector_lookup"]]))

all_sectors <- data.frame(sec_grp)

 

sector_groups <- data.frame(Sector_group = sec_grp, Element_sector = sec_lookup)

sector_groups <- filter(sector_groups, Sector_group != "*")

 

 

flag_names <- as.vector(na.omit(slabels[["Flag_names"]]))

flags <- data.frame(Element_sector = sec_lookup)#BEIS = BEIS_sectors, Intensive = intensive, Optional = optional)

group_flags <- data.frame(Element_sector = sec_grp)

 

for(n in 1:length(flag_names)){

  flags[n+1] <- as.vector(na.omit(slabels[[flag_names[n]]]))

  colnames(flags)[n+1] <- flag_names[n]

  group_flags[n+1] <- as.vector(na.omit(slabels[[flag_names[n]]]))

  colnames(group_flags)[n+1] <- flag_names[n]

}

 

flags <- flags[ !duplicated(flags[, c("Element_sector")], fromLast=T),]

group_flags <- distinct(group_flags, Element_sector, .keep_all = TRUE)

 

#Create an output workbook for scenario comparison

wb <- createWorkbook("output")

ws_scen <- paste0("Scenarios ", years[[1]], "-", years[[2]])

addWorksheet(wb, ws_scen)

 

ws_yoi <- paste0("Scenarios ", yearoi[1])

addWorksheet(wb, ws_yoi)

 

addWorksheet(wb, "Pathway_Charts")

addWorksheet(wb, "Pathways_Ordered")

 

# Create the file names for loading

file_names <- c() # Blank vector

file_names_totals <- c()

file_stems <- c()

 

#Create files if the file name doesn't exist in the input sheet

if (length(files)==0) {

  for (i in 1:length(scenarios)){

    name <- paste0("Reference_scenario_", scenarios[i],".csv") # this allows for scenarios to be compared other than 1,2,3 sequential scenarios

    file_names <- c(file_names, name)

    name2 <- paste0(getwd(),"/",folder,"/","totals_scenario", scenarios[i], "_", ".xlsx")

    file_names_totals <- c(file_names_totals, name2)

    file_stems <- c(file_stems, scenarios[i])

    if(is.na(scenario_names[i])) {

      scenario_names[i] <- file_names[i]

    }

    rm(name)

    rm(name2)

  }

 

} else {

  for (i in 1:length(files)) {

    name <- files[i] # this allows for scenarios to be compared other than 1,2,3 sequential scenarios

    file_names <- c(file_names, name)

    stem <- sub('\\.csv$', '', files[i])

    name2 <- paste0(getwd(),"/",folder,"/","totals_", stem,".xlsx")

    file_names_totals <- c(file_names_totals, name2)

    file_stems <- c(file_stems, stem)

    if(is.na(scenario_names[i])) {

      scenario_names[i] <- file_names[i]

    }

    rm(name)

    }

  scenarios <- file_stems

  }

rm(i)

 

# label the data with positive flag values

is_BEIS <- filter(flags, BEIS == 1) %>% select(Element_sector)

is_intensive <- filter(flags, Intensive == 1) %>% select(Element_sector)

is_optional <- filter(flags, Optional == 1) %>% select(Element_sector)

 

# Create a list (sector_list) with all the sector groups

 

unique_sectors <- all_sectors %>% distinct()

names(unique_sectors) <- "sectors"

 

sum_sectors <- as.data.frame(unique_sectors) # The variable to store totals for each sector

 

sectors <- as.vector(slabels[slabels$Sector_lookup_type==unique_sectors[1,1],'Sector_lookup'])

sector_list <- list(sectors)

list_names <- list(unique_sectors[1,1])

 

for(i in 2:nrow(unique_sectors)){

  new_sector <- as.vector(slabels[slabels$Sector_lookup_type==unique_sectors[i,1],'Sector_lookup'])

  sector_list[[i]] <- new_sector

  list_names[i] <- unique_sectors[i,1]

}

rm(i)

names(sector_list) <- list_names # name the sectors for reference

number_of_sectors <- length(sector_list)

 

 

select_from <- as.vector(na.omit(slabels[["Select_from"]]))

select_to <- as.vector(na.omit(slabels[["Select_to"]]))

##############################################################

# Get meta data - split this into two parts and pulled in electricity info

 

meta_filter_col <- as.vector(na.omit(slabels[["Meta-data_filter1"]])) # list of columns to be retrieved "Total direct emissions abated (MtCO2e)"              "Post REEE baseline emissions (MtCO2e)"            "Baseline emissions (MtCO2e)"     "Remaining direct emissions per year (MtCO2e)"              "AM capex (£m)"            "AM fuel costs (£m)"

 

meta_filter_var <- as.vector(na.omit(slabels[["Filter3"]])) # used to filter data

 

meta_data1 <- make_df("NZIP_meta_data.xlsx") %>%

  filter(`Column Label` %in% meta_filter_col)

 

filter2 <- as.vector(na.omit(slabels[["Meta-data_filter2"]]))

 

meta_data2 <- make_df("NZIP_meta_data.xlsx") %>%

    filter(`Column Label` %in% filter2) %>%

    filter(`Column Header 1` >= years[1] & `Column Header 1` <= years[2]) # This has been changed from 2020 to 2050 so that it can be adjusted from the input data file

 

meta_data3 <- make_df("lookups.xlsx")

technologies <- unique(meta_data3$Technology)

 

meta_data <- rbind(meta_data1, meta_data2)

rm(meta_data1, meta_data2)

 

gcols = enquo(meta_filter_var)

 

################################The output data variables ##################################

############################################################################################

 

calc <- as.data.frame(c(FALSE))

sresults <- list()

results <- list()

yoi_sresults <- list()

yresults <- list()

ytresults <- list() # a list of dataframes for each scenario to hold the sector, technology year timeseries

yoiresults <- list()

ssy_results <- data.frame(matrix(nrow = 0, ncol = (years[2]-years[1]+3)))

 

#sty_results <- data.frame(matrix(nrow = 0, ncol = (years[2]-years[1]+3)))

colnames(ssy_results) <- c("factor", "scenario", years[1]:years[2])

loop <- 1 # Loop holds the sequence of loops and scenario the scenario number

 

# Use the filter to create the common column list from filter1

common_filter <- c()

for(tf in meta_filter_col){

  common_filter <- c(common_filter, meta_data$`Col Index`[meta_data$`Column Label` == tf])

}

############################### Loop through the Scenario ############################

######################################################################################

 

for (scenario in scenarios){

  writeLines("")

  message(paste("Scenario ",scenario, " ", loop, "of", length(scenarios)))

  #scenario=1  #use this line for debugging

  stotals_list <- list() # A list to hold the totals 

  calculations_list <- list()


  ma <- matrix(ncol = 0, nrow = 0)

  stotals_tech_years <- data.frame(ma)# This is the reset for each sector technology times series dataframe between scenarios

  stotals <- data.frame(ma)

 

  ssv <- data.frame(matrix(NA, nrow = nrow(unique_sectors), ncol = length(variables)+1))

  ssv_tech <- data.frame(matrix(0, nrow = nrow(unique_sectors), ncol = length(technologies)))

  ssv_tech_yoi <- data.frame(matrix(0, nrow = nrow(unique_sectors), ncol = length(technologies)))

 

  ssv_yoi <- data.frame(matrix(NA, nrow = nrow(unique_sectors), ncol = length(variables)+1))

  ssy <- data.frame(matrix(0, nrow = (years[2]-years[1]+1), ncol = length(variables)))

  ss_yoi <- data.frame(matrix(0, nrow = 1, ncol = length(variables)))

 

  colnames(ssy) <- variables

  rownames(ssy) <- c(years[1]:years[2])

 

  colnames(ss_yoi) <- variables

  rownames(ss_yoi) <- c(yearoi[1])

 

  colnames(ssv) <- c(variables, "scenario")

  rownames(ssv) <- unique_sectors$sectors

 

  colnames(ssv_tech) <- technologies

  rownames(ssv_tech) <- unique_sectors$sectors

 

  colnames(ssv_tech_yoi) <- technologies

  rownames(ssv_tech_yoi) <- unique_sectors$sectors

 

  colnames(ssv_yoi) <- c(variables, "scenario")

  rownames(ssv_yoi) <- unique_sectors$sectors

 

  # Import data one scenario at a time

  data <- read.csv(file_names[loop])

 

  ssv[,"scenario"] <- scenario  

  ssv_yoi[,"scenario"] <- scenario

 

  stotals_tech <- list()

 

####################################################################################### 

  #Loop through the types of returns# This is only necessary because we are flexibly assigning groups and flags etc

####################################################################################

for(v in 1:length(variables)){

#  v=1

  var_filter <- common_filter #common filter includes those columns that match each time which is appended with each variable for var_filter 

  message(paste("Looping through the variables",v, "of", length(variables)))

    col <- variables[v]

    cf <- columns[v]

    

    var_filter <- c(var_filter, meta_data$`Col Index`[meta_data$`Column Label` == cf])

   

    calc <- data %>% select(all_of(var_filter))

   

    # Rename columns

    rename <- append(meta_filter_var, meta_data$`Column Header 1`[meta_data$`Column Label` == cf])

    colnames(calc) <- rename

   

    grp_col = append(meta_filter_var, "year") # year is added for these calculations but not in meta_filter_var as the year is not always used to filter

#    dots <-lapply(grp_col, as.symbol)

 

    message(paste("Calculating:",cf)) # check it is valid

 

    calc <- calc %>%

      gather(year, value, -!!gcols) %>%

      group_by_at(vars(one_of(grp_col))) %>%

      summarise(temp = sum(value, na.rm = T)) %>%

      ungroup()

 

    grp_calc <- subset(calc, Element_sector %in% sector_groups$Element_sector)

    calc_temp <- grp_calc

    calc_temp[] <- lapply(grp_calc, function(x) sector_groups$Sector_group[match(x, sector_groups$Element_sector)])

    grp_calc$Element_sector <- calc_temp$Element_sector

   

    calc_tech <- merge(grp_calc, meta_data3, by = "selected_option", sort = F, All.x =T) ###########################adding technology type to the extracted data

   

    calc_yoi <- grp_calc %>%

      filter(year == yearoi[1])

    

    calc_tech_yoi <- calc_tech %>%

      filter(year == yearoi[1])

   

    totals <- grp_calc %>%

      group_by(Element_sector) %>%

      summarise(temp = sum(temp, na.rm = T))

   

    totals_tech <- calc_tech %>%

      group_by(Element_sector, Technology) %>%

      summarise(temp = sum(temp, na.rm = T))

   

    totals_tech_yoi <- calc_tech_yoi %>%

      group_by(Element_sector, Technology) %>%

      summarise(temp = sum(temp, na.rm = T))   

    

    totals_tech_years <- calc_tech %>%

      group_by(Element_sector, Technology, year) %>%

      summarise(temp = sum(temp, na.rm = T))

     

    colnames(totals_tech_years)[which(names(totals_tech_years) == "temp")] <- variables[v]

   

    totals_yoi <- calc_yoi %>%

      group_by(Element_sector) %>%

      summarise(temp = sum(temp, na.rm = T))

 

    totals_years <- grp_calc %>%

      group_by(year) %>%

      summarise(temp = sum(temp, na.rm = T))

   

    ma <- matrix(ncol = 0, nrow = 0)

    stotals_years <- data.frame(ma)

   

    for(i in 1:length(list_names)) {

      if(i==1) {

        st <- unlist(list_names) # This is the most reliable way of getting totals (so far discovered) - otherwise the totals were higher than the sum

      }

      else{

        st <- list_names[[i]]

      }

#    } 

      stotals <- subset(totals, (Element_sector %in% st))

      stotals_yoi <- subset(totals_yoi, (Element_sector %in% st))

      ssv[i,v] <- sum(stotals[[2]])

      ssv_yoi[i,v] <- sum(stotals_yoi[[2]])

     

      if(variables[v]==split_tech) {

        sparse_tech <- group_by(totals_tech, Element_sector, Technology) %>% summarise(sum(temp))

        total_tech <- totals_tech %>% group_by(Technology) %>% summarise(sum(temp))

        total_tech["Element_sector"] <- "*"

        sparse_tech <- bind_rows(sparse_tech, total_tech)

        sparse_tech_yoi <- group_by(totals_tech_yoi, Element_sector, Technology) %>% summarise(sum(temp))

        total_tech_yoi <- totals_tech_yoi %>% group_by(Technology) %>% summarise(sum(temp))

        total_tech_yoi["Element_sector"] <- "*"

        sparse_tech_yoi <- bind_rows(sparse_tech_yoi, total_tech_yoi)

       

        for(j in 1:length(sparse_tech$Technology)) {

          n <- as.character(sparse_tech[j,1])

          m <- as.character(sparse_tech[j,2])

          ssv_tech[n,m] <- sparse_tech[j,3]

        }

        #ssv_tech[,1] <- colSums(ssv_tech)

        for(j in 1:length(sparse_tech_yoi$Technology)) {

          n <- as.character(sparse_tech_yoi[j,1])

          m <- as.character(sparse_tech_yoi[j,2])

          ssv_tech_yoi[n,m] <- sparse_tech_yoi[j,3]

        }

      }

     

#      stotals_tech_years <- bind_rows(stotals_tech_years, subset(totals_tech_years, (Element_sector %in% st)))

    }  

    

    stotals_years <- rbind(stotals_years, totals_years)

   

    total_yoi <- stotals_years %>%

      filter(year == yearoi[1])

    

    names(stotals)[names(stotals) == 'temp'] <- variables[v]

    names(total_yoi)[names(total_yoi) == 'temp'] <- variables[v]

   

    ssy[v] <- stotals_years[2]

 

    ss_yoi[v] <- total_yoi[2]

   

#########Subset data for output to compare scenarios

    

    names(calc)[names(calc) == 'temp'] <- variables[v]

   

    if (v == 1){

      calculations_list <- calc

#      stotals_list <- stotals

    } else {

      calculations_list <- append(calculations_list, calc[, ncol(calc)]) # adds the unique last column

#     stotals_list <- append(stotals_list, stotals)

    }

 

  } # variables #loop

  ###################################################################################

 

  for(d in 1:nrow(derived_calcs)){

    ssy[derived_calcs[d,1]] <- ssy[derived_calcs[d,2]] - ssy[derived_calcs[d,3]]

    ss_yoi[derived_calcs[d,1]] <- ss_yoi[derived_calcs[d,2]] - ss_yoi[derived_calcs[d,3]]

    ssy <- ssy %>% relocate(derived_calcs[d,1], .after = derived_calcs[d,3])

    ss_yoi <- ss_yoi %>% relocate(derived_calcs[d,1], .after = derived_calcs[d,3])

  }

  

  scen <- rep(scenario_names[loop], nrow(calc))

 

  calculations_list$scenario = scen

 

  final_results <- as_tibble(calculations_list)

 

  #derived fields added at the end

  # flags first

  for(n in 2:ncol(flags)){

    trueflags <- filter(flags, flags[n] == 1) %>% select(Element_sector)

    final_results[,names(flags)[n]] <- final_results$Element_sector %in% trueflags[,1] #add Sector flag

  }

  for(d in 1:nrow(derived_calcs)){

    final_results[derived_calcs[d,1]] <- final_results[derived_calcs[d,2]] - final_results[derived_calcs[d,3]]

    final_results <- final_results %>% relocate(derived_calcs[d,1], .after = derived_calcs[d,3])

  }

 

# legacy code needs checking

  final_results[,'Technology'] <- 0 #add BEIS Technology flag

  final_results$Technology=meta_data3$Technology[match(final_results$selected_option,meta_data3$selected_option,nomatch = NA)]

 

#  add the sector group - automatically calls it Sector_group based on name of column in sector_groups which is based on the input spreadsheet minus the * rows

  final_results <- merge(final_results, sector_groups, by='Element_sector')

 

  ssv_tech$scenario <- rep(scenario_names[loop], nrow(ssv_tech))

  ssv_tech_yoi$scenario <- rep(scenario_names[loop], nrow(ssv_tech_yoi))

  if(loop==1) {

    tech_calculations <- ssv_tech

  } else {

    tech_calculations <- rbind(tech_calculations, ssv_tech)

  }

  if (verbose == 1){

    t_file_name <- paste(getwd(),"/",folder,"/","technologies_scenario_", file_stems[loop], ".xlsx", sep = "") 

    write.xlsx(ssv_tech, file=t_file_name, asTable = TRUE, rowNames = TRUE, colNames = TRUE, borders = "rows", overwrite = TRUE)

    }

  write.xlsx(ssy, file=file_names_totals[loop], asTable = TRUE, colNames = TRUE, rowNames = TRUE, borders = "rows", overwrite = TRUE)

  c_file_name <- paste(getwd(),"/",folder,"/","scenario_", file_stems[loop], ".xlsx", sep = "")

  write.xlsx(final_results, file=c_file_name, asTable = TRUE, rowNames = TRUE, colNames = TRUE, borders = "rows", overwrite = TRUE)

 

  #The following provides the option to do a technology grouping of results for specific business reasons such as grouping all the furnace types

  # Create a summarised set of results that reduces the number of selected options for a business request

 

  if(length(select_from) > 0) {

 

  c_file_name2 <- paste(getwd(),"/",folder,"/","scenario_", file_stems[loop], "_selected.xlsx", sep = "")

  filtered_results <- final_results

 

  for(x in 1:length(select_from)) {

    filtered_results$selected_option[filtered_results$selected_option == select_from[x]] <- select_to[x]

  }

 

  write.xlsx(filtered_results, file=c_file_name2, asTable = TRUE, rowNames = TRUE, colNames = TRUE, borders = "rows", overwrite = TRUE)

  }

 

  sresults = append(sresults, list(ssv))# store the sum of variables for each scenario in a list

  yoi_sresults= append(yoi_sresults, list(ssv_yoi))

  yresults= append(yresults, list(ssy))

 

  ssy_transposed <- cbind(factor = colnames(ssy), scenario = rep(scenario_names[loop], length(variables)+nrow(derived_calcs)), t(ssy))

 

  ssy_results <- rbind(ssy_results, ssy_transposed)

 

  yoiresults[[loop]] <- ss_yoi

 

  # Create the calculations list for the scenario

    if (loop == 1){

    results <- calculations_list

  } else {

    results <- rbind(results, calculations_list)

  }

       ###########################################

  ####################################################################################################

  # Write the summary data to the spreadsheet sheet

 

  ns <- length(scenarios)

  r <- 3 # the row to start data output

  c <- 6

  o <- 1

 

  for(out in outputs) {

    writeData(wb, sheet= ws_scen, outputs_labels[o], startRow = r+ns-1, startCol = c)

    writeData(wb, sheet= ws_yoi, outputs_labels[o], startRow = r+ns-1, startCol = c)

    t = which(colnames(ssv) == out)

    if (length(t) != 0) {

      for(k in 1:nrow(unique_sectors)) {

        ro = r+k*ns+loop+k

        writeData(wb, sheet= ws_scen, file_stems[loop], startRow = ro, startCol = 2)

        writeData(wb, sheet= ws_yoi, file_stems[loop], startRow = ro, startCol = 2)

        writeData(wb, sheet= ws_scen, scenario_names[loop], startRow = ro, startCol = 3)

        writeData(wb, sheet= ws_yoi, scenario_names[loop], startRow = ro, startCol = 3)

        writeData(wb, sheet= ws_scen, unique_sectors[k,1], startRow = ro, startCol = 1)

        writeData(wb, sheet= ws_yoi, unique_sectors[k,1], startRow = ro, startCol = 1)

        writeData(wb, sheet= ws_scen, paste0(years[1], " to ",years[2]), startRow = ro, startCol = 4)

        writeData(wb, sheet= ws_yoi, yearoi, startRow = ro, startCol = 4)

        writeData(wb, sheet= ws_scen, ssv[k,t], startRow = ro, startCol = c)

        writeData(wb, sheet= ws_yoi, ssv_yoi[k,t], startRow = ro, startCol = c)

    }

      }

  c=c+1 # c is the column for adding the derived fields later

  o=o+1

  }

  o=1

#Write the technologies on the end of the 

  for(tech in technologies) {

    writeData(wb, sheet= ws_scen, technologies[o], startRow = r+ns-1, startCol = c)

    writeData(wb, sheet= ws_yoi, technologies[o], startRow = r+ns-1, startCol = c)

      for(k in 1:nrow(unique_sectors)) {

        ro = r+k*ns+loop+k

        writeData(wb, sheet= ws_scen, ssv_tech[k,tech], startRow = ro, startCol = c)

        writeData(wb, sheet= ws_yoi, ssv_tech_yoi[k,tech], startRow = ro, startCol = c)

    }

    c=c+1

    o=o+1

  }

 

loop=loop+1 # allows for non sequential scenarios like 3, 2, 4

 

} # scenario loop##############################################

############################## scenario loop########################################################

 

if (is.list(results) != TRUE) {

  rownames(results) <- scenarios # this is a large matrix with rows from each scenario

}

if (is.list(sresults) != TRUE) {

  names(sresults) <- scenarios # this is a list of sums for each scenario

}

all_variables = colnames(ssy)

ssy_results_ordered <- arrange(ssy_results, factor(factor, levels = all_variables))

 

writeData(wb, sheet= "Pathway_Charts", ssy_results, startRow = 10, startCol = 3)

writeData(wb, sheet= "Pathways_Ordered", ssy_results_ordered, startRow = 10, startCol = 3)

 

# Creating a new table with the comparison

# excel_output_name

 

saveWorkbook(wb, paste0(getwd(),"/",folder,"/",excel_output_name), overwrite = TRUE)

end_time <- Sys.time()

writeLines("")

message(paste("Time taken is:",(end_time - start_time)))
