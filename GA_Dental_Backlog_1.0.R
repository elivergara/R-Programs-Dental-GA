# CENTURION ENTERPRISE DATA SYSTEMS R HEADER
# Project Name: Georgia Dental Backlog
# R File Location/Name:////ashinfodata//Reports and Data//Programs//Georgia//Dental Reporting//R Programs
# Purpose of File: Tidy data and turn to long format to be used in Excel Dahboard
# Authors: E. Vergara
# Creation Date: 27FEB2024
# Revisions (include date, author, and description of revisions):
# 1.0 Init. ~ GA_Dental_Backlog_1.0.R
# 1.5 14MAR2024, Eli Vergara, adding second part, to import Sick Call Aging
###############################################################################
# SET the environment
###############################################################################
# Clear the global environment
rm(list = ls())
# Set current date
current_date <- Sys.Date()
###############################################################################
# Load packages & set directories
###############################################################################
load_libraries <- function() {
    suppressPackageStartupMessages(library(tidyverse))
    suppressPackageStartupMessages(library(readxl))
    suppressPackageStartupMessages(library(lubridate))
    suppressPackageStartupMessages(library(writexl))
}
load_libraries()


# Set paths and working directory
project_dir <- getwd()


set_paths <- function() {
    # Set Working Directory
    rawdata <- paste0(dirname(project_dir), "/Raw Data")
    
    
    exports_path <- paste0(dirname(project_dir), "/Derived Data")
    
    
    return(list(rawdata = rawdata,
                exports_path = exports_path))
}

# Capture the output of set_paths in a variable
paths <- set_paths()

# ... and use the variable to access in_hsr_path
setwd(paths$rawdata)

####################################################
# Work with files list. Retrieve the latest one for import
#####################################################
# Generate files lists for backlog files and also for aging reports
weekly_backlog_files_list <-
    list.files(pattern = "Weekly Backlog", full.names = FALSE)
aging_report_files_list <-
    list.files(pattern = "Aging Report", full.names = FALSE)

# Extract the date from filename and convert to Date object
extract_date <- function(filename) {
    date_string <- strsplit(filename, "\\s+")[[1]][6]
    as.Date(date_string, format = "%m.%d.%y")
}


# Apply the extract_date function created above, to extract dates from imported file
weekly_dates <- lapply(weekly_backlog_files_list, extract_date)

# Use which.max to find the index of the latest date
latest_index <- which.max(weekly_dates)

# Import the latest file using the list index
latest_file <- weekly_backlog_files_list[latest_index]

# Read all sheets from the latest file:
all_sheets <-
    setNames(lapply(excel_sheets(latest_file), function(sheet)
        read_excel(latest_file, sheet = sheet)),
        excel_sheets(latest_file))


##################################################
# Convert the column names into the actual week names
#################################################
for (i in 1:length(all_sheets)) {
    # Step 2: Convert 2nd column onwards in the first row to numeric, handling mixed types
    # Use sapply to attempt conversion to numeric, reverting to character if conversion fails
    date_col_nums <-
        sapply(all_sheets[[i]][1, 2:ncol(all_sheets[[i]])], function(x) {
            # Attempt to convert to numeric, NA if not possible
            as.numeric(as.character(x))
        })
    
    
    # Filter out NAs in case there are non-convertible values (optional, based on your data)
    date_col_nums <- na.omit(date_col_nums)
    
    # Convert the numeric values to dates
    date_col_names <- as.Date(date_col_nums, origin = "1899-12-30")
    
    # Format dates as mm/dd/yyyy or your preferred format
    formatted_dates <- format(date_col_names, "%m/%d/%Y")
    
    # Prepend "Facility" to the list of formatted date column names
    all_col_names <- c("Facility", formatted_dates)
    
    # Step 3: Set the formatted dates along with "Facility" as column names
    colnames(all_sheets[[i]]) <- all_col_names
    
    # Step 4: Now, remove the first row as it's been used to set column names
    all_sheets[[i]] <- all_sheets[[i]][-1, ]
}

########################################################
# Pivot all dataframes in the all_sheets list, from wide into long format
######################################################
# Cast all columns into character type so we can pivot them
for (j in 1:length(all_sheets)) {
all_sheets[[j]] <- all_sheets[[j]] %>%
    mutate(across(-Facility, as.character))
}


# Initialize all_sheets_long tibble
all_sheets_long <- tibble()
# Extract the sheet names from the file
names(all_sheets) <- excel_sheets(latest_file)

for (k in names(all_sheets)) {
# Pivoting the dataframe to long format
    pivoted_df <- all_sheets[[k]] %>%
        mutate(across(everything(), as.character)) %>% # Convert all columns to character to avoid data type issues
        pivot_longer(
            cols = -Facility, # Selects all columns except `Facility`
            names_to = "Date", # New column for the dates
            values_to = "Value" # New column for the values
        ) %>%
        mutate(Category = k) # Add a new column with the name of the sheet
    
    # Append the pivoted dataframe to the accumulating dataframe
    all_sheets_long <- bind_rows(all_sheets_long, pivoted_df)
}



# Remove "Total" rows from long formatted dataframe
all_sheets_long <- all_sheets_long %>%
    filter(!str_detect(Facility, "Total"))


# Convert 'Date' column from m/d/Y format to Date format
all_sheets_long <- all_sheets_long %>%
    mutate(Date = mdy(Date))
    


#####################################################
# Export data in long format
####################################################
# Export final dataframe into excel
export_file_path <- file.path(paths$exports_path, "Dental_long_data.xlsx")
writexl::write_xlsx(x = all_sheets_long, path = export_file_path)



# Aggregate data by Category and Date
aggregated_data <- all_sheets_long %>%
    group_by(Date, Category) %>%
    summarise(Total_Value = sum(as.numeric(Value), na.rm = TRUE)) %>%
    ungroup()

# Convert Date to character for easier formatting
# aggregated_data$Date <- as.character(aggregated_data$Date)

# Export aggregated data to Excel
export_file_path_aggregated <- file.path(paths$exports_path, "Date_aggregated_data.xlsx")
writexl::write_xlsx(x = aggregated_data, path = export_file_path_aggregated)



#####################################################
# End of Weekly Backlog processing
####################################################


