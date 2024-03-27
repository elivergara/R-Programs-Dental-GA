# CENTURION ENTERPRISE DATA SYSTEMS R HEADER
# Project Name: Georgia Dental Backlog
# R File Location/Name:////ashinfodata//Reports and Data//Programs//Georgia//Dental Reporting//R Programs
# Purpose of File: Tidy data and turn to long format to be used in Excel Dashboard
# Authors: E. Vergara
# Creation Date: 27FEB2024
# Revisions (include date, author, and description of revisions):
# 1.0 Init. ~ GA_Dental_Backlog_1.0.R
# 1.5 14MAR2024, Eli Vergara, adding second part, to import Sick Call Aging
# 2.0 18MAR2024 Eli Vergara, completed code to export weekly backlog and aging report, and two aggregated data files
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
# Weekly Backlog
####################################################
# Work with weekly_backlog_files_list. Retrieve the latest one for import
#####################################################
# Generate files lists for backlog files and also for aging reports
# List files without filtering
weekly_backlog_files_list <-
    list.files(pattern = "Weekly Backlog", full.names = FALSE)
aging_report_files_list <-
    list.files(pattern = "Aging Report", full.names = FALSE)

# Filter out files that start with '~'
weekly_backlog_files_list <-
    weekly_backlog_files_list[!grepl("^~", weekly_backlog_files_list)]
aging_report_files_list <-
    aging_report_files_list[!grepl("^~", aging_report_files_list)]



# Extract the date from filename and convert to Date object
extract_date <- function(filename) {
    # Define regular expressions for date patterns
    date_pattern_dot <- "\\d{1,2}\\.\\d{1,2}\\.\\d{2,4}"
    date_pattern_underscore <- "\\d{1,2}_\\d{1,2}_\\d{2,4}"
    
    # Search for date patterns using regular expressions
    date_string <-
        str_extract(filename,
                    paste(date_pattern_dot, date_pattern_underscore, sep = "|"))
    
    # Determine the format based on the separator
    if (grepl("_", date_string)) {
        format <- "%m_%d_%y"
    } else {
        format <- "%m.%d.%y"
    }
    
    # Convert to date object
    as.Date(date_string, format = format)
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
all_col_names <- purrr::map(all_sheets, ~{
    date_col_nums <- sapply(.x[1, 2:ncol(.x)], function(x){
        as.numeric(as.character(x))
    })
    
    date_col_nums <- na.omit(date_col_nums)
    date_col_names <- as.Date(date_col_nums, origin="1899-12-30")
    formatted_dates <- format(date_col_names, "%m/%d/%Y")
    all_col_names <- c("Facility", formatted_dates)
    })
                                                                        
map(all_sheets, ~{
    colnames(.x) <- all_col_names
    .x <- .x[-1]
})

                                                                        # for (i in 1:length(all_sheets)) {
                                                                        #     # Step 2: Convert 2nd column onwards in the first row to numeric, handling mixed types
                                                                        #     # Use sapply to attempt conversion to numeric, reverting to character if conversion fails
                                                                        #     date_col_nums <-
                                                                        #         sapply(all_sheets[[i]][1, 2:ncol(all_sheets[[i]])], function(x) {
                                                                        #             # Attempt to convert to numeric, NA if not possible
                                                                        #             as.numeric(as.character(x))
                                                                        #         })
                                                                        #     
                                                                        #     
                                                                        #     # Filter out NAs in case there are non-convertible values (optional, based on your data)
                                                                        #     date_col_nums <- na.omit(date_col_nums)
                                                                        #     
                                                                        #     # Convert the numeric values to dates
                                                                        #     date_col_names <- as.Date(date_col_nums, origin = "1899-12-30")
                                                                        #     
                                                                        #     # Format dates as mm/dd/yyyy or your preferred format
                                                                        #     formatted_dates <- format(date_col_names, "%m/%d/%Y")
                                                                        #     
                                                                        #     # Prepend "Facility" to the list of formatted date column names
                                                                        #     all_col_names <- c("Facility", formatted_dates)
                                                                        #     
                                                                        #     # Step 3: Set the formatted dates along with "Facility" as column names
                                                                        #     colnames(all_sheets[[i]]) <- all_col_names
                                                                        #     
                                                                        #     # Step 4: Now, remove the first row as it's been used to set column names
                                                                        #     all_sheets[[i]] <- all_sheets[[i]][-1,]
                                                                        # }

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

# Convert Values to numeric prior to export
all_sheets_long$Value <- as.numeric(all_sheets_long$Value)

#####################################################
# Export data in long format
####################################################
# Export final dataframe into excel
export_file_path <-
    file.path(paths$exports_path, "Weekly_Backlog_long_data.xlsx")
# writexl::write_xlsx(x = all_sheets_long, path = export_file_path)



# Aggregate data by Category and Date
aggregated_data <- all_sheets_long %>%
    group_by(Date, Category) %>%
    summarise(Total_Value = sum(as.numeric(Value), na.rm = TRUE)) %>%
    ungroup()

# Convert Date to character for easier formatting
# aggregated_data$Date <- as.character(aggregated_data$Date)


# Export aggregated data to Excel
export_file_path_aggregated <-
    file.path(paths$exports_path, "Weekly_Backlog_Date_aggregated_data.xlsx")
# writexl::write_xlsx(x = aggregated_data, path = export_file_path_aggregated)


#####################################################
# End of Weekly Backlog processing
####################################################






####################################################
# Aging Report
###################################################

# Apply the extract_date function created above, to extract dates from imported file
aging_dates <- lapply(aging_report_files_list, extract_date)

# Use which.max to find the index of the latest date
latest_aging_index <- which.max(aging_dates)



# Import the latest Aging Report
latest_aging <- aging_report_files_list[latest_aging_index]

# Read the data from all sheets from the latest file:
aging_all_sheets <-
    setNames(lapply(excel_sheets(latest_aging), function(sheet)
        read_excel(
            latest_aging,
            sheet = sheet,
            skip = 3,
            col_names = FALSE
        )),
        excel_sheets(latest_aging))

for (j in 1:length(aging_all_sheets)) {
    aging_all_sheets[[j]] <- aging_all_sheets[[j]] %>%
        rename(Facility = 1) %>%
        mutate(across(-Facility))
    
}


# Now, read the headers. We want only the headers to use later for naming the columns
aging_headers <-
    setNames(lapply(excel_sheets(latest_aging), function(sheet)
        read_excel(
            latest_aging,
            sheet = sheet,
            n_max = 3,
            col_names = FALSE
        )),
        excel_sheets(latest_aging))


for (m in 1:length(aging_headers)) {
    # Replace the first NA with "Facility"
    aging_headers[[m]][1, 1] <- "Facility"
    aging_headers[[m]][3, 1] <- "Facility"
    # Initialize recent_value with NA to handle the case where the first column is NA
    recent_value <- NA
    
    # Loop through each column starting from the second
    for (n in 2:ncol(aging_headers[[m]])) {
        # Check if the value is NA and needs to be replaced
        if (is.na(aging_headers[[m]][1, n])) {
            # If recent_value is not NA (meaning we've encountered a numeric value before), use it
            if (!is.na(recent_value)) {
                aging_headers[[m]][1, n] <- recent_value
            }
        } else {
            # If it's not an NA, update recent_value
            recent_value <- aging_headers[[m]][1, n]
        }
        
        # Additional code: Copy "Total" from row 2 to row 3 if present
        if (aging_headers[[m]][2, n] == "Total") {
            aging_headers[[m]][3, n] <- "Total"
        }
    }
}





# Function to combine three elements into a single string separated by underscores

for (a in 1:length(aging_headers)) {
    for (b in 1:ncol(aging_headers[[a]])) {
        new_header <-
            paste(aging_headers[[a]][1,],
                  aging_headers[[a]][2,],
                  aging_headers[[a]][3,],
                  sep = "_")
        names(aging_headers[[a]]) <- new_header
        
    }
    aging_headers[[a]] <- aging_headers[[a]][FALSE, ]
}





# Merge dataframes in aging_all_sheets with the headers in aging_headers
for (f in seq_along(aging_all_sheets)) {
    # Extract the column names from the corresponding dataframe in aging_headers
    new_headers <- colnames(aging_headers[[f]])
    # Replace the column names in the dataframe from aging_all_sheets with new_headers
    colnames(aging_all_sheets[[f]]) <- new_headers
}




# Use lapply with an anonymous function to process each dataframe
cleaned_aging_all_sheets <- lapply(aging_all_sheets, function(df) {
    # Identify the row with the word "Total"
    total_row_index <-
        which(df$Facility_Facility_Facility == "Total")
    
    # Assuming you want to keep columns that are non-numeric or don't have a 0 in the 'Total' row
    cols_to_keep <- sapply(df, function(column) {
        is.numeric(column) && column[total_row_index] != 0
    })
    
    # Adding TRUE for the 'Name' column or any other column you want to always keep
    cols_to_keep[1] <-
        TRUE  # Assuming the first column is always 'Name' or similar
    
    # Keep only the desired columns
    df <- df[, cols_to_keep]
    
    return(df)
})





########################################################
# Pivot all dataframes in the aging_all_sheets list, from wide into long format
######################################################
# Cast all columns into character type so we can pivot them
for (j in 1:length(cleaned_aging_all_sheets)) {
    cleaned_aging_all_sheets[[j]] <- cleaned_aging_all_sheets[[j]] %>%
        rename(Facility = 1) %>%
        mutate(across(-Facility))
}


# Initialize aging_all_sheets_long tibble
aging_long <- tibble()
# Extract the sheet names from the file
names(aging_all_sheets) <- excel_sheets(latest_aging)

for (k in names(cleaned_aging_all_sheets)) {
    # Pivoting the dataframe to long format
    pivoted_aging <- cleaned_aging_all_sheets[[k]] %>%
        mutate(across(everything(), as.character)) %>% # Convert all columns to character to avoid data type issues
        pivot_longer(
            cols = -Facility, # Selects all columns except `Facility`
            names_to = "Week", # New column for the dates
            values_to = "Value" # New column for the values
        ) %>%
        mutate(Category = k) # Add a new column with the name of the sheet
    
    # Append the pivoted dataframe to the accumulating dataframe
    aging_long <- bind_rows(aging_long, pivoted_aging)
}




# Remove "Total" rows from long formatted dataframe
aging_long <- aging_long %>%
    filter(!str_detect(Facility, "Total"))



# Separate the 'Week' column into three new columns
aging_long <- aging_long %>%
    separate(
        col = Week,
        into = c("Report_Date", "Period", "Range"),
        sep = "_",
        remove = TRUE
    ) # Set to TRUE if you want to remove the original column


# Remove "Total" rows from long formatted dataframe
aging_long <- aging_long %>%
    filter(!str_detect(Period, "Total"))


# Convert 'Date' column from m/d/Y format to Date format
# Convert Excel serial date numbers to Date format in R
aging_long$Report_Date <-
    as.Date(as.numeric(aging_long$Report_Date), origin = "1899-12-30")

# Convert Values to numeric prior to export
aging_long$Value <- as.numeric(aging_long$Value)
#####################################################
# Export data in long format
####################################################
# Export final dataframe into excel
export_file_path <- file.path(paths$exports_path, "aging_long.xlsx")
# writexl::write_xlsx(x = aging_long, path = export_file_path)



# Aggregate data by Category and Date
aggregated_aging <- aging_long %>%
    group_by(Report_Date, Category) %>%
    summarise(Total_Value = sum(as.numeric(Value), na.rm = TRUE)) %>%
    ungroup()

# Convert Date to character for easier formatting
# aggregated_data$Date <- as.character(aggregated_data$Date)

# Export aggregated data to Excel
export_aggregated_aging_path <-
    file.path(paths$exports_path, "aging_date_aggregated_data.xlsx")
# writexl::write_xlsx(x = aggregated_aging, path = export_aggregated_aging_path)

setwd(project_dir)

print("End of program")
#######################################
#############     EOF       ###########
#######################################

# save(list = ls(.GlobalEnv), file = "assets.RData")
