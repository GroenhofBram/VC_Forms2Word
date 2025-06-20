### PACKAGES ###
library(readxl)
library(tidyverse)
library(officer)
### PACKAGES ###

### FUNCTIONS ###
# Read in data
get_data <- function(filepath){
  data_df <- read_excel(filepath)
  return(data_df)
}

# Reformat data
reformat_data <- function(unformatted_df){
  formatted_df <- unformatted_df %>%
    rename("VC lid" = "Geef uw naam") %>%
    mutate(across(where(is.character), ~replace_na(., "geen opmerkingen")))
  
  return(formatted_df)
}

# Create .docx tables
create_docx_tables <- function(data_df){
  
  # Loop through each unique prefix (CG)
  for (prefix in prefixes) {
    # Create a new Word document for the current CG prefix
    doc <- read_docx()
    
    # Get all columns that start with the current prefix
    current_cg_columns <- cg_columns[grepl(paste0("^", prefix),
                                           cg_columns)]
    
    curr_CG_statement = paste0("Nu Bezig Met : ", prefix)
    print(curr_CG_statement)
    
    
    # Loop through each CG column to create a table
    for (cg_col in current_cg_columns) {
      # Create a data frame for the current CG column
      current_df <- data_df %>% select(`VC lid`, !!sym(cg_col))
      
      if (nrow(current_df) > 0) {  # Ensure there is data to print
        
        # Extract simplified column name for the table
        simplified_col_name <- extract_column_name(cg_col)
        colnames(current_df)[2] <- simplified_col_name  # Change the second column name
        
        # Create the document
        doc <- doc %>%
          body_add_par(cg_col, style = "Normal") %>%
          body_add_par(paste("VC", current_date), style = "Normal") %>% # Add VC and the current date
          body_add_table(value = current_df, style = "table_template", alignment = "left") %>%
          body_add_break()  
      }
    }
    
    export_tables(prefix, doc)
  }  
}

# Exports the tables
# Called in create_docx_tables
export_tables <- function(prefix, doc){
  # cleanthe filename for the CG Word document
  sanitized_prefix <- sanitize_filename(prefix)
  
  # Define the filename for the CG Word document
  filename <- paste0(current_date, "_FB Gebundeld_", sanitized_prefix, ".docx")
  
  # Save the document
  print(doc, target = filename)
}

# Function to clean filename
# Called in export_tables
sanitize_filename <- function(name) {
  # Replace any character that is not alphanumeric or underscore with an underscore
  gsub("[^[:alnum:]_]", "_", name)
}

# Function to extract the relevant part of the column name
# Called in extract_column_name
extract_column_name <- function(col_name) {
  # Extract components split by space
  parts <- strsplit(col_name, " ")[[1]]
  
  # Look for the last capitalized word
  for (i in rev(seq_along(parts))) {
    if (grepl("[A-Z]", parts[i])) {
      # Found the last capitalized part
      return(paste(parts[i:length(parts)], collapse = " "))  # Combine from this part to the end
    }
  }
  
  return(col_name)  # Fallback to original if no capitalized word found
}


### RUNNING ###
current_date <- format(Sys.Date(), "%Y-%m-%d")
filepath <- "Feedbackformulier MBO2F vergadering 5 juni 2025.xlsx" # TODO: User can upload file
data_df <- get_data(filepath)
data_df <- reformat_data(data_df)
cg_columns <- names(data_df)[grepl("^CG[0-9]", names(data_df))]  # Start with "CG"
prefixes <- unique(sub("(CG[0-9]+).*", "\\1", cg_columns))
create_docx_tables(data_df)
