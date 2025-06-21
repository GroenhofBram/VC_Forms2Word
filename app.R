library(shiny)
library(readxl)
library(tidyverse)
library(officer)

# --- FUNCTIONS ---

# Reformat data
reformat_data <- function(unformatted_df){
  formatted_df <- unformatted_df %>%
    rename("VC lid" = "Geef uw naam") %>%
    mutate(across(where(is.character), ~replace_na(., "geen opmerkingen")))
  
  return(formatted_df)
}

# Function to clean filename
sanitize_filename <- function(name) {
  gsub("[^[:alnum:]_]", "_", name)
}

# Function to extract relevant part of column name
extract_column_name <- function(col_name) {
  parts <- strsplit(col_name, " ")[[1]]
  for (i in rev(seq_along(parts))) {
    if (grepl("[A-Z]", parts[i])) {
      return(paste(parts[i:length(parts)], collapse = " "))
    }
  }
  return(col_name)
}

# --- SHINY APP ---

ui <- fluidPage(
  titlePanel("Feedback Word Exporter"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Upload Excel File", accept = ".xlsx"),
      actionButton("process", "Generate Word Documents")
    ),
    mainPanel(
      verbatimTextOutput("status")
    )
  )
)

server <- function(input, output) {
  observeEvent(input$process, {
    req(input$file)
    
    tryCatch({
      # Read and process data
      data_df <- read_excel(input$file$datapath)
      data_df <- reformat_data(data_df)
      
      # Get CG-related columns
      cg_columns <- names(data_df)[grepl("^CG[0-9]", names(data_df))]
      prefixes <- unique(sub("(CG[0-9]+).*", "\\1", cg_columns))
      current_date <- format(Sys.Date(), "%Y-%m-%d")
      
      # Process each prefix
      for (prefix in prefixes) {
        doc <- read_docx()
        current_cg_columns <- cg_columns[grepl(paste0("^", prefix), cg_columns)]
        
        for (cg_col in current_cg_columns) {
          current_df <- data_df %>% select(`VC lid`, !!sym(cg_col))
          
          if (nrow(current_df) > 0) {
            simplified_col_name <- extract_column_name(cg_col)
            colnames(current_df)[2] <- simplified_col_name
            
            doc <- doc %>%
              body_add_par(cg_col, style = "Normal") %>%
              body_add_par(paste("VC", current_date), style = "Normal") %>%
              body_add_table(value = current_df, style = "table_template", alignment = "left") %>%
              body_add_break()
          }
        }
        
        sanitized_prefix <- sanitize_filename(prefix)
        filename <- paste0(current_date, "_FB_Gebundeld_", sanitized_prefix, ".docx")
        print(doc, target = filename)
      }
      
      output$status <- renderText({
        "✅ Word documents generated and saved to the working directory."
      })
      
    }, error = function(e) {
      output$status <- renderText({
        paste("❌ Error:", e$message)
      })
    })
  })
}

shinyApp(ui, server)
