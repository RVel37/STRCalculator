library(readODS)
library(writexl)
library(tidyverse)

PATH <- "C:\\Users\\velis\\Downloads\\populationdata2.ods"

# starting data: allele frequency tables for 17 STR loci extracted from Excel file.
# Each sheet represents a distinct population group based on ethnicity/regional background.
# Data from King's College London. 
# (Note: The population group names below reflect the categories used in the source dataset.)

sheets <- list(
  "White" = read_ods(PATH, sheet = 2),
  "Black African & Caribbean" = read_ods(PATH, sheet = 3),
  "Indian" = read_ods(PATH, sheet = 4),
  "Chinese" = read_ods(PATH, sheet = 5),
  "African" = read_ods(PATH, sheet = 6),
  "African Caribbean" = read_ods(PATH, sheet = 7)
)

########### MOVE N TO TOP ##############

reordered <- lapply(sheets, function(df) {
  n_row <- tail(df, 1)
  rest <- head(df, -1)
  
  # Capitalize "n" to "N" in the Allele column only
  n_row$Allele <- ifelse(n_row$Allele == "n", "N", n_row$Allele)
  
  bind_rows(n_row, rest)
})


########## SORT INTO STR TABLES ###########

# pull frequencies for each STR from tables
# and pull non-NA allele numbers -> join to STR table

strDataReformat <- setNames(lapply(names(reordered[[1]])[-1], function(str_col) {
  
  # initialise empty table for STR with an allele column
  str_table <- tibble(Allele = reordered[[1]]$Allele)
  
  for (population in names(reordered)) {
    pop_data <- reordered[[population]]
    
    # Rename first column to "Allele" 
    colnames(pop_data)[1] <- "Allele"
    
    # Select Allele and current STR 
    pop_data <- pop_data %>%
      select(Allele, Frequency = all_of(str_col))
    
    # Extract N value from "N" row
    N_value <- as.numeric(
      pop_data %>% filter(Allele == "N") %>% pull(Frequency)
    )
    
    # Normalise values by N (except the N row itself)
    pop_data <- pop_data %>%
      mutate(Frequency = if_else(
        Allele != "N" & !is.na(Frequency),
        as.character(as.numeric(Frequency) / N_value),
        as.character(Frequency)
      ))
    
    # Join to main STR table
    str_table <- left_join(str_table, pop_data, by = "Allele")
    colnames(str_table)[ncol(str_table)] <- population
  }
  
  # Create character vector of column names
  pop_cols <- setdiff(names(str_table), "Allele")
  
  # Convert data to numeric values
  str_table[pop_cols] <- str_table[pop_cols] %>%
    mutate(across(everything(), ~ suppressWarnings(as.numeric(.))))
  
  # Get highest frequency per allele by applying pmax across population columns
  str_table <- str_table %>%
    mutate(`Highest frequency per allele` = pmax(
      !!!syms(pop_cols), na.rm = TRUE  
    )) %>%
    
    # Replace "N" row value in this column with "-"
    mutate(
      `Highest frequency per allele` = if_else(
        Allele == "N",
        "-", 
        as.character(`Highest frequency per allele`)
      ))
  
  setNames(list(str_table), str_col)
}),
names(reordered[[1]])[-1]
)

############## CLEAN UP #####################

# Convert nested list to list of dataframes
strDataFrames <- map(strDataReformat, ~ {
  df <- .x[[1]] %>%
    as.data.frame() %>%
    column_to_rownames(var = "Allele") 
  
  # Format "N" row
  if ("N" %in% rownames(df)) {
    df["N", ] <- map_chr(df["N", ], function(val) {
      if (!is.na(val) && !is.character(val) && !is.na(as.numeric(val))) {
        paste0("N=", val)
      } else if (!is.na(val) && is.character(val) && grepl("^[0-9]+$", val)) {
        paste0("N=", val)
      } else {
        val  # keep as-is 
      }
    })
  }

  # Extract top two frequencies 
  freq_vals <- df[["Highest frequency per allele"]]
  freq_vals_num <- as.numeric(freq_vals)
  top2_vals <- sort(freq_vals_num, decreasing = TRUE, na.last = NA)[1:2]
  
  # Create empty rows with correct columns
  top_rows <- df[0, , drop = FALSE]  # empty df with same structure
  top_rows["Highest", ] <- NA
  top_rows["Second Highest", ] <- NA
  top_rows["Highest", "Highest frequency per allele"] <- as.character(top2_vals[1])
  top_rows["Second Highest", "Highest frequency per allele"] <- as.character(top2_vals[2])
  
  # Combine
  bind_rows(df, top_rows)
})

# function to remove NAs
remove_all_na_rows <- function(df_list) {
  cleaned <- lapply(df_list, function(df) {
    df %>%
      mutate(
        `Highest frequency per allele` = as.numeric(`Highest frequency per allele`)
      ) %>%
      filter(!is.na(`Highest frequency per allele`) & `Highest frequency per allele` != 0)
  })
  return(cleaned)
}

# remove NAs
cleanedData <- remove_all_na_rows(strDataFrames)

# round to 4dp (requires changing vals back to numeric)
rounded <- cleanedData %>%
  lapply(function(df) {
    df %>%
      mutate(across(where(is.character), ~ as.numeric(.))) %>%
      mutate(across(where(is.numeric), ~ round(., 4)))  
  })

# Add rownames as a column
finalData <- rounded %>%
  map(~ .x %>% 
        tibble::rownames_to_column(var = "Allele"))

write_xlsx(finalData, path = "newData.xlsx")
