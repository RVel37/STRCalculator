# Helper functions to generate the excel formulae used in the calculator.


################## TABLE 1 EXPECTED FREQUENCY FORMULAE ####################

# Function to convert number to Excel column letter
num_to_col <- function(n) {
  letters <- LETTERS
  if (n <= 26) {
    return(letters[n])
  } else {
    first <- (n - 1) %/% 26
    second <- (n - 1) %% 26 + 1
    return(paste0(letters[first], letters[second]))
  }
}

# Generate formula list
generate_excel_formulas <- function(finalData, start_col = 3, start_row = 7) {
  formulas <- list()
  gene_names <- names(finalData)
  
  for (i in seq_along(gene_names)) {
    gene <- gene_names[i]
    df <- finalData[[gene]]
    
    # Range
    last_row <- nrow(df) + 1  # Data starts at A4 in Excel
    last_col_index <- ncol(df)  
    last_col_letter <- num_to_col(last_col_index)
    
    # Cell reference
    cell <- paste0("C", start_row + i - 1)
    range <- paste0("A4:", last_col_letter, last_row)
    
    # Formula
    formula <- paste0(
      '=IF(', cell, '=0,"",VLOOKUP(', cell, ',', gene, '!', range, ',',
      last_col_index, ',FALSE)*VLOOKUP(', cell, ',', gene, '!', range, ',',
      last_col_index, ',FALSE))'
    )
    formulas[[gene]] <- formula
  }
  
  return(formulas)
}

formula_list <- generate_excel_formulas(finalData)
cat(paste(unlist(formula_list), collapse = "\n"))


############ PULL MOST COMMON ALLELES #################

extract_top_alleles <- function(data_list) {
  lapply(data_list, function(df) {
    # Get highest & second highest frequency values
    highest <- df$`Highest frequency per allele`[df$Allele == "Highest"]
    
    # Find alleles that match highest
    top_allele <- df$Allele[df$`Highest frequency per allele` == highest & df$Allele != "Highest" & df$Allele != "Second Highest"]

      highest = top_allele
     
  })
}

top_alleles_list <- extract_top_alleles(finalData)
print(top_alleles_list)


####################### TABLE 2 FORMULAE ####################

formula_list_2 <- mapply(function(df, gene, i) {
  sheet_range_end <- nrow(df) + 1  # skip bottom 2 rows
  cell_row <- 7 + i - 1
  cell <- paste0("H", cell_row) 
  
  paste0('=IF(', cell, '<>"",VLOOKUP(', cell, ',', gene, '!A4:H', sheet_range_end, 
         ',8,FALSE),"N/A")')
},
finalData, names(finalData), seq_along(finalData), SIMPLIFY = TRUE)

cat(formula_list_2, sep = "\n")
