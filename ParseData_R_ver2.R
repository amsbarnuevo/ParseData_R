library(tidyverse)
library(dplyr)
library(readxl)
library(xlsx)
library(writexl)
library(svDialogs)
library(data.table)


setwd("D:/ALLYSA FILE/2023/DMU Files/DMU Projects/Parse Data")

#reading data file
get_file <- dlgInput("Enter a filename:", Sys.info()[" "])$res
df_raw <- read_xlsx(paste(get_file,"xlsx",sep ="."))

# get column name/s from user
get_column <- dlgInput("Please enter one or more column name separated by comma.")$res

# split column name/s
list_colnames <- unlist(strsplit(gsub(" ", "", get_column, fixed = TRUE), ","))
num_cols <- as.numeric(length(list_colnames))

#Create new dataframe containing the ID and the column to be parse
id_cols <- names(df_raw)[1]
new_df <- subset(df_raw, select = c(id_cols))

 

SplitCols <- function(x){

  #Get column name
  split_column <- as.character(x)
  
  #Bind column name and split column
  parse_df <- cbind(new_df,column_split = df_raw[[x]])
  
  #Change Column Name
  setnames(parse_df, "column_split", split_column)
  

  #Split column of comma-separated numbers into multiple columns based on value
  parse_df <- parse_df %>% 
    mutate(column_name = strsplit(df_raw[[x]], ",")) %>% 
    unnest(column_name) %>% 
    mutate(value = 1) %>% 
    pivot_wider(names_from = column_name, values_fill = 0) %>%
    as.data.frame()


  #Replace column value with the name of its respective column
  with_value <- which(parse_df==1,arr.ind=TRUE)
  parse_df[with_value] <- names(parse_df)[with_value[,"col"]]

  #Replace "0" with blank
  parse_df[parse_df == 0] <- ""
  
  #drop column ID
  parse_df = select(parse_df, -1)

return(data.frame(parse_df))

}


if(num_cols == 1) {
  parse_df <- SplitCols(list_colnames)
  
  # Write the dataframe to Excel in the specified output folder
  output_filename <- paste0(list_colnames,"_parse.xlsx")

  
}else{
  split_df <-sapply(list_colnames, SplitCols)
  
  #Split a list into separate data frame in R
  split_df <- purrr::map(split_df, tibble::as_tibble)

  #Then send this to the global environment
  list2env(split_df, envir = .GlobalEnv)
  
  # create list of dataframe
  df_list <- mget(list_colnames)
  
  #merge dataframe
  parse_df <- bind_cols(new_df,df_list)
  parse_df <- as.data.frame(parse_df)
  
  #assign output filename
  output_filename <- paste0("multicolumn_parse.xlsx")
 
}

# Write the dataframe to Excel in the specified output folder
write.xlsx(parse_df, file = output_filename, row.names = FALSE)



#edited version 2
