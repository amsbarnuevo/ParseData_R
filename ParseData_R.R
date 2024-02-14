library(tidyverse)
library(dplyr)
library(readxl)
library(xlsx)
library(writexl)
library(svDialogs)


setwd("D:/ALLYSA FILE/2023/DMU Files/DMU Projects/Parse Data")

#reading data file
get_file <- dlgInput("Enter a filename:", Sys.info()[" "])$res
df <- read_xlsx(paste(get_file,"xlsx",sep ="."))

#Define column to be parse
get_column <- dlgInput("Enter a column name:", Sys.info()[" "])$res
#column_data <- unique(unlist(strsplit(as.character(df[[get_column]]), ",")))

#Create new dataframe containing the ID and the column to be parse
id_cols <- names(df)[1]
new_df = subset(df, select = c(id_cols,get_column))

#Split column of comma-separated numbers into multiple columns based on value
parse_df <- new_df %>% 
  mutate(column_name = strsplit(df[[get_column]], ",")) %>% 
  unnest(column_name) %>% 
  mutate(value = 1) %>% 
  pivot_wider(names_from = column_name, values_fill = 0) %>%
  as.data.frame()

#Replace column value with the name of its respective column
with_value <- which(parse_df==1,arr.ind=TRUE)
parse_df[with_value] <- names(parse_df)[with_value[,"col"]]

#Replace "0" with blank
parse_df[parse_df == 0] <- ""

# Write the dataframe to Excel in the specified output folder
output_filename <- paste0(get_column,"_parse.xlsx")
write.xlsx(parse_df, file = output_filename, row.names = FALSE)


