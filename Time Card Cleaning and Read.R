rm(list=ls())

library(RPostgreSQL)
library(dplyr)
library(dbplyr)
library(data.table)
library(lubridate)
library(reshape2)
library(stringr)
library("readxl")
library(writexl)
library(openxlsx)
library(tidyverse)
library(odbc)

time_card_file <- list.files("C:/Users/lisal/OneDrive - bookxchange.com/Retail Reporting/Raw Data/Time_Card_Raw", pattern="*.csv", full.names=TRUE)
time_card <- lapply(time_card_file, read.csv, header=T, sep=",", fileEncoding="UTF-8-BOM")
for (i in 1:length(time_card)){time_card[[i]]<-cbind(time_card[[i]],time_card_file[i])}
time_card_data <- do.call("rbind", time_card) 

time_card_data$`time_card_file[i]` <- gsub("^.*_", "", time_card_data$`time_card_file[i]`)
time_card_data$`time_card_file[i]` <- gsub(".csv", "", time_card_data$`time_card_file[i]`)

workers <- read.xlsx("C:/Users/lisal/OneDrive - bookxchange.com/Retail Reporting/Reports/Warehouse Analysis/Total Workers.xlsx")
ly_time_card <- read.csv("C:/Users/lisal/OneDrive - bookxchange.com/Retail Reporting/Reports/Warehouse Analysis/Time Card LY.csv")
colnames(ly_time_card) <- c("Position.ID", "Last.Name", "Supervisor", "Date", "Hours")

ly_time_card$Date <- format(as.Date(ly_time_card$Date, "%m/%d/%Y"), "%m/%d/%Y")

time_card_data <- time_card_data %>% 
  filter(Hours != 0) %>% 
  mutate(Date = sub(" .*", "", In.time)) %>% 
  left_join(workers, by = c("Last.Name")) %>% 
  filter(Warehouse %in% "YES") %>% 
  select(Position.ID, Last.Name, Supervisor, Date, Hours) 

time_card_data$Date <- format(as.Date(time_card_data$Date, "%m/%d/%Y"),"%m/%d/%Y")

time_card_all <- rbind(ly_time_card, time_card_data)

time_card_all <- time_card_all %>% 
  mutate(key = paste0(Position.ID, Date, Hours)) %>% 
  distinct(key,.keep_all = TRUE) %>% 
  select(Position.ID, Last.Name, Supervisor, Date, Hours) 

con.microsoft.sql <- DBI::dbConnect(odbc::odbc(),
                                    Driver   = "SQL Server",
                                    Server   = "52.86.56.66",
                                    Database = "PROCUREMENTDB",
                                    UID      = "LisaLi",
                                    PWD      = "t4vUByNaANWqszXP",
                                    Port     =  1433)

dbWriteTable(con.microsoft.sql, DBI::SQL("Warehouse.TimeCard"), as.data.frame(time_card_all), overwrite = TRUE)
