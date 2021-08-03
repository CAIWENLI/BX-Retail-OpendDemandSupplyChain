rm(list=ls())
knitr::opts_chunk$set(echo = TRUE)

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

open_demand_retail <- read_xlsx("C:/Users/lisal/OneDrive - bookxchange.com/Retail Reporting/Raw Data/Open_Demand_Retail/Open_Demand_Retail.xlsx")

open_demand_retail$ISBN <- as.character(open_demand_retail$ISBN)

tbl_vars(open_demand_retail)

open_demand_retail <- open_demand_retail %>% 
  mutate(DUP = paste(ISBN, WAREHOUSE, `DUE DATE`)) %>% 
  arrange(ISBN, `GUIDE NAME`) %>% 
  distinct(DUP, .keep_all = TRUE)
open_demand_retail$DUP <- NULL
tbl_vars(open_demand_retail)
colnames(open_demand_retail) <- c("ISBN", "SUB_ISBN","QUANTITY", "BB_PRICE", "WAREHOUSE", "DUE_DATE", "GUIDE_NAME", "DEMAND_DATE")

library(odbc)
library(DBI)

con.microsoft.sql <- DBI::dbConnect(odbc::odbc(),
                                    Driver   = "SQL Server",
                                    Server   = "52.86.56.66",
                                    Database = "PROCUREMENTDB",
                                    UID      = "LisaLi",
                                    PWD      = "t4vUByNaANWqszXP",
                                    Port     =  1433)

dbWriteTable(con.microsoft.sql, DBI::SQL("Sourcing.OpenDemand"), as.data.frame(open_demand_retail), overwrite = TRUE)
