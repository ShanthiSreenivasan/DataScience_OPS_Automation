rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(stringr)
library(lubridate)
library(data.table)
library(DBI)
library(RPostgreSQL)
library(RMySQL)
library(logging)
#library(mailR)
library(xtable)
library(yaml)
library(openxlsx)
library(bit64)
library(ggrepel)
library(lookup)

setwd("C:\\R\\Ref_Dialler OPS Automation")

getwd()

thisdate<-format(Sys.Date(),'%Y-%m-%d')

source('.\\Function file.R')

if(!dir.exists(paste("./Output/",thisdate,sep="")))
{
  dir.create(paste("./Output/",thisdate,sep=""))
} 

LTD07_dump <- fread('./Input/Referrals Consolidated SBPL Dialer Dumps-LTD07.csv')

LTD07 <- LTD07_dump %>% filter(!grepl(07,product_status,ignore.case=TRUE), 
                               grepl('SHRIRAM',lenders_presented,ignore.case=TRUE), 
                               !customer_type == 'Red') %>% mutate(Base_Name = 'SHRIRAM', 'Base_Type' = `product_status`)

LTD07 <- LTD07[rev(order(LTD07$first_profile_month)), ]

LTD07 <-  LTD07 %>% 
  mutate(
    Language = case_when(
      grepl('100|50|51|52|53|54', pincode, ignore.case = T) ~ 'Telugu',
      grepl('56|57|58|59', pincode, ignore.case = T) ~ 'Kannada',
      grepl('60|61|62|63|64|65', pincode, ignore.case = T) ~ 'Tamil',
      grepl('66|67|68|69', pincode, ignore.case = T) ~ 'Malayalam',
    )
  )
LTD07$Language[is.na(LTD07$Language)] <- 'Hindi'


LTD07<-LTD07 %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date, Language)

write.csv(LTD07, file = "C:\\R\\Ref_Dialler OPS Automation\\SHRIRAM_APTG_06.csv")

