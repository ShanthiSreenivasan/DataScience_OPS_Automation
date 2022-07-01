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

scuf_06_200_230_dump <- fread('./Input/Shriram PL 06 200 230.csv')



scuf_APTG <- scuf_06_200_230_dump %>% filter(grepl('200|230',product_status,ignore.case=TRUE), 
                                             grepl('AP | TG',ap_tg_split,ignore.case=TRUE)) %>% 
  mutate(Base_Name = `sku_slug` , 'Base_Type' = `product_status`) %>% 
  select(user_id, lead_id, phone_home, Base_Name, Base_Type)#first_profile_month, login_date, Language

write.csv(scuf_APTG, file = "C:\\R\\Ref_Dialler OPS Automation\\SHRIRAM_APTG_200_230.csv")

#scuf_ROI <- scuf_base[rev(order(scuf_base$first_profile_month)), ]

scuf_ROI <- scuf_06_200_230_dump %>% filter(grepl('200|230',product_status,ignore.case=TRUE), 
                                             grepl('Rest of India',ap_tg_split,ignore.case=TRUE)) %>% 
  mutate(Base_Name = `sku_slug`, 'Base_Type' = `product_status`) %>% 
  select(user_id, lead_id, phone_home, Base_Name, Base_Type)#first_profile_month, login_date, Language


write.csv(scuf_ROI, file = "C:\\R\\Ref_Dialler OPS Automation\\SHRIRAM_ROI_200_230.csv")

