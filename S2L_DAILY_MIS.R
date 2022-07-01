rm(list=ls())
setwd('E:\\Automation\\MIS File')
Sys.getenv("R_ZIPCMD", "zip")

library(data.table)
library(dplyr)
library(lubridate)
library(data.table)
library(readxl)
library(openxlsx)
library(stringr)
library(stringi)

ydate <- Sys.Date()

month<- format(Sys.Date(),"%m")
#month <-as.numeric(month)-1
year<- format(Sys.Date(),"%Y")
thismonth <-paste(year,"-",month,"-01",sep = "")

thismonth <-as.Date(thismonth)-1

bank_fdate <- Sys.Date()-1
#bank_fdate <- Sys.Date()-2
referral_date <- Sys.Date()-1
#referral_date <- Sys.Date()-2
referral_date_7 <- Sys.Date()-3

referral_date <-as.character(referral_date)
bank_fdate <-as.Date(bank_fdate)
referral_date_7 <-as.character(referral_date_7)
thismonth <-as.character(thismonth)

y_date <- Sys.Date()- 1
dump_name <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",y_date,".csv",sep ='')
dump <- fread(dump_name)

dump$bank_feedback_date <- as.Date(dump$bank_feedback_date,"%Y-%m-%d")

Ref_snet <- dump %>% filter(date_of_referral == referral_date & appops_status_code == "270" )%>% select(name,appops_status_code)

Not_received <- dump %>% filter(date_of_referral > thismonth & date_of_referral < referral_date_7 & appops_status_code == "270" & lender_notes =="" & applied_date >thismonth)

Not_received <- Not_received %>% select(status,name,appops_status_code)

Ref_received <- dump %>% filter(bank_feedback_date == bank_fdate & appops_status_code > "269"  & !lender_notes =="")

Ref_received <- Ref_received %>% select(status,name,appops_status_code)


wb <- createWorkbook()
addWorksheet(wb,"Ref_sent")
addWorksheet(wb,"Not_received")
addWorksheet(wb,"Ref_received")
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
setColWidths(wb,"Ref_sent",cols = 1:ncol(Ref_snet),widths = 10)
writeData(wb,"Ref_sent",Ref_snet,borders = "all",headerStyle = hs1)

setColWidths(wb,"Not_received",cols = 1:ncol(Not_received),widths = 10)
writeData(wb,"Not_received",Not_received,borders = "all",headerStyle = hs1)

setColWidths(wb,"Ref_received",cols = 1:ncol(Ref_received),widths = 10)
writeData(wb,"Ref_received",Ref_received,borders = "all",headerStyle = hs1)

path <- paste('Mis File_',Sys.Date(),'.xlsx',sep='')
saveWorkbook(wb,path,overwrite = T)
openXL(path)


