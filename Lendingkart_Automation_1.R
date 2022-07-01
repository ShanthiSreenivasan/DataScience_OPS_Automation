#Load Basic Libraries
rm(list = ls())
setwd('E:\\Automation\\LendingKart')

library(data.table)
library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)
ydate <- Sys.Date()-1
today <- format(Sys.Date(),"%d-%m-%Y")

#Load Dump
dump_path <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",ydate,".csv",sep='')
dump <- fread(dump_path)
dump <- dump %>% filter(name == 'Lending Kart')

path <-paste(".//Input//Credit Mantri ",today,".xlsx",sep = "" )
Lending <- read_excel(path,sheet = "Credit Mantri (Last 60 days)",na = c('',NA))

Lending_Status <- Lending %>% filter(`Lead Status` != "Rejected")

Lending_Sub <- Lending %>% filter(`Lead Status` == "Rejected")


appopscode <- read_excel('Appops Moneyview Code.xlsx',sheet = 'App Ops Status')
mapping_Status <- read_excel('Appops Moneyview Code.xlsx',sheet = 'Sheet1')
mapping_Sub <- read_excel('Appops Moneyview Code.xlsx',sheet = 'Sheet2')

#Lead Status

Lending_Status$`Primary Phone` <- as.character(Lending_Status$`Primary Phone`)
dump$phone_home <- as.character(dump$phone_home)

df_Status <- left_join(Lending_Status, dump[,c('phone_home',
                                  'offer_application_number',
                                  'appops_status_code',
                                  'status')],
                by = c('Primary Phone' = 'phone_home'))

df_Status <- left_join(df_Status, appopscode, by = c('appops_status_code'='App Ops Status Code'))
df_Status <- left_join(df_Status, mapping_Status, by = "Lead Status")

#LEAD SUB

Lending_Sub$`Primary Phone` <- as.character(Lending_Sub$`Primary Phone`)
dump$phone_home <- as.character(dump$phone_home)

df_Sub <- left_join(Lending_Sub, dump[,c('phone_home',
                                               'offer_application_number',
                                               'appops_status_code',
                                               'status')],
                       by = c('Primary Phone' = 'phone_home'))

df_Sub <- left_join(df_Sub, appopscode, by = c('appops_status_code'='App Ops Status Code'))
df_Sub <- left_join(df_Sub, mapping_Sub, by = "Lead Sub Status")



Lending_file <- bind_rows(df_Status,df_Sub)

Lending_file$Comments <-paste(Lending_file$`Leads Number`,"/",Lending_file$`Application ID`,"/",Lending_file$`Lead Status`,"/",Lending_file$`Lead Sub Status`,"/",Lending_file$Comments)


Lending_file <- Lending_file %>% select("Created Date","Primary Phone",offer_application_number,Pre,post,
                    "Lead Status","Lead Sub Status",
                    Comments,XX1,XX2)

Lending_file<-Lending_file[ !is.na(Lending_file$offer_application_number)  , ] %>% distinct(Mobile_Number, .keep_all = T)


wb <-createWorkbook()
addWorksheet(wb,"Leads")
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
setColWidths(wb,"Leads",cols = 1:ncol(Lending_file),widths = 15)
writeData(wb,"Leads",Lending_file,borders = "all",headerStyle = hs1)
path <- paste('./Output/LendingKart_',Sys.Date(),'.xlsx',sep='')

saveWorkbook(wb,path,overwrite = T)
openXL(path)






