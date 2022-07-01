#Load Basic Libraries
rm(list = ls())
setwd('E:\\Automation\\MoneyTap')

library(data.table)
library(dplyr)
library(readxl)
library(openxlsx)

today <- Sys.Date()
ydate <- Sys.Date()-1

#Load Dump
dump_path <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",ydate,".csv",sep='')
dump <- fread(dump_path)
dump <- dump %>% filter(name == 'Money Tap') #Filter Only MoneyTap

#MoneyTap_Mapping
appopscode <- read_excel('MoneyTap_LenderMapping.xlsx',sheet = 'App Ops Status')
mapping <- read_excel('MoneyTap_LenderMapping.xlsx',sheet = 'Sheet1')


#Load MoneyTap Input File
money_in <- fread(paste("credit_mantri_",today,'_csv.csv',sep=""))

#Join Phone Number vs Creditmantri Application Number
money_in$phone <- as.character(money_in$phone)
dump$phone_home <- as.character(dump$phone_home)

df <- left_join(money_in, dump[,c('phone_home',
                                  'offer_application_number',
                                  'appops_status_code',
                                  'status')],
                by = c('phone' = 'phone_home'))

df <- left_join(df, appopscode, by = c('appops_status_code'='App Ops Status Code'))
df <- left_join(df, mapping, by = 'appstatus')

#Select Required Coloumns
df %>% names()

col_names <- c('ref_id','fullname','phone','loanamount','status',
               'offer_application_number','App Ops Status','AppOpsCode_New',
               'appstatus','X1','X2')

df_new <- df %>% select(col_names)
df_new$X3 <- paste(df_new$ref_id,"/",df_new$appstatus)
names(df_new)[names(df_new) == 'App Ops Status'] <- 'Pre'
df_new1<-df_new[ !is.na(df_new$offer_application_number)  , ] %>% distinct(phone, .keep_all = T)

#Write Data in Excel
wb <- createWorkbook()
addWorksheet(wb,"MoneyTap")
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
setColWidths(wb,"MoneyTap",cols = 1:ncol(df_new1),widths = 15)
writeData(wb,"MoneyTap",df_new1,borders = "all",headerStyle = hs1)
path <- paste('./Output/MoneyTap_Remarks_',Sys.Date(),'.xlsx',sep='')
saveWorkbook(wb,path,overwrite = T)
openXL(path)

dim(money_in)
