#Load Basic Libraries
rm(list = ls())
setwd('E:\\Automation\\ICICI')

library(data.table)
library(dplyr)
library(readxl)
library(openxlsx)

today <- Sys.Date()
ydate <- Sys.Date()-1

#Load Dump
dump_path <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",ydate,".csv",sep='')
dump <- fread(dump_path)
dump <- dump %>% filter(name == 'ICICI Bank' & status %in% c("CC-021","CC-022","CC-05","CC-051","CC-06","CC-07","CC-110","CC-200","CC-210","CC-230","CC-240","CC-250","CC-270","CC-275","CC-295","CC-310","CC-340","CC-390","CC-395","CC-396","CC-397","CC-400","CC-410","CC-421","CC-423","CC-430","CC-440","CC-490","CC-510","CC-540","CC-680","CC-690","CC-710","CC-88","CC-89","CC-90","CC-99","CC-990")) #Filter Only MoneyTap


#ICICI_Mapping
appopscode <- read_excel('ICICI_LenderMapping.xlsx',sheet = 'App Ops Status')
mapping1 <- read_excel('ICICI_LenderMapping.xlsx',sheet = 'Sheet1')
mapping2 <- read_excel('ICICI_LenderMapping.xlsx',sheet = 'Sheet2')


#Load ICICI Input File

ICICI_Data <- read_excel("ICICI_input.xlsx",sheet = "Data")
ICICI_Return <- read_excel("ICICI_input.xlsx",sheet = "Return")

ICICI_Data$`MOB NO` <- as.character(ICICI_Data$`MOB NO`)
ICICI_Return$`MOBILE NO` <-as.character(ICICI_Return$`MOBILE NO`)
dump$phone_home <- as.character(dump$phone_home)


df_data <- left_join(ICICI_Data, dump[,c('phone_home',
                                  'offer_application_number',
                                  'appops_status_code',
                                  'status')],
                by = c('MOB NO' = 'phone_home'))

df_Return <- left_join(ICICI_Return, dump[,c('phone_home',
                                  'offer_application_number',
                                  'appops_status_code',
                                  'status')],
                by = c('MOBILE NO' = 'phone_home'))

df_data <- left_join(df_data, appopscode, by = c('appops_status_code'='App Ops Status Code'))
df_data <- left_join(df_data, mapping1, by = c('PICKUP STATUS'='appstatus'))

df_data <- df_data %>% distinct(`MOB NO`, .keep_all = T)

df_Return <- left_join(df_Return, appopscode, by = c('appops_status_code'='App Ops Status Code'))
df_Return <- left_join(df_Return, mapping2, by = c('REMARKS'='appstatus'))

df_Return <- df_Return %>% distinct(`MOBILE NO`, .keep_all = T)


df_data <- df_data %>% select("CITY NAME","CUSTOMER NAME","MOB NO","offer_application_number","App Ops Status","AppOpsCode_New","PICKUP STATUS","REMARKS","LOGIN STATUS","DECLINE REASON","LOGIN DATE","xx1","xx2","Reject Reasons")
df_Return <- df_Return %>% select("CITY NAME","CUSTOMER NAME","MOBILE NO","offer_application_number","App Ops Status","AppOpsCode_New","PICKUP STATUS","REMARKS","xx1","xx2","Reject Reasons")

df_data$`Reject Reasons` <- paste(df_data$`PICKUP STATUS`,"/",df_data$REMARKS,"/",df_data$`LOGIN STATUS`,"/",df_data$`DECLINE REASON`,df_data$`LOGIN DATE`)
df_Return$`Reject Reasons` <- paste(df_Return$`PICKUP STATUS`,"/",df_Return$REMARKS)



wb <- createWorkbook()
addWorksheet(wb,"DATA")
addWorksheet(wb,"Return")
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
setColWidths(wb,"DATA",cols = 1:ncol(df_data),widths = 15)
setColWidths(wb,"Return",cols = 1:ncol(df_Return),widths = 15)

writeData(wb,"DATA",df_data,borders = "all",headerStyle = hs1)
writeData(wb,"Return",df_Return,borders = "all",headerStyle = hs1)

path <- paste('./Output/ICICI_Remarks_',Sys.Date(),'.xlsx',sep='')
saveWorkbook(wb,path,overwrite = T)
openXL(path)







