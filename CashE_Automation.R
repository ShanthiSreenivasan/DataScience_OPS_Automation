#Remove existing List
rm(list=ls())
setwd('E:\\Automation\\CashE')


#Load Libraries
library(data.table)
library(readxl)
library(openxlsx)
library(dplyr)


#Read dump
y_date <- Sys.Date()-1
dump_name <- paste('E:\\Automation\\AppOpsDump\\appopsdump_',y_date,'.csv',sep ='')
dump <- fread(dump_name)
dump <- dump %>% filter(name == 'CashE')

#Read Input file
cashe_input <- read_excel('CashE_input_file.xlsx')
cashe_input <- cashe_input %>% select(customer_name,
                                      mobile_no,
                                      customer_status_name)

#Read Mapping File
cashe_mapping <- read_excel('./Source/CashE_Mapping.xlsx',sheet = 'CashE_Mapping')
AppOpsMapping <- read_excel('./Source/CashE_Mapping.xlsx',sheet = 'AppOpsMapping')

#Map Phone with application Number and Pre-existing Remarks
dump$phone_home <- as.character(dump$phone_home)
cashe_input$mobile_no <- as.character(cashe_input$mobile_no)

df_cashe <- left_join(cashe_input,
                      dump[,c('phone_home','offer_application_number',
                              'appops_status_code')],
                      by = c('mobile_no' = 'phone_home'))


df_cashe$appops_status_code <- as.character(df_cashe$appops_status_code)
AppOpsMapping$`App Ops Status Code` <- as.character(AppOpsMapping$`App Ops Status Code`)

df_cashe <- left_join(df_cashe,
                      AppOpsMapping,
                      by = c('appops_status_code' = 'App Ops Status Code'))

names(df_cashe)[names(df_cashe) == 'App Ops Status'] <- 'Pre'


#Mapping Post Remarks
df <- left_join(df_cashe,
                cashe_mapping,
                by = c('customer_status_name' = 'Sub Status'))

names(df)[names(df) == 'CM Status'] <- 'Post'



cashe_names <- c('customer_name','mobile_no','offer_application_number',
                 'Pre','Post','customer_status_name','Rejection_Tag',
                 'Rejection_Category')


df <- df %>% select(cashe_names)


#Create Output file
wb <- createWorkbook()
addWorksheet(wb,"CashE")
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
setColWidths(wb,"CashE",cols = 1:ncol(df),widths = 15)
writeData(wb,"CashE",df,borders = "all",headerStyle = hs1)
path <- paste('./Output/Cashe_Remarks_',Sys.Date(),'.xlsx',sep = '')
saveWorkbook(wb,path,overwrite = T)
openXL(path)







