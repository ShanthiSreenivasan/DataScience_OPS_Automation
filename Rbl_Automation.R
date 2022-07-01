#Load Basic Libraries
rm(list = ls())
setwd('E:\\Automation\\RBL')

library(data.table)
library(dplyr)
library(readxl)
library(openxlsx)

today <- Sys.Date()
ydate <- Sys.Date()-1

#Load Dump
dump_path <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",ydate,".csv",sep='')
dump <- fread(dump_path)
dump <- dump %>% filter(name == 'RBL')
Fdump <- dump %>% filter(status %in% c('CC-99','CC-310','CC-88','CC-05','CC-110','CC-440','CC-410','CC-990','CC-510','CC-275','CC-210','CC-340','CC-051','CC-540','CC-240','CC-490','CC-390','CC-391','CC-422','CC-423','CC-395','CC-396','CC-330','CC-421','CC-397','CC-89','CC-710','CC-07','CC-430','CC-295','CC-400','CC-90','CC-270','CC-520','CC-021','CC-06')) #Filter Only RBL


appopscode <- read_excel('Appops RBL Code.xlsx',sheet = 'App Ops Status')
mapping <- read_excel('Appops RBL Code.xlsx',sheet = 'Sheet1')
mapping2 <- read_excel('Appops RBL Code.xlsx',sheet = 'Sheet2')

Rbl_in <- read_excel('RBL_Input.xlsx',sheet = 'Appointment Fix')

Rbl_in1 <- Rbl_in %>% filter(`OPS Status 1` %in% c('Declined','Approved','Incurable','Picked But Not Login','WIP','Under Curing'))



Rbl_in1$`Mobile No` <- as.character(Rbl_in1$`Mobile No`)
Fdump$phone_home <- as.character(Fdump$phone_home)


df <- left_join(Rbl_in1, dump[,c('phone_home',
                                  'offer_application_number',
                                  'appops_status_code',
                                  'status')],
                by = c('Mobile No' = 'phone_home'))


df <- left_join(df, appopscode, by = c('appops_status_code'='App Ops Status Code'))

df <- left_join(df, mapping, by = 'OPS Status 1')

df<-df[ !is.na(df$offer_application_number)  , ] %>% distinct(`Mobile No`, .keep_all = T)

df$Comments=paste(df$`App Reference Number`,"/",df$`Field Status`,"/",df$`Application`,"/",df$`OPS Status 1`,"/",df$`Approved Date`,"/",df$`Declined Reasons`)

col_names <- c('App Reference Number','Customer Name','Mobile No','offer_application_number',
               'Pre','POST','Field Status','OPS Status 1','Approved Date','Declined Reasons','Comments')

df_new <- df %>% select(col_names)

















