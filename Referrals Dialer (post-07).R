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

ref_post07_dump <- fread('./Input/Referrals Dialer Dumps(post-07).csv')

names(ref_post07)

ref_post07 <- ref_post07_dump %>% filter(grepl('CashE|Early Salary|KREDITBEE|EarlySalary|MONEYVIEW|Paysense',lender,ignore.case=TRUE), 
                                         !grepl('DNC|NI|NE|NI2|COVID REFUSE',crm_status_text,ignore.case=TRUE), 
                                         !latest_dispo %in% c('CD','DNC','NI'), !has_reject_status %in% c(1), 
                                         !grepl('021|022|03|05|06|88|89|98|210|310',product_status,ignore.case=TRUE), 
                                         attempts %in% c(0,1,2,3,4,5))



ref_post07 <-  ref_post07 %>% 
  mutate(
    Language = case_when(
      grepl('100|50|51|52|53|54', pincode, ignore.case = T) ~ 'Telugu',
      grepl('56|57|58|59', pincode, ignore.case = T) ~ 'Kannada',
      grepl('60|61|62|63|64|65', pincode, ignore.case = T) ~ 'Tamil',
      grepl('66|67|68|69', pincode, ignore.case = T) ~ 'Malayalam',
    )
  )

ref_post07$Language[is.na(ref_post07$Language)] <- 'Hindi'


ref_post07<-ref_post07 %>% mutate('Base_Name' = paste0(`lender`, ' ', `product`), 'Base_Type' = `appops_status_code`)


ref_post07_HINDI<-ref_post07 %>% filter(Language == 'Hindi') %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, profile_vintage, applied_date)
write.csv(ref_post07_HINDI, file = "C:\\R\\Ref_Dialler OPS Automation\\STBL_Docs_Hindi.csv")

ref_post07_Tamil<-ref_post07 %>% filter(Language == 'Tamil') %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, profile_vintage, applied_date)
write.csv(ref_post07_Tamil, file = "C:\\R\\Ref_Dialler OPS Automation\\STBL_Docs_Tamil.csv")

ref_post07_Telugu<-ref_post07 %>% filter(Language == 'Telugu') %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, profile_vintage, applied_date)
write.csv(ref_post07_Telugu, file = "C:\\R\\Ref_Dialler OPS Automation\\STBL_Docs_Telugu.csv")


ref_post07_Kannada<-ref_post07 %>% filter(Language == 'Kannada') %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, profile_vintage, applied_date)
write.csv(ref_post07_Kannada, file = "C:\\R\\Ref_Dialler OPS Automation\\STBL_Docs_Kannada.csv")

ref_post07_Malayalam<-ref_post07 %>% filter(Language == 'Malayalam') %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, profile_vintage, applied_date)
write.csv(ref_post07_Malayalam, file = "C:\\R\\Ref_Dialler OPS Automation\\STBL_Docs_Malayalam.csv")

