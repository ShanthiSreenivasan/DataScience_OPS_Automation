rm(list = ls())

library(magrittr)
library(tibble)
library(dplyr)
library(purrr)
library(tidyr)
library(tidyverse)

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
library(purrr)
library(janitor)

setwd("C:\\R\\Lender Feedback")

getwd()

thisdate<-format(Sys.Date(),'%Y-%m-%d')

source('.\\Function file.R')

if(!dir.exists(paste("./Output/",thisdate,sep="")))
{
  dir.create(paste("./Output/",thisdate,sep=""))
} 

lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\CashE.xlsx",sheet = 'consoildated_feedback_cases')
appos_map_CashE <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Cashe')
appos_dump <- read.xlsx("C:\\R\\Lender Feedback\\Input\\appopsdump.xlsx",sheet = 'Sheet1') %>% select(phone_home,offer_application_number,status,name,appops_status_code)
########################CashE LenderFeedBack
CashE_appos_dump <- appos_dump %>% filter(name %in% c('CashE'))

#feedback_file = appos_dump %>% left_join(appos_dump$offer_application_number, by = c('phone_home'))

lender_MIS$offer_application_number<-CashE_appos_dump$offer_application_number[match(lender_MIS$phone_home, CashE_appos_dump$phone_home)]

lender_MIS$appops_status<-CashE_appos_dump$appops_status_code[match(lender_MIS$phone_home, CashE_appos_dump$phone_home)]

lender_MIS <-lender_MIS %>%
  mutate(New_appops_status = case_when(
    !is.na(loan_amount) ~ "990",
    appops_status <=280 ~ "300",
    appops_status >280 & appops_status <=400  ~ "400",
    appops_status >400 & appops_status <=480  ~ "480",
    appops_status >480 & appops_status <=490  ~ "490",
    appops_status >490 & appops_status <=500  ~ "500",
    appops_status >500 & appops_status <=590  ~ "590",
    appops_status ==990 ~ "990"
    ))

lender_MIS$New_appops_description<-appos_map_CashE$Status_Description[match(lender_MIS$New_appops_status, appos_map_CashE$New.Status)]


desc_lender_MIS <- lender_MIS %>% group_by(New_appops_description) %>% dplyr::summarise(Total = n_distinct(phone_home, na.rm = TRUE)) %>% adorn_totals("row")


CashE_upload<- lender_MIS %>% filter(New_appops_description %in% c("Initial FB - Contact successful", "Docs stage - Rejected")) %>% select(offer_application_number,New_appops_description) %>% 
  mutate(`Bank_Feedback_Date`=Sys.Date()-1,`Appointment_Date`="Nil",`Notes`="API",`Offer_Reference_Number`=" ",`Loan_Sanctioned_Disbursed_Amount`=" ",`Booking_Date`=" ",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

list_of_datasets <- list("FeedBack_update" = lender_MIS, "FeedBack_summary" = desc_lender_MIS)
write.xlsx(list_of_datasets, file = "C:\\R\\Lender Feedback\\CashE_feedback.xlsx")

write.xlsx(CashE_upload, file = "C:\\R\\Lender Feedback\\CashE_upload.xlsx")

names(KB_lender_MIS)

########################KB LenderFeedBack

KB_lender_MIS <- read.xlsx("C:\\R\\Lender Feedback\\Input\\KB.xlsx",sheet = 'Sheet1') %>% select(uId,State,user_subState,latest_loan_state,first_loan_gmv,mobile)
appos_map_KB <- read.xlsx("C:\\R\\Lender Feedback\\Revised Referrals Mapping.xlsx",sheet = 'Krazy Bee')

KB_appos_dump <- appos_dump %>% filter(name %in% c('Kredit Bee')) %>% select(phone_home,offer_application_number,status,name,appops_status_code)

KB_lender_MIS$offer_application_number<-KB_appos_dump$offer_application_number[match(KB_lender_MIS$mobile, KB_appos_dump$phone_home)]

KB_lender_MIS$appops_status<-KB_appos_dump$appops_status_code[match(KB_lender_MIS$mobile, KB_appos_dump$phone_home)]

KB_lender_MIS <-KB_lender_MIS %>%
  mutate(New_appops_status = case_when(
    !is.na(first_loan_gmv) ~ "990",
    appops_status >280 & appops_status <=380  ~ "480",
    appops_status >380 & appops_status <=500  ~ "490",
    appops_status >500 & appops_status <=590  ~ "590",
    appops_status ==990 ~ "990"
  ))


KB_lender_MIS$New_appops_description<-appos_map_KB$Status_Description[match(KB_lender_MIS$New_appops_status, appos_map_KB$New.Status)]


KB_desc_lender_MIS <- KB_lender_MIS %>% group_by(New_appops_description) %>% dplyr::summarise(Total = n_distinct(uId, na.rm = TRUE)) %>% adorn_totals("row")


KB_upload<- KB_lender_MIS %>% filter(New_appops_description %in% c("Initial FB - Contact successful", "Docs stage - Rejected")) %>% select(offer_application_number,New_appops_description) %>% 
  mutate(`Bank_Feedback_Date`=Sys.Date()-1,`Appointment_Date`="Nil",`Notes`="API",`Offer_Reference_Number`=" ",`Loan_Sanctioned_Disbursed_Amount`=" ",`Booking_Date`=" ",`Rejection_Tag`="Nil",`Rejection_Category`="Nil")

list_of_datasets <- list("FeedBack_update" = KB_lender_MIS, "FeedBack_summary" = KB_desc_lender_MIS)
write.xlsx(list_of_datasets, file = "C:\\R\\Lender Feedback\\KB_feedback.xlsx")

write.xlsx(KB_upload, file = "C:\\R\\Lender Feedback\\KB_upload.xlsx")

