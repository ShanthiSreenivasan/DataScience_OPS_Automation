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

ref_pre07_dump <- fread('./Input/Referrals Consolidated Dialer Dumps.csv')

ref_dialer <- ref_pre07_dump %>% filter(!grepl(07,product_status,ignore.case=TRUE), !lenders_presented %in% c(NA, ''), !customer_type == 'Red')

ref_dialer <-  ref_dialer %>% 
  mutate(
    Language = case_when(
      grepl('100|50|51|52|53|54', pincode, ignore.case = T) ~ 'Telugu',
      grepl('56|57|58|59', pincode, ignore.case = T) ~ 'Kannada',
      grepl('60|61|62|63|64|65', pincode, ignore.case = T) ~ 'Tamil',
      grepl('66|67|68|69', pincode, ignore.case = T) ~ 'Malayalam',
    )
  )

ref_dialer$Language[is.na(ref_dialer$Language)] <- 'Hindi'

ref_dialer <- ref_dialer[rev(order(ref_dialer$first_profile_month)), ]

ref_dialer <-  ref_dialer %>% 
  mutate(
    Lender = case_when(
      grepl('Early Salary|EarlySalary', lenders_presented, ignore.case = T) ~ 'Early Salary',
      grepl('CASHE', lenders_presented, ignore.case = T) ~ 'CASHE',
      grepl('MONEYVIEW', lenders_presented, ignore.case = T) ~ 'MONEYVIEW',
      grepl('KREDITBEE', lenders_presented, ignore.case = T) ~ 'KREDITBEE',
      grepl('MYSHUBLIFE', lenders_presented, ignore.case = T) ~ 'MYSHUBLIFE',
      grepl('PAYSENSE', lenders_presented, ignore.case = T) ~ 'PAYSENSE',
      grepl('RBL', lenders_presented, ignore.case = T) ~ 'RBL',
      grepl('IDFC', lenders_presented, ignore.case = T) ~ 'IDFC',
      grepl('AXIS', lenders_presented, ignore.case = T) ~ 'AXIS',
      grepl('SCB', lenders_presented, ignore.case = T) ~ 'SCB',
      grepl('CITI', lenders_presented, ignore.case = T) ~ 'CITI',
      grepl('SBI', lenders_presented, ignore.case = T) ~ 'SBI',
      grepl('YESBANK', lenders_presented, ignore.case = T) ~ 'YESBANK',
      grepl('HDFC', lenders_presented, ignore.case = T) ~ 'HDFC',
      grepl('LENDINGKART', lenders_presented, ignore.case = T) ~ 'LENDINGKART',
      grepl('FLEXILOAN', lenders_presented, ignore.case = T) ~ 'FLEXILOAN',
      grepl('IIFL', lenders_presented, ignore.case = T) ~ 'IIFL',
      grepl('INDIFI', lenders_presented, ignore.case = T) ~ 'INDIFI',
      grepl('PROTIUM', lenders_presented, ignore.case = T) ~ 'PROTIUM',
      grepl('AYE', lenders_presented, ignore.case = T) ~ 'AYE',
      grepl('SMECORNER', lenders_presented, ignore.case = T) ~ 'SMECORNER',
      grepl('LNT', lenders_presented, ignore.case = T) ~ 'LNT',
      grepl('MUTHOOT', lenders_presented, ignore.case = T) ~ 'MUTHOOT',
      grepl('FAIRCENT', lenders_presented, ignore.case = T) ~ 'FAIRCENT',
      grepl('FINZY', lenders_presented, ignore.case = T) ~ 'FINZY',
      grepl('SHRIRAM', lenders_presented, ignore.case = T) ~ 'SHRIRAM',
      grepl('PNB', lenders_presented, ignore.case = T) ~ 'PNB',
      grepl('INDIASHELTER', lenders_presented, ignore.case = T) ~ 'INDIASHELTER',
      grepl('CANARA', lenders_presented, ignore.case = T) ~ 'CANARA'
    )
  )

ref_dialer<-ref_dialer %>% mutate('Base_Name' = paste0(`Lender`, ' ', `product`), 'Base_Type' = `product_status`)

write.csv(ref_dialer, file = "C:\\R\\Ref_Dialler OPS Automation\\ref_dialer-07.csv")


ref_dialer_CC_Hindi<-ref_dialer %>% filter(Language == 'Hindi', grepl('CC' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_CC_Hindi, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_CC_HINDI-", Sys.Date(), '.csv'))

ref_dialer_CC_Tamil<-ref_dialer %>% filter(Language == 'Tamil', grepl('CC' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_CC_Tamil, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_CC_TAMIL-", Sys.Date(), '.csv'))

ref_dialer_CC_Telugu<-ref_dialer %>% filter(Language == 'Telugu', grepl('CC' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_CC_Telugu, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_CC_TELUGU-", Sys.Date(), '.csv'))

ref_dialer_CC_Mal<-ref_dialer %>% filter(Language == 'Malayalam', grepl('CC' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_CC_Mal, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_CC_MALAYALAM-", Sys.Date(), '.csv'))

ref_dialer_CC_kan<-ref_dialer %>% filter(Language == 'Kannada', grepl('CC' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_CC_kan, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_CC_Kannada-", Sys.Date(), '.csv'))

###############For STBL base

ref_dialer_STBL_Hindi<-ref_dialer %>% filter(Language == 'Hindi', grepl('STBL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_STBL_Hindi, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_STBL_HINDI-", Sys.Date(), '.csv'))

ref_dialer_STBL_Tamil<-ref_dialer %>% filter(Language == 'Tamil', grepl('STBL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_STBL_Tamil, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_STBL_TAMIL-", Sys.Date(), '.csv'))

ref_dialer_STBL_Telugu<-ref_dialer %>% filter(Language == 'Telugu', grepl('STBL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_STBL_Telugu, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_STBL_TELUGU-", Sys.Date(), '.csv'))

ref_dialer_STBL_Mal<-ref_dialer %>% filter(Language == 'Malayalam', grepl('STBL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_STBL_Mal, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_STBL_MALAYALAM-", Sys.Date(), '.csv'))

ref_dialer_STBL_kan<-ref_dialer %>% filter(Language == 'Kannada', grepl('STBL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_STBL_kan, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_STBL_Kannada-", Sys.Date(), '.csv'))

###############For PL base

ref_dialer_PL_Hindi<-ref_dialer %>% filter(Language == 'Hindi', grepl('PL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_PL_Hindi, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_PL_HINDI-", Sys.Date(), '.csv'))

ref_dialer_PL_Tamil<-ref_dialer %>% filter(Language == 'Tamil', grepl('PL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_PL_Tamil, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_PL_TAMIL-", Sys.Date(), '.csv'))

ref_dialer_PL_Telugu<-ref_dialer %>% filter(Language == 'Telugu', grepl('PL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_PL_Telugu, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_PL_TELUGU-", Sys.Date(), '.csv'))

ref_dialer_PL_Mal<-ref_dialer %>% filter(Language == 'Malayalam', grepl('PL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_PL_Mal, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_PL_MALAYALAM-", Sys.Date(), '.csv'))

ref_dialer_PL_kan<-ref_dialer %>% filter(Language == 'Kannada', grepl('PL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_PL_kan, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_PL_Kannada-", Sys.Date(), '.csv'))


###############For SBPL base

ref_dialer_SBPL_Hindi<-ref_dialer %>% filter(Language == 'Hindi', grepl('SBPL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_SBPL_Hindi, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_SBPL_HINDI-", Sys.Date(), '.csv'))

ref_dialer_SBPL_Tamil<-ref_dialer %>% filter(Language == 'Tamil', grepl('SBPL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_SBPL_Tamil, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_SBPL_TAMIL-", Sys.Date(), '.csv'))

ref_dialer_SBPL_Telugu<-ref_dialer %>% filter(Language == 'Telugu', grepl('SBPL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_SBPL_Telugu, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_SBPL_TELUGU-", Sys.Date(), '.csv'))

ref_dialer_SBPL_Mal<-ref_dialer %>% filter(Language == 'Malayalam', grepl('SBPL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_SBPL_Mal, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_SBPL_MALAYALAM-", Sys.Date(), '.csv'))

ref_dialer_SBPL_kan<-ref_dialer %>% filter(Language == 'Kannada', grepl('SBPL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_SBPL_kan, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_SBPL_Kannada-", Sys.Date(), '.csv'))

###############For HL base

ref_dialer_HL_Hindi<-ref_dialer %>% filter(Language == 'Hindi', grepl('HL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_HL_Hindi, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_HL_HINDI-", Sys.Date(), '.csv'))

ref_dialer_HL_Tamil<-ref_dialer %>% filter(Language == 'Tamil', grepl('HL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_HL_Tamil, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_HL_TAMIL-", Sys.Date(), '.csv'))

ref_dialer_HL_Telugu<-ref_dialer %>% filter(Language == 'Telugu', grepl('HL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_HL_Telugu, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_HL_TELUGU-", Sys.Date(), '.csv'))

ref_dialer_HL_Mal<-ref_dialer %>% filter(Language == 'Malayalam', grepl('HL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_HL_Mal, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_HL_MALAYALAM-", Sys.Date(), '.csv'))

ref_dialer_HL_kan<-ref_dialer %>% filter(Language == 'Kannada', grepl('HL' ,Base_Type,ignore.case=TRUE)) %>% select(user_id, lead_id, phone_home, Base_Name, Base_Type, first_profile_month, login_date)
write.csv(ref_dialer_HL_kan, file = paste0("C:\\R\\Ref_Dialler OPS Automation\\Allset_Referrals_HL_Kannada-", Sys.Date(), '.csv'))




