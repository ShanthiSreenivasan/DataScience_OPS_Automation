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
library(mailR)
library(xtable)
library(yaml)
library(openxlsx)
library(bit64)
library(ggrepel)
library(lookup)
library(fs)
library(htmlTable)
#install.packages("htmlTable")

#install.packages("ggrepel")
#install.packages("lookup")

setwd("C:\\R\\CIS_DIALER OPS Automation")

getwd()

thisdate<-format(Sys.Date(),'%Y-%m-%d')

source('.\\Function file.R')

if(!dir.exists(paste("./Output/",thisdate,sep="")))
{
  dir.create(paste("./Output/",thisdate,sep=""))
} 
#################################CIS NEW V0 DIALER BASE 

CIS_v0_dump <- fread('./Input/CIS_NEW_V0_DIALER_BASE.csv')
CIS_v0_dump$user_id <- as.numeric(CIS_v0_dump$user_id)
CIS_v0_dump$total_amount_overdue <- round(CIS_v0_dump$total_amount_overdue +1,0)

names(CIS_v0_dump)[4]<-c("Lead_id")
names(CIS_v0_dump)[17]<-c("No_of_Accounts")

CIS_v0_dump$Customer_Name <- str_c(CIS_v0_dump$first_name,'',CIS_v0_dump$last_name)
CIS_v0_dump %>% drop_na(latest_login)

CIS_v0_dump$latest_login <- convertToDate(CIS_v0_dump$latest_login)


CIS_v0_dump$latest_login <- as.Date(CIS_v0_dump$latest_login,             # Change class of date column
                      format = "%d-%m-%Y")
CIS_v0_dump <- CIS_v0_dump[rev(order(CIS_v0_dump$latest_login)), ]
# CIS_v0_dump$latest_login <- CIS_v0_dump[rev(order(as.Date(CIS_v0_dump$latest_login, format("%d/%m/%Y")))),]
# 
# CIS_v0_dump$desc_latest_login <- arrange(desc(dmy(CIS_v0_dump$latest_login)))

MIVR_tamil <- CIS_v0_dump %>% filter(!No_of_Accounts == 0, !latest_dispo %in% c('CD','DNC','NI','PTP'), Language=='Tamil', attempts %in% c(0,1,2), effective_attempts %in% c(0,1,2))


MIVR_tamil <- MIVR_tamil %>% select(user_id, Lead_id, phone_home, Customer_Name, No_of_Accounts) %>% slice(1:500)


write.csv(MIVR_tamil, file = "C:\\R\\CIS_DIALER OPS Automation\\V0_MIVR_TAMIL.csv")

MIVR_hindi <- CIS_v0_dump %>% filter(!No_of_Accounts == 0, !latest_dispo %in% c('CD','DNC','NI','PTP'), Language=='Hindi', attempts %in% c(0,1,2,3,4,5,6), effective_attempts %in% c(0,1,2,3,4)) %>% select(user_id, Lead_id, phone_home, Customer_Name, No_of_Accounts) %>% slice(1:6500)
write.csv(MIVR_hindi, file = "C:\\R\\CIS_DIALER OPS Automation\\V0_MIVR_HINDI.csv")

MIVR_telugu <- CIS_v0_dump %>% filter(!No_of_Accounts == 0, !latest_dispo %in% c('CD','DNC','NI','PTP'), Language=='Telgu', attempts %in% c(0,1,2), effective_attempts %in% c(0,1,2)) %>% select(user_id, Lead_id, phone_home, Customer_Name, No_of_Accounts) %>% slice(1:600)
write.csv(MIVR_telugu, file = "C:\\R\\CIS_DIALER OPS Automation\\V0_MIVR_TELUGU.csv")


#################################CHR BYS CONSOILATED DIALER DUMP   ( IVR )
CHR_BYS_dump <- fread('./Input/CHR BYS Consolidated Dialer Dump.csv')
CHR_BYS_dump$user_id <- as.numeric(CHR_BYS_dump$user_id)

IVR_tamil <- CHR_BYS_dump %>% filter(is.na(bhr_subscribed_flag) %in% c(0,NA), chr_subscribed_flag %in% c(0,NA), bys_subscribed_flag %in% c(0,NA), cis_subscribed_flag %in% c(0,NA), 
                                     !is.na(CHR.BYS.product_status) %in% c("CHRA-DX-XX-070","BYS-DX-XX-070","BYS-PP-XX-070","BYS-UD-XX-070","CHRA-DX-XX-100","BYS-DX-XX-100","CHR-DX-XX-100"), 
                                     !is.na(residential_pincode) %in% c(NA), 
                                     !employment_type %in% c('Student', "Working Executive", "NA",""), customer_type %in% c('Green', 'Red'), credit_score<725, calls == 0, 
                                     contacted==0 & !grepl('CD|DNC|NI|PTP|NE|NI2|COVID REFUSE',crm_status_text,ignore.case=TRUE))


IVR_tamil$latest_profile_date <- as.Date(IVR_tamil$latest_profile_date,             # Change class of date column
                                    format = "%d-%m-%Y")
IVR_tamil <- IVR_tamil[rev(order(IVR_tamil$latest_profile_date)), ]

IVR_tamil <-IVR_tamil %>% drop_na(residential_pincode, employment_type)


IVR_tamil <-  IVR_tamil %>% 
  mutate(
    Language = case_when(
      grepl('100|50|51|52|53|54', residential_pincode, ignore.case = T) ~ 'Telugu',
      grepl('56|57|58|59', residential_pincode, ignore.case = T) ~ 'Kannada',
      grepl('60|61|62|63|64|65', residential_pincode, ignore.case = T) ~ 'Tamil',
      grepl('66|67|68|69', residential_pincode, ignore.case = T) ~ 'Malayalam',
    )
  )

IVR_tamil$Language[is.na(IVR_tamil$Language)] <- 'Hindi'
IVR_tamil$Customer_Name <- str_c(IVR_tamil$first_name,'',IVR_tamil$last_name)

IVR_tamil <-  IVR_tamil %>% 
  mutate(
    Emp_type_flag = case_when(
      grepl('Self employed|Self employed professional|Self employed business|selfEmployedBusiness|Self-Employed|selfemployee|selfEmployedProfessional', employment_type, ignore.case = T) ~ '1',
      grepl('salaried|Salaried|Salaried Doctor|salariedDoctor', employment_type, ignore.case = T) ~ '2',
      
    )
  )

IVR_tamil$latest_profile_month <- format(as.Date(IVR_tamil$latest_profile_date, format="%d/%m/%Y"),"%m")
thisdate<-format(as.Date(Sys.Date()-1, format="%d/%m/%Y"),"%m")


IVR_tamil <-IVR_tamil %>%
  mutate(latest_profile_vintage_flag = case_when(
    latest_profile_month < thisdate ~ 'M0+',
    latest_profile_month == thisdate ~ 'M0',
  
  ))


IVR_tamil <-IVR_tamil %>%
  mutate(green_m0_flag = case_when(
    Emp_type_flag == 1 & customer_type == 'Green' ~ 'TRUE',
    Emp_type_flag == 1 & latest_profile_vintage_flag == 'M0' ~ 'TRUE'
    
  ))

IVR_tamil$green_m0_flag[is.na(IVR_tamil$green_m0_flag)] <- 'FALSE'



TCN_dump <- read.xlsx("C:\\R\\CIS_DIALER OPS Automation\\Input\\TCN_dump.xlsx",sheet = 'Sheet1')
#names(IVR_tamil)

IVR_tamil$ph_no<-TCN_dump$Phone.Number[match(IVR_tamil$phone_home, TCN_dump$Phone.Number)]
IVR_tamil$ph_no[is.na(IVR_tamil$ph_no)] <- 'FALSE'

# MIVR_Red <- CIS_v0_dump
# 
# MIVR_Red$ph_no<-TCN_dump$Phone.Number[match(MIVR_Red$phone_home, TCN_dump$Phone.Number)]


IVR_tamil <- IVR_tamil %>% filter(green_m0_flag %in% c('FALSE') & ph_no %in% c('FALSE'))


write.csv(IVR_tamil, file = "C:\\R\\CIS_DIALER OPS Automation\\CONSOILDATED_IVR_Tamil.csv")                                     
IVR_tamil_df<- IVR_tamil %>% filter(Language %in% c('Tamil')) %>% select(user_id, lead_id, phone_home, Customer_Name)
write.csv(IVR_tamil_df, file = "C:\\R\\CIS_DIALER OPS Automation\\IVR_Tamil.csv")                                     

IVR_Hindi<- IVR_tamil %>% filter(Language %in% c('Hindi')) %>% select(user_id, lead_id, phone_home, Customer_Name)
write.csv(IVR_Hindi, file = "C:\\R\\CIS_DIALER OPS Automation\\IVR_Hindi.csv")                                     

IVR_Telugu<- IVR_tamil %>% filter(Language %in% c('Telugu')) %>% select(user_id, lead_id, phone_home, Customer_Name)
write.csv(IVR_Telugu, file = "C:\\R\\CIS_DIALER OPS Automation\\IVR_Telugu.csv")                                     

#################################RED BASE \ CIS NEW V0 DIALER BASE


MIVR_Red <- CIS_v0_dump

MIVR_Red$ph_no<-TCN_dump$Phone.Number[match(MIVR_Red$phone_home, TCN_dump$Phone.Number)]
MIVR_Red<- MIVR_Red[-which(is.na(MIVR_Red$phone_home)), ]

# MIVR_Red<- MIVR_Red %>% 
#   filter(is.na(ph_no))
# 

MIVR_Red <-MIVR_Red %>%
  mutate(CIS_cases_flag = case_when(
    is.na(ph_no) ~ '0',
    !is.na(ph_no) ~ '1'
  ))

write.csv(MIVR_Red, file = "C:\\R\\CIS_DIALER OPS Automation\\CONSOILDATED_V0_RED_MIVR.csv")
MIVR_Red_Hindi<- MIVR_Red %>% filter(CIS_cases_flag == 0, No_of_Accounts == 0, !latest_dispo %in% c('CD','DNC','NI','PTP'), Language=='Hindi', attempts %in% c(0,1,2,3,4,5,6), effective_attempts %in% c(0,1,2,3,4)) %>% select(user_id, Lead_id, phone_home, Customer_Name) %>% slice(1:6500)
write.csv(MIVR_Red_Hindi, file = "C:\\R\\CIS_DIALER OPS Automation\\V0_RED_MIVR_Hindi.csv")

MIVR_Red_Tamil<- MIVR_Red %>% filter(CIS_cases_flag == 0, No_of_Accounts == 0, !latest_dispo %in% c('CD','DNC','NI','PTP'), Language=='Tamil', attempts %in% c(0,1,2,3,4,5,6), effective_attempts %in% c(0,1,2,3,4)) %>% select(user_id, Lead_id, phone_home, Customer_Name) %>% slice(1:500)
write.csv(MIVR_Red_Tamil, file = "C:\\R\\CIS_DIALER OPS Automation\\V0_RED_MIVR_Tamil.csv")

MIVR_Red_Telugu<- MIVR_Red %>% filter(CIS_cases_flag == 0, No_of_Accounts == 0, !latest_dispo %in% c('CD','DNC','NI','PTP'), Language=='Telgu', attempts %in% c(0,1,2,3,4,5,6), effective_attempts %in% c(0,1,2,3,4)) %>% select(user_id, Lead_id, phone_home, Customer_Name) %>% slice(1:600)
write.csv(MIVR_Red_Tamil, file = "C:\\R\\CIS_DIALER OPS Automation\\V0_RED_MIVR_Telugu.csv")

#################################BHR DIALER DUMP BASE

BHR_dump <- fread('./Input/BHR Dialer - dump.csv')
BHR_dump$user_id <- as.numeric(BHR_dump$user_id)
BHR_dump$phone_home <- suppressWarnings(as.numeric(BHR_dump$phone_home))

BHR_IVR <- BHR_dump %>% filter(is.na(bhr_subscription_flag) %in% c(0,NA), chr_subscription_flag %in% c(0,NA), bys_subscription_flag %in% c(0,NA), 
                               cis_subscription_flag %in% c(0,NA), !is.na(pincode) %in% c(NA), 
                                     !employment_type %in% c('Student', "Working Executive", "NA",""), customer_type %in% c('Green', 'Red'), credit_score<=700, calls == 0, 
                                     contacted==0 & !grepl('XX-070|XX-100',all_pos,ignore.case=TRUE), !grepl('CD|DNC|NI|PTP|NE|NI2|COVID REFUSE',crm_status_text,ignore.case=TRUE))


# BHR_IVR$latest_profile_date <- as.Date(BHR_IVR$latest_profile_date,             # Change class of date column
#                                          format = "%d-%m-%Y")
# BHR_IVR <- BHR_IVR[rev(order(BHR_IVR$latest_profile_date)), ]

#IVR_tamil <-IVR_tamil %>% drop_na(residential_pincode, employment_type)


BHR_IVR <-  BHR_IVR %>% 
  mutate(
    Language = case_when(
      grepl('100|50|51|52|53|54', pincode, ignore.case = T) ~ 'Telugu',
      grepl('56|57|58|59', pincode, ignore.case = T) ~ 'Kannada',
      grepl('60|61|62|63|64|65', pincode, ignore.case = T) ~ 'Tamil',
      grepl('66|67|68|69', pincode, ignore.case = T) ~ 'Malayalam',
    )
  )

BHR_IVR$Language[is.na(BHR_IVR$Language)] <- 'Hindi'
BHR_IVR$Customer_Name <- str_c(BHR_IVR$first_name,'',BHR_IVR$last_name)

MIVR_Red<-MIVR_Red
TCN_dump<-TCN_dump
BHR_IVR$ph_no<-TCN_dump$Phone.Number[match(BHR_IVR$phone_home, TCN_dump$Phone.Number)]

BHR_IVR <-BHR_IVR %>%
  mutate(TCN_cases_flag = case_when(
    is.na(ph_no) ~ '0',
    !is.na(ph_no) ~ '1'
  ))

BHR_IVR<- BHR_IVR %>% filter(TCN_cases_flag == 0)
#names(MIVR_Red)
BHR_IVR$phone_home <- as.numeric(BHR_IVR$phone_home)


BHR_IVR$dup_chk<-MIVR_Red$phone_home[match(as.numeric(BHR_IVR$phone_home), as.numeric(MIVR_Red$phone_home))]

BHR_IVR <-BHR_IVR %>%
  mutate(MIVR_cases_flag = case_when(
    is.na(dup_chk) ~ '0',
    !is.na(dup_chk) ~ '1'
  ))

#full_dup_chk <- union_all(MIVR_Red$phone_home, suppressWarnings(as.integer(IVR_tamil$phone_home))) %>% data.frame() %>% drop_na()
#names(full_dup_chk)
# BHR_IVR$phone_home<-as.numeric(BHR_IVR$phone_home)
# 
# 
#BHR_IVR$dup_chk2<-full_dup_chk$.[match(as.numeric(BHR_IVR$phone_home), as.numeric(full_dup_chk$.))]
# 
#BHR_IVR <-BHR_IVR %>%
#   mutate(mivr_cases_flag = case_when(
#     is.na(mivr_ph_no) ~ '0',
#     !is.na(mivr_ph_no) ~ '1'
#   ))
# 
BHR_IVR<- BHR_IVR %>% filter(MIVR_cases_flag == 0)

names(BHR_IVR)[2]<-c("Lead_id")
names(BHR_IVR)[24]<-c("No_of_Accounts")

#write.csv(full_dup_chk, file = "C:\\R\\CIS_DIALER OPS Automation\\CONSOILDATED_full_dup_chk_IVR.csv")

write.csv(BHR_IVR, file = "C:\\R\\CIS_DIALER OPS Automation\\CONSOILDATED_BHR_IVR.csv")


BHR_IVR_Hindi<- BHR_IVR %>% filter(MIVR_cases_flag == 0, No_of_Accounts == 0, Language=='Hindi') %>% select(user_id, Lead_id, phone_home, Customer_Name, No_of_Accounts) %>% slice(1:6500)
write.csv(BHR_IVR_Hindi, file = "C:\\R\\CIS_DIALER OPS Automation\\BHR_IVR_Hindi.csv")

BHR_IVR_Tamil<- BHR_IVR %>% filter(MIVR_cases_flag == 0, No_of_Accounts == 0, Language=='Tamil') %>% select(user_id, Lead_id, phone_home, Customer_Name, No_of_Accounts) %>% slice(1:500)
write.csv(BHR_IVR_Tamil, file = "C:\\R\\CIS_DIALER OPS Automation\\BHR_IVR_Tamil.csv")

BHR_IVR_Telugu<- BHR_IVR %>% filter(MIVR_cases_flag == 0, No_of_Accounts == 0, Language=='Telugu') %>% select(user_id, Lead_id, phone_home, Customer_Name, No_of_Accounts) %>% slice(1:600)
write.csv(BHR_IVR_Telugu, file = "C:\\R\\CIS_DIALER OPS Automation\\BHR_IVR_Telugu.csv")

############################CONSOLIDATED_CIS_M0_DIALER_DUMP_2PM

cis_m0_dump<-fread('./Input/CONSOLIDATED_CIS_M0_DIALER_DUMP.csv')

names(cis_m0_dump)[6]<-c("Lead_id")
cis_m0_dump$Customer_Name <- str_c(cis_m0_dump$first_name,'',cis_m0_dump$last_name)
cis_m0_dump <-  cis_m0_dump %>% 
  mutate(
    Language = case_when(
      grepl('100|50|51|52|53|54', zip, ignore.case = T) ~ 'Telugu',
      grepl('56|57|58|59', zip, ignore.case = T) ~ 'Kannada',
      grepl('60|61|62|63|64|65', zip, ignore.case = T) ~ 'Tamil',
      grepl('66|67|68|69', zip, ignore.case = T) ~ 'Malayalam',
    )
  )
cis_m0_dump$Language[is.na(cis_m0_dump$Language)] <- 'Hindi'

cis_m0_dump <- cis_m0_dump %>% filter(attempts %in% c(0,1,2) & effective_attempts%in% c(0))

write.csv(cis_m0_dump, file = "C:\\R\\CIS_DIALER OPS Automation\\CONSOILDATED_cis_m0_dialer base.csv")

cis_m0_Hindi<- cis_m0_dump %>% filter(Language=='Hindi') %>% select(user_id, Lead_id, phone_home, Customer_Name)# %>% slice(1:6500)
write.csv(cis_m0_Hindi, file = "C:\\R\\CIS_DIALER OPS Automation\\cis_m0_Hindi.csv")

cis_m0_Tamil<- cis_m0_dump %>% filter(Language=='Tamil') %>% select(user_id, Lead_id, phone_home, Customer_Name)# %>% slice(1:500)
write.csv(cis_m0_Tamil, file = "C:\\R\\CIS_DIALER OPS Automation\\cis_m0_Tamil.csv")

cis_m0_Telugu<- cis_m0_dump %>% filter(Language=='Telugu') %>% select(user_id, Lead_id, phone_home, Customer_Name)# %>% slice(1:600)
write.csv(cis_m0_Telugu, file = "C:\\R\\CIS_DIALER OPS Automation\\cis_m0_Telugu.csv")


############################CONSOLIDATED_CIS_REPEAT_USER_DIALER_DUMP

cis_repeat_dump<-fread('./Input/CONSOLIDATED_CIS_REPEAT_USER_DIALER_DUMP.csv')

names(cis_repeat_dump)[4]<-c("Lead_id")

cis_repeat_dump$phone_home <- suppressWarnings(as.numeric(cis_repeat_dump$phone_home))

cis_repeat_dump<- cis_repeat_dump[!is.na(cis_repeat_dump$phone_home),]


cis_repeat_dump <- cis_repeat_dump %>% filter(!grepl('XX-070|XX-100',product_status,ignore.case=TRUE),
                                              decile %in% c(0,1,2,3,4,5,6,7,8,9,10),
                                              !ever_050_6mon %in% c(1), nsaleable %in% c(1,2,3,4))


cis_repeat_dump$latest_login_date <- as.Date(cis_repeat_dump$latest_login_date,             # Change class of date column
                                          format = "%d-%m-%Y")
cis_repeat_dump <- cis_repeat_dump[rev(order(cis_repeat_dump$latest_login_date)), ]

write.csv(cis_repeat_dump, file = "C:\\R\\CIS_DIALER OPS Automation\\CONSOILDATED_cis_repeat_user.csv")

cis_repeat_dump_Hindi<- cis_repeat_dump %>% filter(language=='Hindi', attempts %in% c(0,1,2,3,4,5,6), effective_attempts %in% c(0,1,2,3,4)) %>% select(user_id, Lead_id, phone_home, customer_name)# %>% slice(1:6500)
write.csv(cis_repeat_dump_Hindi, file = "C:\\R\\CIS_DIALER OPS Automation\\IVR_Hindi_v2.csv")


############################CONSOLIDATED_CIS Repeat User Hourly Dialer Dump 4PM




cis_repeat_hr_dump<-fread('./Input/CONSOLIDATED_CIS_REPEAT_USER_DIALER_DUMP_HOURLY.csv')

names(cis_repeat_hr_dump)
names(cis_repeat_hr_dump)[2]<-c("Lead_id")

cis_repeat_hr_dump$phone_home <- suppressWarnings(as.numeric(cis_repeat_hr_dump$phone_home))

cis_repeat_hr_dump<- cis_repeat_hr_dump[!is.na(cis_repeat_hr_dump$phone_home),]

cis_repeat_hr_dump <-  cis_repeat_hr_dump %>% 
  mutate(
    Language = case_when(
      grepl('100|50|51|52|53|54', zip, ignore.case = T) ~ 'Telugu',
      grepl('56|57|58|59', zip, ignore.case = T) ~ 'Kannada',
      grepl('60|61|62|63|64|65', zip, ignore.case = T) ~ 'Tamil',
      grepl('66|67|68|69', zip, ignore.case = T) ~ 'Malayalam',
    )
  )
cis_repeat_hr_dump$Language[is.na(cis_repeat_hr_dump$Language)] <- 'Hindi'

cis_repeat_dump_Hindi<- cis_repeat_hr_dump %>% filter(language=='Hindi', attempts %in% c(0,1,2), effective_attempts %in% c(0)) %>% select(user_id, Lead_id, phone_home, customer_name)# %>% slice(1:6500)
write.csv(cis_repeat_dump_Hindi, file = "C:\\R\\CIS_DIALER OPS Automation\\SD_MIVR_Hindi_v2.csv")

cis_repeat_dump_Tamil<- cis_repeat_hr_dump %>% filter(language=='Tamil', attempts %in% c(0,1,2), effective_attempts %in% c(0)) %>% select(user_id, Lead_id, phone_home, customer_name)# %>% slice(1:6500)
write.csv(cis_repeat_dump_Tamil, file = "C:\\R\\CIS_DIALER OPS Automation\\SD_MIVR_Tamil_v2.csv")

cis_repeat_dump_Telugu<- cis_repeat_hr_dump %>% filter(language=='Telugu', attempts %in% c(0,1,2), effective_attempts %in% c(0)) %>% select(user_id, Lead_id, phone_home, customer_name)# %>% slice(1:6500)
write.csv(cis_repeat_dump_Telugu, file = "C:\\R\\CIS_DIALER OPS Automation\\SD_MIVR_Telugu_v2.csv")

############################CONSOLIDATED_CIS Repeat User Hourly Dialer Dump 2PM




cis_repeat_hr_dump_2pm<-fread('./Input/CONSOLIDATED_CIS_REPEAT_USER_DIALER_DUMP_HOURLY_2pm.csv')

names(cis_repeat_hr_dump_2pm)[2]<-c("Lead_id")

cis_repeat_hr_dump_2pm$phone_home <- suppressWarnings(as.numeric(cis_repeat_hr_dump_2pm$phone_home))

cis_repeat_hr_dump_2pm<- cis_repeat_hr_dump_2pm[!is.na(cis_repeat_hr_dump_2pm$phone_home),]

cis_repeat_hr_dump_2pm <-  cis_repeat_hr_dump_2pm %>% 
  mutate(
    Language = case_when(
      grepl('100|50|51|52|53|54', zip, ignore.case = T) ~ 'Telugu',
      grepl('56|57|58|59', zip, ignore.case = T) ~ 'Kannada',
      grepl('60|61|62|63|64|65', zip, ignore.case = T) ~ 'Tamil',
      grepl('66|67|68|69', zip, ignore.case = T) ~ 'Malayalam',
    )
  )
cis_repeat_hr_dump_2pm$Language[is.na(cis_repeat_hr_dump_2pm$Language)] <- 'Hindi'

cis_repeat_hr_dump_2pm_Hindi<- cis_repeat_hr_dump_2pm %>% filter(Language=='Hindi', attempts %in% c(0,1,2), effective_attempts %in% c(0)) %>% select(user_id, Lead_id, phone_home, customer_name)# %>% slice(1:6500)
write.csv(cis_repeat_dump_2pm_Hindi, file = "C:\\R\\CIS_DIALER OPS Automation\\SD_MIVR_Hindi_m0.csv")

cis_repeat_dump_2pm_Tamil<- cis_repeat_hr_dump_2pm %>% filter(Language=='Tamil', attempts %in% c(0,1,2), effective_attempts %in% c(0)) %>% select(user_id, Lead_id, phone_home, customer_name)# %>% slice(1:6500)
write.csv(cis_repeat_dump_2pm_Tamil, file = "C:\\R\\CIS_DIALER OPS Automation\\SD_MIVR_Tamil_m0.csv")

cis_repeat_dump_2pm_Telugu<- cis_repeat_hr_dump_2pm %>% filter(Language=='Telugu', attempts %in% c(0,1,2), effective_attempts %in% c(0)) %>% select(user_id, Lead_id, phone_home, customer_name)# %>% slice(1:6500)
write.csv(cis_repeat_dump_2pm_Telugu, file = "C:\\R\\CIS_DIALER OPS Automation\\SD_MIVR_Telugu_m0.csv")


############################CIS Subscribed Base Hourly




cis_sub<-fread('./Input/CIS Subscribed Base Hourly.csv')

names(cis_sub)

cis_sub$user_id <- suppressWarnings(as.numeric(cis_sub$user_id))
cis_sub$phone_home <- suppressWarnings(as.numeric(cis_sub$phone_home))


cis_sub = cis_sub[!duplicated(cis_sub$phone_home),]


cis_sub$subscription_date <- as.Date(cis_sub$subscription_date,             # Change class of date column
                                     format = "%d-%m-%Y","%d-%m-%Y")
cis_sub <- cis_sub[rev(order(cis_sub$subscription_date)), ]

tilldate<-format(as.Date(Sys.Date()-3, format="%d/%m/%Y"),"%d/%m/%Y")

cis_sub<- cis_sub %>% filter(!order_total==0, is_chr==0, subscription_date >= tilldate)

write.csv(cis_sub, file = "C:\\R\\CIS_DIALER OPS Automation\\CHR_South.csv")

#################################CIS NEW V0 DIALER BASE {CURRENT DATE}karthick Agency



v0_dump <- fread('./Input/CIS_NEW_V0_DIALER_BASE.csv')
v0_dump$user_id <- as.numeric(v0_dump$user_id)
names(v0_dump)
names(v0_dump)[4]<-c("Lead_id")
names(v0_dump)[17]<-c("No_of_Accounts")

v0_dump$Customer_Name <- str_c(v0_dump$first_name,'',v0_dump$last_name)
v0_dump %>% drop_na(latest_login)

v0_dump$latest_login <- convertToDate(v0_dump$latest_login)

v0_dump$latest_login <- as.Date(v0_dump$latest_login,             # Change class of date column
                                    format = "%d-%m-%Y")
v0_dump <- CIS_v0_dump[rev(order(v0_dump$latest_login)), ]

v0_tamil <- v0_dump %>% filter(!No_of_Accounts == 0, !latest_dispo %in% c('CD','DNC','NI','PTP'), Language=='Tamil', attempts %in% c(0,1,2), effective_attempts %in% c(0,1,2))


v0_tamil <- v0_tamil %>% select(user_id, Lead_id, phone_home, Customer_Name, No_of_Accounts)


write.csv(v0_tamil, file = "C:\\R\\CIS_DIALER OPS Automation\\V0_TAMIL.csv")

v0_hindi <- v0_dump %>% filter(!No_of_Accounts == 0, !latest_dispo %in% c('CD','DNC','NI','PTP'), Language=='Hindi', attempts %in% c(0,1,2,3,4,5,6), effective_attempts %in% c(0,1,2,3,4)) %>% select(user_id, Lead_id, phone_home, Customer_Name, No_of_Accounts) %>% slice(1:3500)
write.csv(MIVR_hindi, file = "C:\\R\\CIS_DIALER OPS Automation\\V0_HINDI.csv")

v0_telugu <- v0_dump %>% filter(!No_of_Accounts == 0, !latest_dispo %in% c('CD','DNC','NI','PTP'), Language=='Telgu', attempts %in% c(0,1,2), effective_attempts %in% c(0,1,2)) %>% select(user_id, Lead_id, phone_home, Customer_Name, No_of_Accounts)
write.csv(v0_telugu, file = "C:\\R\\CIS_DIALER OPS Automation\\V0_TELUGU.csv")

v0_kannada <- v0_dump %>% filter(!No_of_Accounts == 0, !latest_dispo %in% c('CD','DNC','NI','PTP'), Language=='Kannada', attempts %in% c(0,1,2), effective_attempts %in% c(0,1,2)) %>% select(user_id, Lead_id, phone_home, Customer_Name, No_of_Accounts)
write.csv(v0_kannada, file = "C:\\R\\CIS_DIALER OPS Automation\\V0_KANNADA.csv")

v0_malayalam <- v0_dump %>% filter(!No_of_Accounts == 0, !latest_dispo %in% c('CD','DNC','NI','PTP'), Language=='Malayalam', attempts %in% c(0,1,2), effective_attempts %in% c(0,1,2)) %>% select(user_id, Lead_id, phone_home, Customer_Name, No_of_Accounts)
write.csv(v0_kannada, file = "C:\\R\\CIS_DIALER OPS Automation\\V0_MALAYALAM.csv")



#################################COGENT PTP EVER PTP


ptp_dump <- fread('./Input/EVER PTP CASES.csv')

ptp_dump$phone_home <- as.numeric(ptp_dump$phone_home)

ptp_dump = ptp_dump[!duplicated(ptp_dump$phone_home),]

names(ptp_dump)[24]<-c("No_of_Accounts")


ptp_dump$Customer_Name <- str_c(ptp_dump$first_name,'',ptp_dump$last_name)

ptp_dump$latest_log_date <- as.Date(ptp_dump$latest_log_date,            
                                          format = "%d-%m-%Y")
ptp_dump <- ptp_dump[rev(order(ptp_dump$latest_log_date)), ]

ptp_dump$log_date <- as.Date(ptp_dump$log_date,            
                                    format = "%d-%m-%Y")
ptp_dump <- ptp_dump[rev(order(ptp_dump$log_date)), ]

tilldate<-format(as.Date(Sys.Date()-3, format="%d/%m/%Y"),"%d/%m/%Y")


ptp_dump <- ptp_dump %>% filter(chr_subscription_flag == 0, cis_subscription_flag ==0,
                                grepl('CG',latest_oic,ignore.case=TRUE), log_date >= tilldate) %>% select(user_id, lead_id, phone_home, Customer_Name, No_of_Accounts)

write.csv(ptp_dump, file = "C:\\R\\CIS_DIALER OPS Automation\\Cogent_ptp.csv")

###################################count

csvFileName = list.files("C:\\R\\CIS_DIALER OPS Automation", pattern=NULL, all.files=FALSE,
           full.names=FALSE)

csv.file <- list.files("C:\\R\\CIS_DIALER OPS Automation\\Output\\2022-06-29") # Directory with your .csv files
data.frame.output <- data.frame(Date = NA,
                                Base_Name = NA, 
                                Counts = NA) #The df to be written


MyF <- function(x){
  
  csv.read.file <- data.table::fread(
    paste("C:\\R\\CIS_DIALER OPS Automation\\Output", x, sep = "/")
  )
  
  number.of.rows <- nrow(csv.read.file)
  
  data.frame.output <<- add_row(data.frame.output, 'Date'=Sys.Date(),
                                Base_Name = str_remove_all(x,".csv"),
                                Counts = number.of.rows) %>% 
    filter(!is.na(Base_Name))
  
}

map(csv.file, MyF)

data.table::fwrite(data.frame.output, file = "CIS_Base_Count.csv")


###########send mail

sender <- "shanthi.s@creditmantri.com"  # Replace with a valid address
recipients <- c("rakshith.thangaraj@creditmantri.com")  # Replace with one or more valid addresses
today<-Sys.Date()
msg<-paste("Hi Team,      Please find CIS upload base summary for ",today,".",sep='')
#msg<- c(paste0("C:\\R\\CIS_DIALER OPS Automation\\CIS_Base_Count.csv"))
email <- send.mail(from = sender,
                    to = recipients,
                    subject="CIS_Base_Count",
                    body = msg,
                    smtp = list(host.name = "email-smtp.ap-south-1.amazonaws.com", port = 587,
                                user.name = "AKIA6IP74RHPZOVGY5QM",
                                passwd = "BERtlTNx3XLQP3JOUi89sFKfiZpj9mg+y8z9EiKpceij" , ssl = TRUE),
                    authenticate = TRUE,
                   attach.files = c("C:\\R\\CIS_DIALER OPS Automation\\CIS_Base_Count.csv"),
                    send = TRUE)
 
 
 #email$send() # execute to send email
