#Load Basic Libraries
rm(list = ls())
setwd('E:\\Automation\\MoneyView')

library(data.table)
library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)

today <- Sys.Date()
ydate <- Sys.Date()-1

month<- format(Sys.Date(),"%m")
year<- format(Sys.Date(),"%Y")
thismonth <-paste(year,"-",month,"-01",sep = "")



#Load Dump
dump_path <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",ydate,".csv",sep='')
dump <- fread(dump_path)
dump <- dump %>% filter(name == 'Money View')



#Filter Only MoneyView
appopscode <- read_excel('Appops Moneyview Code.xlsx',sheet = 'App Ops Status')
mapping_leads <- read_excel('Appops Moneyview Code.xlsx',sheet = 'leads')
mapping_install_1 <- read_excel('Appops Moneyview Code.xlsx',sheet = 'installed1')
mapping_install_2 <- read_excel('Appops Moneyview Code.xlsx',sheet = 'installed2')

#Load MoneyView leads Input File

path <-paste(".//Input//CreditMantriLead_MIS_External_",today,".xlsx",sep = "" )
money_in <- read_excel(path,sheet = "leads")

money_leads <- money_in %>% filter(lead_date >= thismonth & lead_status == "CHECK_ELIGIBILITY_SUCCESS"
                                        & app_install == "1")

money_leads$ph_no <- as.character(money_leads$ph_no)
dump$phone_home <- as.character(dump$phone_home)



money_leads <- money_leads %>% select('name','ph_no','lead_status','portal_offer') 

df_leads <- left_join(money_leads, dump[,c('phone_home',
                                  'offer_application_number',
                                  'appops_status_code',
                                  'status')],
                by = c('ph_no' = 'phone_home'))

df_leads <- left_join(df_leads, appopscode, by = c('appops_status_code'='App Ops Status Code'))

df_leads <- left_join(df_leads, mapping_leads, by = 'lead_status')

df_new <- df_leads %>%  select('name','ph_no','status','offer_application_number','Pre','POST','lead_status','portal_offer','x1','x2')

df_new1<-df_new[ !is.na(df_new$offer_application_number)  , ] %>% distinct(ph_no, .keep_all = T)

#installed -----------------------------------------------

money_inst <- read_excel(path,sheet = "installed") %>% filter(submission_date == "NULL" )

                                         
money_inst$ph_no <- as.character(money_inst$ph_no)


money_inst <- money_inst %>% select('name','ph_no','loan_status') 

df_inst <- left_join(money_inst, dump[,c('phone_home',
                                           'offer_application_number',
                                           'appops_status_code',
                                           'status')],
                      by = c('ph_no' = 'phone_home'))

df_inst <- left_join(df_inst, appopscode, by = c('appops_status_code'='App Ops Status Code'))

df_inst <- left_join(df_inst, mapping_install_1, by = 'loan_status')

df_new_inta <- df_inst %>%  select('name','ph_no','status','offer_application_number','Pre','POST','loan_status','xx1','xx2')

df_new_inta1<-df_new_inta[ !is.na(df_new_inta$offer_application_number)  , ] %>% distinct(ph_no, .keep_all = T)


#installed1 -----------------------------------------------


money_inst_this <- read_excel(path,sheet = "installed")


money_inst_this <- money_inst_this %>% filter(submission_date >= thismonth & submission_date !="NULL")


money_inst_this$ph_no <- as.character(money_inst_this$ph_no)


money_inst_this <- money_inst_this %>% select('name','ph_no','loan_status') 

df_inst1 <- left_join(money_inst_this, dump[,c('phone_home',
                                         'offer_application_number',
                                         'appops_status_code',
                                         'status')],
                     by = c('ph_no' = 'phone_home'))

df_inst1 <- left_join(df_inst1, appopscode, by = c('appops_status_code'='App Ops Status Code'))

df_inst1 <- left_join(df_inst1, mapping_install_2, by = 'loan_status')

df_new_inta2 <- df_inst1 %>%  select('name','ph_no','status','offer_application_number','Pre','POST','loan_status','xx1','xx2')

df_new_inta2<-df_new_inta2[ !is.na(df_new_inta$offer_application_number)  , ] %>% distinct(ph_no, .keep_all = T)




df_installed <- rbind(df_new_inta1,df_new_inta2) 

df_installed <- df_installed %>% distinct(offer_application_number, .keep_all = T)


#disbursed cases------------------------------------------------

money_disbursed <- read_excel(path,sheet = "disbursed")

money_disbursed <- money_disbursed %>% filter(disbursal_date >= thismonth)

money_disbursed$ph_no <- as.character(money_disbursed$ph_no)


money_disbursed <- money_disbursed %>% select('name','ph_no','loan_status','disb_amount','loan_id') 

df_disb <- left_join(money_disbursed, dump[,c('phone_home',
                                               'offer_application_number',
                                               'appops_status_code',
                                               'status')],
                      by = c('ph_no' = 'phone_home'))

df_disb <- left_join(df_disb, appopscode, by = c('appops_status_code'='App Ops Status Code'))

df_disb$POST="Loan Disbursed"
df_disb$x="Nil"
df_disb$y="Nil"

df_new_disb <- df_disb %>%  select('name','ph_no','status','offer_application_number','Pre','POST','loan_status','disb_amount','loan_id','x','y')

df_new_disb1<-df_new_disb[ !is.na(df_new_disb$offer_application_number)  , ] %>% distinct(ph_no, .keep_all = T)



#Create excel--------------------------------------------------------------------------

wb <-createWorkbook()
addWorksheet(wb,"Leads")
addWorksheet(wb,"Installed")
addWorksheet(wb,"Disbursed")
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
setColWidths(wb,"Leads",cols = 1:ncol(df_new1),widths = 15)
writeData(wb,"Leads",df_new1,borders = "all",headerStyle = hs1)

setColWidths(wb,"Installed",cols = 1:ncol(df_installed),widths = 15)
writeData(wb,"Installed",df_installed,borders = "all",headerStyle = hs1)

setColWidths(wb,"Disbursed",cols = 1:ncol(df_new_disb1),widths = 15)
writeData(wb,"Disbursed",df_new_disb1,borders = "all",headerStyle = hs1)

path <- paste('./Back up/MoneyView_Remarks_',Sys.Date(),'.xlsx',sep='')

saveWorkbook(wb,path,overwrite = T)
openXL(path)


