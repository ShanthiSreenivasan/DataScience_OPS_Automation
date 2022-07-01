rm(list = ls())
setwd('E:\\Automation\\Muthoot')

library(data.table)
library(dplyr)
library(readxl)
library(openxlsx)
library(lubridate)

today <- Sys.Date()
ydate <- Sys.Date()-1

dump_path <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",ydate,".csv",sep='')
dump <- fread(dump_path)
dump <- dump %>% filter(name == 'MUTHOOT')


appopscode <- read_excel('Appops Moneyview Code.xlsx',sheet = 'App Ops Status')
mapping_call <- read_excel('Appops Moneyview Code.xlsx',sheet = 'Call center')
mapping_Loaction <- read_excel('Appops Moneyview Code.xlsx',sheet = 'Location')


Muthoot_Call_path <- paste("E:\\Automation\\Muthoot\\Muthoot Call center.xlsx",sep = '')
Muthoot_Call <- read_excel(Muthoot_Call_path) %>% select("Loan Reference Number","Mobile Number","Customer Name","Loan Amount","Disposition","Remarks")
Muthoot_Call$Type="Call Center"


Muthoot_Loca_path <-paste("E:\\Automation\\Muthoot\\Muthoot Location.xlsx",sep = '')
Muthoot_Location <- read_excel(Muthoot_Loca_path) %>% select("Loan Reference Number","Mobile Number","Customer Name","Loan Amount","Final status","Remarks")
Muthoot_Location$Type="Location"

dump$phone_home <-as.character(dump$phone_home)
Muthoot_Call$`Mobile Number` <-as.character(Muthoot_Call$`Mobile Number`)
Muthoot_Location$`Mobile Number` <-as.character(Muthoot_Location$`Mobile Number`)
Muthoot_Call$`Loan Reference Number` <-as.character(Muthoot_Call$`Loan Reference Number`)
Muthoot_Location$`Loan Reference Number` <-as.character(Muthoot_Location$`Loan Reference Number`)


df_Call <-left_join(Muthoot_Call,dump[,c('phone_home',
                                          'offer_application_number',
                                          'appops_status_code',
                                          'status')],
                     by=c("Mobile Number"="phone_home"))


df_Location <-left_join(Muthoot_Location,dump[,c('phone_home',
                                          'offer_application_number',
                                          'appops_status_code',
                                          'status')],
                     by=c("Mobile Number"="phone_home"))


df_Call <- left_join(df_Call, appopscode, by = c('appops_status_code'='App Ops Status Code'))

df_Call <- left_join(df_Call, mapping_call, by = 'Disposition')

df_Location <- left_join(df_Location, appopscode, by = c('appops_status_code'='App Ops Status Code'))

df_Location <- left_join(df_Location, mapping_Loaction, by = 'Final status')

df_Call <- df_Call %>% select("Loan Reference Number","Customer Name","Mobile Number","Loan Amount","offer_application_number","Pre","POST","Type","Disposition","Remarks","Rejection_Tag","Rejection_Category")

df_Location <- df_Location %>% select("Loan Reference Number","Customer Name","Mobile Number","Loan Amount","offer_application_number","Pre","POST","Type","Final status","Remarks","Rejection_Tag","Rejection_Category")

names(df_Location) <-c("Loan Reference Number","Customer Name","Mobile Number","Loan Amount","offer_application_number","Pre","POST","Type","Disposition","Remarks","Rejection_Tag","Rejection_Category")

df_Call$`Loan Amount` <- as.character(df_Call$`Loan Amount`)
df_Location$`Loan Amount` <-as.character(df_Location$`Loan Amount`)



df_new <- bind_rows(df_Call,df_Location)

df_new$Comment <-paste(df_new$`Loan Reference Number`,"/",df_new$Disposition,"/",df_new$Remarks)


df_data2 <- df_new %>% filter(is.na(Rejection_Category))

df_data1 <- df_new %>% filter(!is.na(Rejection_Category))



#Regex creation-----------------------------------------
docs_reject_regex <- "(CUSTOMER DON'T HAVE|customer dont have)"
geo_location_regex <- "(CUSTOMER IS STAYING|NOT AN APPROVED|not apporved|APPROVED LOCATION)"
rented_regex <- "(rent case|surety|rented|ranted|staying|surity|residance|resi |guarantor|not residing|rent hourse|rent house|customer address|not address)"
not_contact_regex <- "(call| contact|switched off|respon|reachable|phone|ph |fake ptp|switch |wrong num|connect|not answer|not lift|not rechible)"
not_intrest_regex <- "(customer delay|not int|no need|not req|dont want loan|no requirement|not come|just info|no requirenment|customer not  interst|rate of intrest high|tenure|no  requirement|need|loan req|no  requirement|interest rate)"
cash_regex <- "(SALARY BY HAND|SALARY BY CASH|salary by hand|CUSTOMER SALARY TAKEN BY CASH)"
income_regex <- "(SALARY IS|CUSTOMER NTH SALARY)"
default_regex <- "(CUSTOMER IS WORKING)"
existing_regex <- "(alredy|already|dedupe|loan live|scuf|customer allready existing loan|log decctention|allready apporved loan|exsting|multiple loan|existing customer)"
foir_dbr_regex <- "(poor bank)"
duplicate_regex <- "(duplicate|double)"
cibil_regex <- "(fi |cibil|f.i|customer very low banking|customer working in|no banking|low banking|co applicant|tym to tym|not eligible|rejected|negative|many bouncing)"
policy_regex <- "(call back|NOT QUALIFIED FOR PL)"
Loan_amount <- "(overliverage|log decctention)"


#Rejection Tag-------------------------------
df_data2 <- df_data2 %>%
  mutate(
    Rejection_Tag = ifelse((grepl(docs_reject_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(geo_location_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(rented_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(not_contact_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NC",
                    ifelse((grepl(not_intrest_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NI",
                    ifelse((grepl(cash_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(income_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(default_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(existing_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(foir_dbr_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(duplicate_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(cibil_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(policy_regex,.$Remarks)&(is.na(.$Rejection_Tag))),"NE",
                    ifelse( is.na(.$Remarks),"NE", "")))))))))))))))


#Rejection_Category--------------------------
df_data2 <- df_data2 %>%
  mutate(
    Rejection_Category =  ifelse((grepl(docs_reject_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"No documents",
                          ifelse((grepl(geo_location_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"GeoLocation",
                          ifelse((grepl(rented_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"Residence Type",
                          ifelse((grepl(not_contact_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"Not Contactable",
                          ifelse((grepl(not_intrest_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"Rate Shopping / Enquiries in case of future need",
                          ifelse((grepl(cash_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"Salary Type",
                          ifelse((grepl(income_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"Income Slab",
                          ifelse((grepl(default_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"Firm Type",
                          ifelse((grepl(existing_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"Existing Product Holder",
                          ifelse((grepl(foir_dbr_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"FOIR / DBR Reasons",
                          ifelse((grepl(duplicate_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"Duplicate lead",
                          ifelse((grepl(cibil_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"CIBIL Reject - Negative record in CIBIL",
                          ifelse((grepl(policy_regex,.$Remarks)&(!is.na(.$Rejection_Tag))),"Miscellaneous Policy",
                          ifelse( is.na(.$Remarks),"Miscellaneous Policy", "")))))))))))))))


df_new1 <- rbind(df_data1,df_data2)

#create Excel -------------------------------------------------------

wb <- createWorkbook()
addWorksheet(wb,"Muthoot")
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
setColWidths(wb,"Muthoot",cols = 1:ncol(df_new1),widths = 15)
writeData(wb,"Muthoot",df_new1,borders = "all",headerStyle = hs1)
path <- paste("./Output/Muthoot_Remarks_",today,".xlsx",sep = "")

saveWorkbook(wb,path,overwrite = T)
openXL(path)
