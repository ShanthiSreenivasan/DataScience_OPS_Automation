#Remove existing list
rm(list =ls())
setwd('E:\\Automation\\Paysense')

#Load Libraries
library(data.table)
library(readxl)
library(openxlsx)
library(dplyr)

#Get require file name
today <- format(Sys.Date(),'%Y-%m-%d')
yesterdate <- format(Sys.Date()-1, '%Y-%m-%d')
dump_name <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",yesterdate,".csv",sep='')
api_file_name <- paste('./CreditMantri_api_',today,'.csv',sep='')
#crm_file_name <- paste('./1_Paysense/CreditMantri_crm_',today,'.csv',sep='')

#Load Files
dump <- fread(dump_name, na.strings = c('',NA))
api_df <- fread(api_file_name,na.strings = c('',NA))
#crm_df <- fread(crm_file_name, na.strings = c('',NA))
#api_df <- bind_rows(api_df,crm_df)

#Reshape Dump file
dump <- dump %>% 
  filter(name == 'Paysense') %>% 
  select(phone_home,offer_application_number,appops_status_code)

#Get Phone, CustName 
api_df$Phone <- substr(api_df$phone_id, 3,12)
api_df$CustName <- paste(api_df$first_name, api_df$last_name)
api_df$Notes <- paste(api_df$user_status,api_df$application_current_status, sep = '/')

#Join api_df and df
api_df$Phone <- as.character(api_df$Phone)
dump$phone_home <- as.character(dump$phone_home)

#Required data
df <- left_join(api_df,
                dump, 
                by = c( 'Phone' = 'phone_home')) %>% 
  filter(!is.na(offer_application_number)) %>% 
  select(CustName, Phone,offer_application_number,appops_status_code,Notes,amount,
         credit_approved,loan_application_created_at,perfios_done,
         application_submitted,application_rejected,application_preapproved,
         application_approved,disbursed_date) %>% 
  distinct(Phone,.keep_all = T) 


#PostCode
df$PostCode <- NA
df$PostCode[complete.cases(df$disbursed_date)] <- '990'
df$PostCode[is.na(df$PostCode) & complete.cases(df$application_approved)] <- '690'
df$PostCode[is.na(df$PostCode) & complete.cases(df$application_preapproved)] <- '590'
df$PostCode[is.na(df$PostCode) & complete.cases(df$application_rejected)] <- '580'
df$PostCode[is.na(df$PostCode) & complete.cases(df$application_submitted)] <- '500'
df$PostCode[is.na(df$PostCode) & complete.cases(df$perfios_done)] <- '490'
df$PostCode[is.na(df$PostCode) & complete.cases(df$loan_application_created_at)] <- '400'
df$PostCode[is.na(df$PostCode) & complete.cases(df$credit_approved)] <- '390'
df$Notes <- paste(df$Notes,df$PostCode, sep = '/')

#Get Unique Values
df_unique <- df[!is.na(df$PostCode),]
df_unique <- df_unique[df_unique$appops_status_code != df_unique$PostCode,]
df_unique$amount[!df_unique$PostCode %in% c(690,990)] <- NA
df_unique$code <- NA


#Split Data by appops code wise
#disbursed_date
df_990 <- df_unique %>% 
  filter(PostCode %in% c(690,990)) %>% 
  filter(appops_status_code != 990)  %>%        
  mutate(code = 'Loan Booked')

#application_preapproved
df_590 <- df_unique %>% 
  filter(PostCode == 590) %>% 
  filter(!appops_status_code >= 590) %>% 
  mutate(code = 'LOS approved')

#application_rejected
df_580 <- df_unique %>% 
  filter(PostCode == 580) %>% 
  filter(appops_status_code <= 600) %>% 
  mutate(code = 'LOS rejected')

#application_submitted
df_500 <- df_unique %>% 
  filter(PostCode == 500)%>% 
  filter(appops_status_code <= 590) %>% 
  mutate(code = 'LOS pending decision')

#perfios_done
df_490 <- df_unique %>% 
  filter(PostCode == 490)%>% 
  filter(appops_status_code <= 480) %>% 
  mutate(code = 'LOS pending decision')
  
#loan_application_created_at
df_400 <- df_unique %>% 
  filter(PostCode == 400) %>% 
  filter(appops_status_code <= 480) %>% 
  mutate(code = 'Docs in process - Security Not identified/Not Applicable')

#credit_approved
df_390 <- df_unique %>% 
  filter(PostCode == 390) %>% 
  filter(appops_status_code <= 390) %>% 
  mutate(code = 'AIP Approved/ Interested in docs')

#Combine all together
df_final <- bind_rows(df_390,
                      df_400,
                      df_490,
                      df_500,
                      df_580,
                      df_590,
                      df_990)

df_final$Rejection_Tag <- case_when(
  
  df_final$PostCode %in% c(390,400,490,500,590,690,990) ~ 'Nil',
  df_final$PostCode == 580 ~ 'NE',
  T ~ ''
  
)

df_final$Rejection_Category <- case_when(
  
  df_final$PostCode %in% c(390,400,490,500,590,690,990) ~ 'Nil',
  df_final$PostCode == 580 ~ 'Miscellaneous Policy',
  T ~ ''
  
)

#Create Up2 File
df_up2 <- df_final
df_up2 <- df_up2 %>% 
  select(offer_application_number,code,Notes,amount,Rejection_Tag,Rejection_Category)


#Create Excel File
wb <- createWorkbook()
#addWorksheet(wb, 'df_final')
addWorksheet(wb, 'df_up2')
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
#setColWidths(wb, 'df_final',cols = 1:ncol(df_final),widths = 'auto')
#writeData(wb,'df_final',df_final,headerStyle = hs1,borders = 'all')
setColWidths(wb, 'df_up2',cols = 1:ncol(df_up2),widths = 'auto')
writeData(wb,'df_up2',df_up2,headerStyle = hs1,borders = 'all')

Filename <- paste('./Output/Paysense_Feedback_', format(Sys.Date(), '%Y-%m-%d'),'.xlsx',sep = '')
saveWorkbook(wb = wb,file = Filename,overwrite = T)
openXL(Filename)

#Finished-----------------------------------------------------------------------