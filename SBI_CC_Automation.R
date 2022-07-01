#Remove existing List
rm(list=ls())
setwd('E:\\Automation\\SBI')


#Load Libraries
library(data.table)
library(readxl)
library(openxlsx)
library(dplyr)


#Read dump
y_date <- Sys.Date()-1
dump_name <- paste('E:\\Automation\\AppOpsDump\\appopsdump_',y_date,'.csv',sep ='')
dump <- fread(dump_name)
dump <- dump %>% filter(name == 'SBI') %>% 
  mutate(customer_name = paste(first_name,last_name))


#Read Input file
sbi_name <- paste('SBI CC Feedback - ',Sys.Date(),'.xlsx',sep = '')
sbi <- read_excel(sbi_name)
sbi <- sbi %>% 
  filter(ApplicationState != 'FD') %>% 
  select(LeadRefNo,ApplicationState,StatusDesc,CreditLimit,lead_id)

#Read Mapping file
sbi_mapping <- read_excel('./Source/SBI_Mapping.xlsx', sheet = 'SBI_Mapping')
AppOpsMapping <- read_excel('./Source/SBI_Mapping.xlsx', sheet = 'AppOpsMapping')

#Mapping leadId and Application Number
sbi$lead_id <- as.character(sbi$lead_id)
dump$leadid <- as.character(dump$leadid)

sbi <- left_join(sbi,
                 dump[,c('leadid',
                         'customer_name',
                         'phone_home',
                         'offer_application_number',
                         'appops_status_code',
                         'city','status',
                         'customer_type')],
                 by = c('lead_id' = 'leadid'))

sbi <- sbi[!is.na(sbi$offer_application_number),]


#Mapping AppOpsStatus Code
sbi$appops_status_code <- as.character(sbi$appops_status_code)
AppOpsMapping$`App Ops Status Code` <- as.character(AppOpsMapping$`App Ops Status Code`)

sbi <- left_join(sbi,
                 AppOpsMapping,
                 by = c('appops_status_code' = 'App Ops Status Code'))

#Split FA
sbi_fa <- sbi %>% filter(ApplicationState == 'FA') %>% 
  mutate(Remarks = 'Loan Disbursed',
         concatenate = paste(ApplicationState,
                             LeadRefNo,
                             StatusDesc,
                             CreditLimit, sep = ' / '),
         Rejection_Tag = 'Nil',
         Rejection_Category = 'Nil')

#Split WIP
sbi_wip <- sbi %>% filter(ApplicationState == 'WIP') %>% 
  left_join(sbi_mapping, by = c('StatusDesc' = 'Lender Feedback')) %>% 
  mutate(concatenate = paste(ApplicationState,
                             LeadRefNo,
                             StatusDesc, sep = ' / '))


sbi_consolidate <- bind_rows(sbi_fa, sbi_wip)

col_names <- c('lead_id','LeadRefNo','customer_name','phone_home','CreditLimit','city','customer_type',
               'offer_application_number','App Ops Status','Remarks','status','Rejection_Tag',
               'Rejection_Category','concatenate')

sbi_consolidate <- sbi_consolidate %>% 
  select(col_names) %>% 
  distinct(offer_application_number, .keep_all = T)

names(sbi_consolidate)[names(sbi_consolidate) == 'App Ops Status'] <- 'Pre'
names(sbi_consolidate)[names(sbi_consolidate) == 'Remarks'] <- 'Post'


#Write Data in Excel
wb <- createWorkbook()
addWorksheet(wb,"SBI")
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
setColWidths(wb,"SBI",cols = 1:ncol(sbi_consolidate),widths = 15)
writeData(wb,"SBI",sbi_consolidate,borders = "all",headerStyle = hs1)
path <- paste('./Output/SBI_Remarks_',Sys.Date(),'.xlsx',sep='')
saveWorkbook(wb,path,overwrite = T)
openXL(path)

















