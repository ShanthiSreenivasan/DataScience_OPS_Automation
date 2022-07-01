#Load Basic Libraries
rm(list = ls())
setwd('E:\\Automation\\PNB_REMARKS')

library(data.table)
library(dplyr)
library(readxl)
library(openxlsx)

today <- Sys.Date()
today <-format(today, format="%m%d%Y")
ydate <- Sys.Date()-1

#Load Dump
dump_path <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",ydate,".csv",sep='')
dump <- fread(dump_path)
dump <- dump %>% filter(name == 'PNB')


path1=paste(".\\Input\\Instant Process",today,".xlsx",sep = '')
path2=paste(".\\Input\\Instant Reject",today,".xlsx",sep = '')
path3=paste(".\\Input\\Post Login Queue",today,".xlsx",sep ='')
path4=paste(".\\Input\\RO Process",today,".xlsx",sep = '')
path5=paste(".\\Input\\RO Reject",today,".xlsx",sep ='')


#Read excels
PNB_in1 <- readxl::read_excel(path1)
PNB_in2 <- readxl::read_excel(path2)
PNB_in3 <- readxl::read_excel(path3)
PNB_in4 <- readxl::read_excel(path4)
PNB_in5 <- readxl::read_excel(path5)

PNB_in1$stage = "Instant Process"
PNB_in2$stage = "Instant Reject"
PNB_in3$stage = "Post Login Queue"
PNB_in4$stage = "RO Process"
PNB_in5$stage = "RO Reject"



#filter Excel

PNB_in1 <- PNB_in1 %>% select('Lead_no','Mobile','Loan_amount','Call_status','stage')
PNB_in2 <- PNB_in2 %>% select('Lead_no','Mobile','Loan_amount','Call_status','stage')
PNB_in3 <- PNB_in3 %>% select('Lead_no','Mobile','Loan_amount','Call_status','stage')
PNB_in4 <- PNB_in4 %>% select('Lead_no','Mobile','Loan_amount','Call_status','stage')
PNB_in5 <- PNB_in5 %>% select('Lead_no','Mobile','Loan_amount','Call_status','stage')

PNB_Final <-rbind(PNB_in1,PNB_in2,PNB_in3,PNB_in4,PNB_in5)



appopscode <- read_excel('Appops Code.xlsx',sheet = 'App Ops Status')
mapping <- read_excel('Appops Code.xlsx',sheet = 'Sheet1')

PNB_Final$Mobile <- as.character(PNB_Final$Mobile)
dump$phone_home <- as.character(dump$phone_home)


df <- left_join(PNB_Final, dump[,c('phone_home',
                                  'offer_application_number',
                                  'appops_status_code',
                                  'status')],
                by = c('Mobile' = 'phone_home'))



df <- left_join(df, appopscode, by = c('appops_status_code'='App Ops Status Code'))
df <- left_join(df, mapping, by = c('stage'='appstatus'))

df$Comments <-paste(df$Lead_no,"/",df$stage,"/",df$Call_status)

col_names <- c('Lead_no','Mobile','offer_application_number','App Ops Status','AppOpsCode_New','stage','Call_status','Loan_amount','X1','X2','Comments')

df_new <- df %>% select(col_names)

df_new1<-df_new[ !is.na(df_new$offer_application_number)  , ] %>% distinct(ph_no, .keep_all = T) %>% distinct(offer_application_number, .keep_all = T)


wb <- createWorkbook()
addWorksheet(wb,"PNB")
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
setColWidths(wb,"PNB",cols = 1:ncol(df_new1),widths = 15)
writeData(wb,"PNB",df_new1,borders = "all",headerStyle = hs1)
path <- paste('PNB_Remarks_',Sys.Date(),'.xlsx',sep='')
saveWorkbook(wb,path,overwrite = T)
openXL(path)











