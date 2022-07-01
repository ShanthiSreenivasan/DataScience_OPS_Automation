#Load Basic Libraries
rm(list = ls())
setwd('E:\\Automation\\SBI CC Remarks')

library(data.table)
library(dplyr)
library(readxl)
library(openxlsx)

today <- Sys.Date()
ydate <- Sys.Date()-1

#Load Dump
dump_path <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",ydate,".csv",sep='')
dump <- fread(dump_path)

SBI_in <- read.xlsx(".//SBI CC_Remarks_input.xlsx")
appopscode <- read_excel('SBI_CC_Comments.xlsx',sheet = 'App Ops Status Code')
mapping <- read_excel('SBI_CC_Comments.xlsx',sheet = 'Sheet1')
mapping3 <- read_excel('SBI_CC_Comments.xlsx',sheet = 'Sheet3')


dump$leadid<-as.character(dump$leadid)
SBI_in$GE_MID1<-as.character(SBI_in$GE_MID1)


df <- left_join(SBI_in,dump[,c("leadid","phone_home",
                               "offer_application_number",
                               "appops_status_code",
                               "status")],
                by = c("GE_MID1"="leadid"))

df1<- left_join(df,appopscode,by = c('appops_status_code'='App Ops Status Code') )



df1 <- df1 %>% select("AGENCY.STATUS","GE_MID1","Soft.Decision","Decline.Code",
                    "Decline.Description","offer_application_number","Pre",
                    "phone_home","appops_status_code",
                    "status")
df_Dec <- df1 %>% filter(AGENCY.STATUS == 'Final Decline') 

df_Dec<-left_join(df_Dec,mapping,by = 'Soft.Decision')
df_Dec$XX=""
df_Dec$XX1=""
df_Dec$XX2=""





df_Pen<-df1 %>% filter(AGENCY.STATUS %in% c('Pending for Dispatch','Post Dispatch Stage',
                                            'Pre Dispatch  Decline','Null/Pre soft decision'))

df_Pen<-left_join(df_Pen,mapping3,by = 'AGENCY.STATUS')

df_Pen$XX=""
df_Pen$XX1=""
df_Pen$XX2=""


Col_names<-c('AGENCY.STATUS','GE_MID1','Soft.Decision','Decline.Code','Decline.Description',
              'offer_application_number','Pre','Post_upload','XX','XX1','XX2','appops_status_code','Post_Status')




df_Dec <- df_Dec %>% select(Col_names)
df_Pen <- df_Pen %>% select(Col_names)

df_final  <- rbind(df_Dec,df_Pen) 
df_final <- df_final[ !is.na(df_final$offer_application_number)  , ]

df_final$XX2 <- paste(df_final$AGENCY.STATUS,"/",df_final$Soft.Decision,"/",df_final$Decline.Code,"/",df_final$Decline.Description)


df_final_all <-df_final %>% filter(Post_upload != "AIP rejected")


df_final_300 <- df_final %>% 
  filter(Post_upload == "AIP rejected",
         Pre %in% 
           c("Applied Online","Closed prior to S2L","Send to Lender","Application Forwarded to Lender",
             "Initial FB - Contact successful","Resent to Lender - Initial FB - Security Identified",
             "Resent to Lender - Initial FB - Security Not Identified/Not Applicable","AIP Approved/ Interested in docs",
              NA))
df_final_300$Post_upload <- "AIP rejected"  

df_final_400 <- df_final %>% filter(Post_upload == "AIP rejected",
         Pre %in% c("Docs in process - Security Not identified/Not Applicable","Docs in process - Security identified",
                    "Resent to Lender - Docs - Security Identified","Resent to Lender - Docs - Security Not Identified/Not Applicable",
                    "Docs NC","Docs NI","Docs - LTC","Docs stage - Rejected","Docs Complete",NA))

df_final_400$Post_upload<-"Docs stage - Rejected"

df_final_500 <- df_final %>% filter(Post_upload == "AIP rejected",
                                    Pre %in% c("LOS pending decision","LOS approved",NA))

df_final_500$Post_upload<-"LOS rejected"

df_fin<- rbind(df_final_300,df_final_400,df_final_500,df_final_all)

wb <- createWorkbook()
addWorksheet(wb,"SBI CC Remarks")
hs1 <- createStyle(fgFill = "#4F81BD", 
                   halign = "CENTER", 
                   textDecoration = "Bold",
                   border = "Bottom", 
                   fontColour = "white")
setColWidths(wb,"SBI CC Remarks",cols = 1:ncol(df_fin),widths = 15)
writeData(wb,"SBI CC Remarks",df_fin,borders = "all",headerStyle = hs1)
path <- paste('SBI_CC_Remarks_',Sys.Date(),'.xlsx',sep='')
saveWorkbook(wb,path,overwrite = T)
openXL(path)






