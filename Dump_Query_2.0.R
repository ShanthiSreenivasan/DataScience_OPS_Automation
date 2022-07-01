#Remove existing List
rm(list=ls())

#Load Library
library(nanotime)
library(data.table)
library(openxlsx)
library(dplyr)
library(sqldf)
library(lubridate)

#Set Working Directory
setwd("E:\\Automation\\Dump")

ydate <- Sys.Date()-1
#Read Download Master
DMPath <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",ydate,".csv",sep='')
dm <- fread(DMPath, sep = ",",na.strings = NULL ) 
dm$appops_status_code <- as.integer(dm$appops_status_code)
# dm[, `:=`(
#   date_of_referral = dmy_hm(date_of_referral)
# )]

#Rename duplicate column name
# names(dm)[34] <- "New_name"


#Read AppOpsCode
Ops <- read.xlsx(paste(".\\AppOpsCode\\AppOpsCode.xlsx"),1)
names(Ops) <- c("Ops_Status_Code","Ops_Status")
Ops$Ops_Status_Code <- as.integer(Ops$Ops_Status_Code)
#Write Query
Quary <- "SELECT 
            dm.leadid,
            dm.first_name,
            dm.last_name,
            dm.first_name || \" \" || dm.last_name AS Concat,
            dm.phone_home,
            dm.offer_reference_number,
            dm.offer_application_number,
            dm.date_of_referral,
            dm.bank_feedback_date,
            dm.followup_date,
            dm.lender_followup_date,
            dm.status,
            dm.name,
            dm.appops_status_code,
            Ops.Ops_Status AS App_Ops_State,
            dm.customer_type
            FROM dm
            LEFT JOIN Ops ON dm.appops_status_code = Ops.Ops_Status_Code"

#Get Data by using Query
DMFinal <- sqldf(Quary)
DMFinal$phone_home <- as.numeric(DMFinal$phone_home)


#Write dump file into excel
wb <- createWorkbook()
addWorksheet(wb,"New")
setColWidths(wb,"New", cols=1:ncol(DMFinal), widths = 12)
writeData(wb, "New", DMFinal)
fname <- paste(".\\",format(Sys.Date(), "%d-%m-%y"),".xlsx",sep="")
saveWorkbook(wb, fname, overwrite = T)
openXL(fname)

print("Dump Created")

