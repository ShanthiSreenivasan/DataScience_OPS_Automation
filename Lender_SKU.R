  rm(list=ls())
  setwd('E:\\Automation\\Lender_SKU')
  Sys.getenv("R_ZIPCMD", "zip")
  
  library(data.table)
  library(dplyr)
  library(lubridate)
  library(data.table)
  ydate <- Sys.Date()
  
  bank_fdate <- Sys.Date()-2
  #bank_fdate <- '2019-08-31'
  referral_date <- Sys.Date()-2
  #referral_date <- '2019-08-31'
  referral_date_7 <- Sys.Date()-7
  
  y_date <- Sys.Date()- 1
  dump_name <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",y_date,".csv",sep ='')
  dump <- fread(dump_name)
  dump$date_of_referral <- format(ymd(dump$date_of_referral),'%Y-%m-%d')
  dump$bank_feedback_date <- as.Date(format(ymd_hms(dump$bank_feedback_date),'%Y-%m-%d'))
  
  #350,360
  dump_d1 <- dump[dump$bank_feedback_date >= bank_fdate &
                    dump$date_of_referral >= referral_date_7,]
  
  
  NI_NC <- dump_d1[grepl("350|360|450|460",dump_d1$appops_status_code),]
  
  NI_NC_CC <- NI_NC[grepl("ICICI|IndusInd|RBL|SBI|Yes",NI_NC$name,ignore.case = T) &
                      grepl('CC',NI_NC$status),]
  
  NI_NC_PL <- NI_NC[grepl("HDB|HDFC|RBL|Shriram",NI_NC$name,ignore.case = T) &
                      grepl('PL',NI_NC$status),]
  
  #MoneyTap
  MoneyTap <- NI_NC[grepl('Money Tap',NI_NC$name,ignore.case = T),]
  
  #CaptialFloat
  capital <- NI_NC[grepl('CAPITAL FLOAT',NI_NC$name,ignore.case = T) ,]
  
  #270
  dump_d2 <- dump[dump$date_of_referral == referral_date,]
  CashE_ES <- dump_d2[grepl('Early Salary|CashE', dump_d2$name, ignore.case = T) & 
                        grepl('VX|VT', dump_d2$applied_oic, ignore.case = T) &
                        appops_status_code == 270,]
  
  #Credy
  credy <- dump[grepl('Credy',dump$name, ignore.case = T) &
                  grepl('Stage 2', dump$lender_notes, ignore.case = T) &
                  grepl('IM', dump$applied_oic, ignore.case = T),]
  
  credy <- credy %>% filter(bank_feedback_date == max(credy$bank_feedback_date))
  
  
  
  final_df <- bind_rows(NI_NC_CC,
                        NI_NC_PL,
                        CashE_ES,
                        credy,
                        MoneyTap,
                        capital)
  
  final_df <- as.data.table(final_df)
  final_df_bu <- final_df[grepl('(System|PORTFOLIO|andriodApp)', oic, ignore.case = T)]
  final_df_assisted <- final_df[!grepl('(System|PORTFOLIO|andriodApp)', oic, ignore.case = T)]
  
  source_name <- paste("./Lender Reverse Feedback SKUs - ", Sys.Date(),'.csv', sep = '')
  zip_name <- paste("./Lender Reverse Feedback SKUs - ", Sys.Date(),'.zip', sep = '')
  fwrite(final_df, source_name, row.names = F)
  
  source_name_1 <- paste("./Lender Reverse Feedback SKUs BU - ", Sys.Date(),'.csv', sep = '')
  zip_name_1 <- paste("./Lender Reverse Feedback SKUs BU - ", Sys.Date(),'.zip', sep = '')
  fwrite(final_df_bu, source_name_1, row.names = F)
  
  
  source_name_2 <- paste("./Lender Reverse Feedback SKUs Assisted - ", Sys.Date(),'.csv', sep = '')
  zip_name_2 <- paste("./Lender Reverse Feedback SKUs Assisted - ", Sys.Date(),'.zip', sep = '')
  fwrite(final_df_assisted, source_name_2, row.names = F)
  
  Sys.setenv('R_ZIPCMD' = 'C:/Rtools/bin/zip.exe')
  zip(zip_name,source_name, flags = paste("-j -r9Xj -P", 'Sku!123'))
  zip(zip_name_1,source_name_1, flags = paste("-j -r9Xj -P", 'Sku!123'))
  zip(zip_name_2,source_name_2, flags = paste("-j -r9Xj -P", 'Sku!123'))
  
