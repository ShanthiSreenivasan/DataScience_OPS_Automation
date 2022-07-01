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

setwd('D:\\Automated Dumps - Profiled')

getwd()

thisdate<-format(Sys.Date(),'%Y-%m-%d')


if(!dir.exists(paste("./Output/",thisdate,sep="")))
{
  dir.create(paste("./Output/",thisdate,sep=""))
} 

#380 & 370 Excel Read

hs <- openxlsx::createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize=12,fontName="Arial Narrow", fgFill = "#4F80BD",border = "TopBottomLeftRight",borderColour="black")


Arsdump=fread('./Input/ARS PROFILED CAMPAIGN BASE.csv')

Arsdump_1<-Arsdump %>%filter(nchar(phone_home)>0 & nchar(phone_home) <11) 

Arsdump_2 <- Arsdump_1#[!(Arsdump_1$email_id==''),]


Output_fields <- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'Main sheet')

Output_fields <- Output_fields %>% filter(Command =='1')
Output_fields_name <- Output_fields[,"Output_fields"] 

ARS_DUMP<- Arsdump_2 %>% select(Output_fields_name)

 

for(i in 1: nrow(Output_fields)) {
  
  x<-Output_fields[i,1]
  assign(paste("test",x[1],sep= "_"),x)
  
  
}
ARS_DUMP_1<-ARS_DUMP
if( exists("test_negative_account") )
{
  negative_account <- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'negative_account')
  
  less <- negative_account$less
  Greater <-negative_account$Greater
  n_Command <-negative_account$Command
  
  if(n_Command=='Yes'){
  if(negative_account$NA_if_needed =='No'){
  ARS_DUMP_1 <- ARS_DUMP %>% filter(!is.na(negative_account))
  
  ARS_DUMP_1 <- ARS_DUMP_1 %>% filter(negative_account < less ,
                                  negative_account >= Greater)
  
  
} else{
  
  ARS_DUMP_1 <- ARS_DUMP %>% filter((negative_account < less &
                                      negative_account >= Greater )| is.na(negative_account))
  
  }
}
}

ARS_DUMP_2 <- ARS_DUMP_1

if( exists("test_language_text") )
{
  language_text <- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'language_text')
  
  language_text <- language_text[!is.na(language_text$Logic),]
  
  lang <- language_text$Logic 
  
  language <-language_text$Command[1]
  
  if( language=='Yes'){
    ARS_DUMP_2 <- ARS_DUMP_2 %>% filter(language_text %in% lang)
  } 
}

ARS_DUMP_3 <- ARS_DUMP_2

if( exists("test_Priority") )
{
  Priority <- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'Priority')
  
  Priority_caes <- Priority$Logic 
  
  P_Command <-Priority$Command[1]
  
  P_NA_if_needed <-Priority$NA_if_needed[1]
  
  if( P_Command=='Yes'){
  if(P_NA_if_needed =='No'){
    ARS_DUMP_3 <- ARS_DUMP_3 %>% filter(Priority %in% Priority_caes)
    
  } else{
    ARS_DUMP_3 <- ARS_DUMP_3 %>% filter(Priority %in% Priority_caes| is.na(Priority) | Priority =='NA')
    
  }
}
}

ARS_DUMP_4 <- ARS_DUMP_3

if( exists("test_latest_credit_score") )
{
  latest_credit_score <- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'latest_credit_score')
  
  less <- latest_credit_score$less
  Greater <-latest_credit_score$Greater
  l_command <-latest_credit_score$Command 
  
  if(l_command =='Yes'){
  ARS_DUMP_4 <- ARS_DUMP_4 %>% filter(latest_credit_score < less ,
                                      latest_credit_score >= Greater)
  }
    
}

ARS_DUMP_5 <- ARS_DUMP_4


if( exists("test_Product_family") )
{
  Product_family<- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'Product_family')
  
  CC <- Product_family$CC
  Retail <-Product_family$Retail
  pf_command <-Product_family$Command 
  
  if(pf_command =='Yes'){
   if(CC =='No' & Retail == 'Yes'){
     
     ARS_DUMP_5 <- ARS_DUMP_5[!grepl('credit card', Product_family, ignore.case = T)]
   } else if(CC =='Yes' & Retail == 'No'){
      
      ARS_DUMP_5 <- ARS_DUMP_5[grepl('credit card', Product_family, ignore.case = T)]
    } 
  }
  }
  
ARS_DUMP_6 <- ARS_DUMP_5

if( exists("test_New_Lender_name") )
{
  New_Lender_name <- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'New_Lender_name')
  
  New_Lender_name <- New_Lender_name[!is.na(New_Lender_name$Logic),]
  
  Lender_name <- New_Lender_name$Logic
  
  LN_Command <-New_Lender_name$Command[1]
  
  
  if( LN_Command=='Yes'){
    
    ARS_DUMP_6 <- ARS_DUMP_6 %>% filter(New_Lender_name %in% Lender_name)
  }
}

ARS_DUMP_7 <- ARS_DUMP_6

if( exists("test_decile") )
{
  decile <- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'decile')
  
  less <- decile$less
  Greater <-decile$Greater
  d_Command <-decile$Command
  
  if(d_Command=='Yes'){
    if(decile$NA_if_needed =='No'){
      ARS_DUMP_7 <- ARS_DUMP_7 %>% filter(!is.na(decile))
      
      ARS_DUMP_7 <- ARS_DUMP_7 %>% filter(decile < less ,
                                          decile >= Greater)
      
    } else{
      
      ARS_DUMP_7 <- ARS_DUMP_7 %>% filter((decile < less &
                                             decile >= Greater )| is.na(decile))
      
    }
  }
}

ARS_DUMP_8 <- ARS_DUMP_7

if( exists("test_account_status") )
{
  account_status <- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'account_status')
  
  account_status_1 <- account_status[!is.na(account_status$Logic),]
  
  status <- account_status_1$Logic
  
  account <- account_status %>%  filter(Type %in% status)
  
  fields <- account$fields
  
  LN_Command <-account_status$Command[1]
  
  
  if( LN_Command=='Yes'){
    
    ARS_DUMP_8 <- ARS_DUMP_8 %>% filter(account_status %in% fields)
  }
}

ARS_DUMP_9 <- ARS_DUMP_8

if( exists("test_monthly_income_split") )
{
  monthly_income_split <- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'monthly_income_split')
  
  monthly_income_split <- monthly_income_split[!is.na(monthly_income_split$Logic),]
  
  Split <- monthly_income_split$Logic
  
  LN_Command <-monthly_income_split$Command[1]
  
  
  if( LN_Command=='Yes'){
    
    ARS_DUMP_9 <- ARS_DUMP_9 %>% filter(monthly_income_split %in% Split)
  }
}


ARS_DUMP_10 <- ARS_DUMP_9

if( exists("test_outstanding_amount") )
{
  outstanding_amount <- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'outstanding_amount')
  
  less <- outstanding_amount$less
  Greater <-outstanding_amount$Greater
  l_command <-outstanding_amount$Command 
  
  if(l_command =='Yes'){
    ARS_DUMP_10 <- ARS_DUMP_10 %>% filter(outstanding_amount < less ,
                                          outstanding_amount >= Greater)
  }
  
}


#ARS_DUMP_10 <- ARS_DUMP_10[!is.na(ARS_DUMP_10$cm_base_id)]
ARS_DUMP_10 <- ARS_DUMP_10[!is.na(ARS_DUMP_10$lender_id)]

source <- read.xlsx(".\\Input\\Automation_Note.xlsx",sheet = 'Source')

File_name <- source$Source[1]

Name1 <- source$Source[2]
Name2<- source$Source[3]
Name3 <- source$Source[4]
Name4 <- source$Source[5]

ARS_DUMP_10$campaign_name <- paste(Name1,Name2,Name3,Name4,"&utm_term=CMBASEAUTO100&tXcf=",
                                   ARS_DUMP_10$tkn,
                                   "&utm_lender_id=",
                                   ARS_DUMP_10$lender_id,
                                   "&utm_cmbase_id=",
                                   ARS_DUMP_10$cm_base_id,
                                   sep = '')

source('.\\Function file.R')

hs <- openxlsx::createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize=12,fontName="Arial Narrow", fgFill = "#4F80BD",border = "TopBottomLeftRight",borderColour="black")


Filename <- paste('.\\Output\\',thisdate,'\\',File_name,sep = '')
FileCreate(ARS_DUMP_10,sheet_name="Sheet1",Filename)
