rm(list = ls()) # Clear environment
knitr::opts_chunk$set(comment="", echo = FALSE, message=FALSE, warning=FALSE)
library(plyr); library(dplyr)
library(car);library(xlsx)
library(stringr);library(excel.link)
library(mailR);library(xtable)
library(lubridate);library (formattable)
library(openxlsx)
Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")


setwd(Sys.getenv('CIS_Batch'))

hs <- openxlsx::createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize=12,fontName="Arial Narrow", fgFill = "#4F80BD",border = "TopBottomLeftRight",borderColour="black")

dump_wh <- read.csv("E:\\Input\\CISDump_WH.csv",stringsAsFactors = FALSE)

dump_wh$Lead_Check_Wh <- paste(dump_wh$lead_id,"/",dump_wh$lender,"/",dump_wh$Account_No,sep = '') 

today <- format(Sys.Date(),"%Y-%m-%d")
today1 <-format(Sys.Date(),"%d-%m-%Y")

Sanker_Folder <- paste('./Output/',today,'/Sanker_Batch',sep = '')

Santra_Folder <- paste('./Output/',today,'/Santra_Batch',sep = '')


#Fullerton_New Validation

Fullerton_New <- paste(Sanker_Folder,'./Fullerton_New Cases.xlsx' ,sep="")

Fullerton_New_validation <-read_excel(Fullerton_New)


Fullerton_New_validation$Lead_Check <- paste(Fullerton_New_validation$CISserialno,"/FULLERTON/",Fullerton_New_validation$Account_No,sep = '')

Fullerton_New1 <- dump_wh %>% filter(dump_wh$Lead_Check_Wh %in% Fullerton_New_validation$Lead_Check)

c<-sum(complete.cases(Fullerton_New_validation$Lead_Check))
d<-sum(complete.cases(Fullerton_New1$Lead_Check_Wh))



  

#SBI Validation


SBI <- paste(Sanker_Folder,'./SBI Cards_New Cases.xlsx' ,sep="")

SBI_validation <-read_excel(SBI)


SBI_validation$Lead_Check <- paste(SBI_validation$CISserialno,"/SBI Cards/",SBI_validation$Account_No,sep = '')

SBI1 <- dump_wh %>% filter(dump_wh$Lead_Check_Wh %in% SBI_validation$Lead_Check)

e<-sum(complete.cases(SBI_validation$Lead_Check))
f<-sum(complete.cases(SBI1$Lead_Check_Wh))



#HDFC_CC Validation


HDFC_CC <- paste(Sanker_Folder,'./HDFC Bank_CC_New Cases.xlsx' ,sep="")

HDFC_CC_validation <-read_excel(HDFC_CC)


HDFC_CC_validation$Lead_Check <- paste(HDFC_CC_validation$CISserialno,"/HDFC/",HDFC_CC_validation$Account_No,sep = '')

HDFC_CC1 <- dump_wh %>% filter(dump_wh$Lead_Check_Wh %in% HDFC_CC_validation$Lead_Check)

g<-sum(complete.cases(HDFC_CC_validation$Lead_Check))
h<-sum(complete.cases(HDFC_CC1$Lead_Check_Wh))



#HDFC_Retail Validation


HDFC_Retail <- paste(Sanker_Folder,'./HDFC Bank_RETAIL_New Cases.xlsx' ,sep="")

HDFC_Retail_validation <-read_excel(HDFC_Retail)


HDFC_Retail_validation$Lead_Check <- paste(HDFC_Retail_validation$CISserialno,"/HDFC/",HDFC_Retail_validation$Account_No,sep = '')

HDFC_Retail1 <- dump_wh %>% filter(dump_wh$Lead_Check_Wh %in% HDFC_Retail_validation$Lead_Check)

i<-sum(complete.cases(HDFC_Retail_validation$Lead_Check))
j<-sum(complete.cases(HDFC_Retail1$Lead_Check_Wh))
sku <- list()


if (c==d) {
  append(sku, "Fullerton Batch Correctly Generated")
}else{
  
  append(sku, "Fullerton Batch Wrongly Generated")
}


if (e==f) {
  append(sku, "SBI Batch Correctly Generated")
}else{
  
  append(sku, "SBI Batch Wrongly Generated")
}


if (g==h) {
  append(sku, "HDFC_CC Batch Correctly Generated")
}else{
  
  append(sku, "HDFC_CC Batch Wrongly Generated")
}



if (i==j) {
  append(sku, "HDFC_Retail Batch Correctly Generated")
}else{
  
  append(sku, "HDFC_Retail Batch Wrongly Generated")
}











