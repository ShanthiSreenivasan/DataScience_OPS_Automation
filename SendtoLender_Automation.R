rm(list = ls()) # Clear environment
knitr::opts_chunk$set(comment="", echo = FALSE, message=FALSE, warning=FALSE)
library(plyr); library(dplyr)
library(car);
library(stringr);
library(mailR);library(xtable)
library(xlsx);library(lubridate)
library(excel.link)
library(sqldf)
thisDate = Sys.Date()
refDate = thisDate
thisMonth = "2016-05-03"
#setwd("F:/DailyMIS/Amrish/AppopsNew")
setwd("Z:/Amrish/AppopsNew")
if(!dir.exists(paste("./Output/",thisDate,sep=""))){
    dir.create(paste("./Output/",thisDate,sep=""))
}
dir.create(paste("./Output/",thisDate,"/OtherNew",sep=""))
dir.create(paste("./Output/",thisDate,"/IndusIndNew",sep=""))
dir.create(paste("./Output/",thisDate,"/ICICINew",sep=""))
dir.create(paste("./Output/",thisDate,"/FullertonNew",sep=""))
dir.create(paste("./Output/",thisDate,"/ShriramNew",sep=""))

#Read all input files
dfDownloadMaster=read.csv("./Input/downloadmaster.csv")
dffullertonSpoc=read.csv("./Input/FullertonSPOCFormat.csv")
dfIndusIndSpoc=read.csv("./Input/IndusIndSPOCcityMaster.csv")
dfICICIPLSpoc=read.csv("Z:/Amrish/AppopsLTD/Input/ICICI SPOC city master.csv")
dfShriramState=read.xlsx("Z:/Amrish/AppopsLTD/Input/Shriram SPOC Mapping.xlsx","PIN to State Mapping")
dfShriramSPOCPin=read.xlsx("Z:/Amrish/AppopsLTD/Input/Shriram SPOC Mapping.xlsx", sheetName ="Pincodelist")

#Filter the required data
#dfDownloadMaster=subset(dfDownloadMaster,AppOps_Status == "Send to Lender")
dfDownloadMaster$Current_Residence_City=toupper(dfDownloadMaster$Current_Residence_City)
dfDownloadMaster=subset(dfDownloadMaster,Application_Number != "")
dfDownloadMaster$RefDate=as.character(as.Date(dfDownloadMaster$Date,format = "%Y-%m-%d"))
dfDownloadMaster$What_date_do_you_want_to_give_an_appointment_to_the_lender = as.character(dfDownloadMaster$What_date_do_you_want_to_give_an_appointment_to_the_lender)
# dfDownloadMaster$What_date_do_you_want_to_give_an_appointment_to_the_lender=as.character(as.Date(dfDownloadMaster$What_date_do_you_want_to_give_an_appointment_to_the_lender,format = "%Y-%m-%d"))
#dfDownloadMaster$RefDate=as.character(thisDate)
dfDownloadMaster$What_date_do_you_want_to_give_an_appointment_to_the_lender[which(dfDownloadMaster$What_date_do_you_want_to_give_an_appointment_to_the_lender == "00-00-0000")]=""
dfICICIPL=dfDownloadMaster
#01-09-2017 adding code to send yes bank new cases in requested format
dfYESCCDM=dfDownloadMaster
# 01-09-2017 change part of it ends here check end for remaining code
#26-09-2017 change done to send CASHE leads 
dfCasheDM=dfDownloadMaster 

# taking copies for HDFC HLBT and HLLAP
dfdm <- dfDownloadMaster

dfDownloadMaster=subset(dfDownloadMaster,select=c("Date","RefDate","Application_Number","Product_Name","Product_Status","AppOps_Status","Customer_Name","Email","Phone_No","Current_Residence_City","Company_Name","Total_Work_Experience","Customer_Segment","Net_Take_Home_Per_Month","Annual_Turnover","Profit_After_Tax","Loan_Amount_Required","What_date_do_you_want_to_give_an_appointment_to_the_lender","What_time","Do_you_want_the_appointment_in_office_or_residence","Appointment_Address_for_doc_pick_up","Customer_docs_ready_for_pick_up","CMOL_Status","Tenor_Opted","EMI_Eligible","Residence_Pincode"))


#Merge Fullerton SPOC
dffullertonPL=subset(dfDownloadMaster,Product_Name == "Fullerton Personal Loan")
dffullertonSpoc$City=toupper(dffullertonSpoc$City)
dffullertonPL=merge(dffullertonPL,dffullertonSpoc,by.x="Current_Residence_City",by.y="City",all.x=TRUE)
dffullertonPL$SPOC=as.character(dffullertonPL$SPOC)
dffullertonPL$SPOC[is.na(dffullertonPL$SPOC)]="others"
dffullertonPL$SPOC=as.factor(dffullertonPL$SPOC)

#Merge IndusInd SPOC
dfIndusIndPL=subset(dfDownloadMaster,Product_Name == "IndusInd Bank Personal Loan")
dfIndusIndGL=subset(dfDownloadMaster,Product_Name == "INDUSIND Bank Gold Loan")
dfIndusIndPLFullList=dfIndusIndPL
dfIndusIndSpoc$Location=toupper(dfIndusIndSpoc$Location)
dfIndusIndPL=merge(dfIndusIndPL,dfIndusIndSpoc,by.x="Current_Residence_City",by.y="Location",all.x = TRUE)
# 
# #Merge ICICIPL SPOC
# dfICICIPL=subset(dfICICIPL,Product_Name == "ICICI Bank Personal Loan")
# dfICICIPLSpoc$Location=toupper(dfICICIPLSpoc$Location)
# dfICICIPL=merge(dfICICIPL,dfICICIPLSpoc,by.x = "Current_Residence_City",by.y = "Location",all.x = TRUE)

#create DCB files
dfDCBSCC=subset(dfDownloadMaster,Product_Name == "DCB PayLess Secured Card")
dfDCBGL=subset(dfDownloadMaster,Product_Name == "DCB Bank Gold Loan")

#Create Magma HL File
dfMagmaHL=subset(dfDownloadMaster,Product_Name == "Magma Fincorp Home Loan")
dfMagmaLAP=subset(dfDownloadMaster,Product_Name == "Magma Fincorp Loan Against Property")

#Create Sundaram Files
#28-12-2017 code added to include sundaram hlbt
dfSundaramHL=subset(dfDownloadMaster,Product_Name %in% c("Sundaram BNP Paribas Home Finance","Sundaram BNP Paribas Home Loan Balance Transfer"))
dfSundaramLAP=subset(dfDownloadMaster,Product_Name == "Sundaram BNP Paribas Loan Against Property")

#create Subham files
dfSubhamHL=subset(dfDownloadMaster,Product_Name == "SHUBHAM Home Loan")
dfSubhamLAP=subset(dfDownloadMaster,Product_Name == "SHUBHAM Loan Against Property")

#Create Shriram PL
pincodelist=dfShriramSPOCPin$PIN
TN01Pin=factor(unique(dfShriramSPOCPin$TN01PIN))
TN02Pin=factor(unique(dfShriramSPOCPin$TN02PIN))
dfShriramPL=subset(dfDownloadMaster, Product_Name == "Shriram City Union Personal Loan")
dfShriramPL$PIN=substr(dfShriramPL$Residence_Pincode, 1, 2)
dfShriramPL=merge(dfShriramPL,dfShriramState,by.x = "PIN",by.y = "State.Code",all.x = TRUE)
dfShriramPL$SPOC = as.character(dfShriramPL$SPOC)
dfShriramPL$SPOC[dfShriramPL$respincode %in% pincodelist]="Praveen"
dfShriramPL$SPOC[dfShriramPL$respincode %in% TN01Pin]="TN01"
dfShriramPL$SPOC[dfShriramPL$respincode %in% TN02Pin]="TN02"
dfShriramPL$SPOC=as.character(dfShriramPL$SPOC)
dfShriramPL$SPOC[is.na(dfShriramPL$SPOC)]="others"

#Create Bajaj PL
dfBajajPL=subset(dfDownloadMaster, Product_Name == "Bajaj Finance Personal Loan")




######################################################################################################################

#Function 1
GenerateFile = function(df,fname,sku){
  print(fname)
  dffinal = subset(df,select=c("RefDate","Application_Number","Customer_Name","Email","Phone_No","Current_Residence_City","Loan_Amount_Required","What_date_do_you_want_to_give_an_appointment_to_the_lender","Net_Take_Home_Per_Month","Profit_After_Tax"))
  names(dffinal)=c("Date","ApplicationNumber","CustomerName","Email","Phone","Location","LoanAmountRequired","AppointmentDate","NetSalary","ProfitAfterTax")

   
  # ##########################################
  
  wb = createWorkbook(type = "xls")
  sheet1 = createSheet(wb, sheetName = "New")
  
  thisDate = as.character(thisDate)
  
  dfT1 = dffinal
  #dfT2$date_of_referral = as.Date(dfT2$date_of_referral)
  
  getMySheet = function(df, sheet, sheetName) {
    myBorder = Border(color = "black", 
                      position = c("TOP","BOTTOM","LEFT","RIGHT"),
                      pen = c("BORDER_THIN"))
    
    # Create various cell styles
    cs1 = CellStyle(wb) + Font(wb, isItalic = TRUE) # rowcolumns
    cs2 = CellStyle(wb) 
    #+ Font(wb, color = "darkgrey")
    cs3 = CellStyle(wb) + Font(wb, isBold = TRUE) + myBorder +
      Fill(foregroundColor = "green")# header
    cs4 = CellStyle(wb) + myBorder
    #+ Font(wb, color = "darkgrey") 
    
    
    # Declare rows
    rows1  = createRow(sheet, rowIndex = 1:4)
    
    # Declare Cells for inserting textinfo and date
    cell1.1 = createCell(rows1, colIndex = 1:5)[[2,3]]
    cell1.2 = createCell(rows1, colIndex = 1:5)[[2,4]]
    
    # Combine cell values and info text
    setCellValue(cell1.1, sheetName)
    setCellValue(cell1.2, thisDate)
    
    
    # For the dataframe create a list
    dfColIndex = rep(list(cs4), dim(df)[2])
    names(dfColIndex) = seq(1, dim(df)[2], by = 1)
    
    
    # add the data frame
    addDataFrame(df, sheet,
                 row.names = FALSE,
                 startRow = 6,
                 startColumn = 2,
                 colnamesStyle = cs3,
                 rownamesStyle = cs1,
                 colStyle = dfColIndex)
    
    autoSizeColumn(sheet, 1:(dim(df)[2]+3))
    
    setCellStyle(cell1.1, cs3)
    setCellStyle(cell1.2, cs3)
    
  }
  
  getMySheet(dfT1, sheet1, "New")
  
  #fileName = paste("AppOps-Conversion", thisDate, ".xlsx")
  saveWorkbook(wb, fname)
  
  if(!sku %in% c("Fullerton Personal Loan","DCB PayLess Secured Card","DCB Bank Gold Loan")){
  xls=xl.get.excel() 
  xl.workbook.open(fname,password = "")
  xl.workbook.save(fname,password = "cm@1234")
  xl.workbook.close()
  
  }
  
}
###########################################################
#Generate Formatted Files#
##########################

GenerateFile2 = function(df,fname,sku){
  print(fname)
  dffinal = df
   
  # ##########################################
  
  wb = createWorkbook(type = "xls")
  sheet1 = createSheet(wb, sheetName = "New")
  
  thisDate = as.character(thisDate)
  
  dfT1 = dffinal
  #dfT2$date_of_referral = as.Date(dfT2$date_of_referral)
  
  getMySheet = function(df, sheet, sheetName) {
    myBorder = Border(color = "black", 
                      position = c("TOP","BOTTOM","LEFT","RIGHT"),
                      pen = c("BORDER_THIN"))
    
    # Create various cell styles
    cs1 = CellStyle(wb) + Font(wb, isItalic = TRUE) # rowcolumns
    cs2 = CellStyle(wb) 
    #+ Font(wb, color = "darkgrey")
    cs3 = CellStyle(wb) + Font(wb, isBold = TRUE) + myBorder
    # header
    cs4 = CellStyle(wb) + myBorder
    #+ Font(wb, color = "darkgrey") 
    
    
    # Declare rows
    rows1  = createRow(sheet, rowIndex = 1:4)
    
    # Declare Cells for inserting textinfo and date
    cell1.1 = createCell(rows1, colIndex = 1:5)[[2,3]]
    cell1.2 = createCell(rows1, colIndex = 1:5)[[2,4]]
    
    # Combine cell values and info text
    setCellValue(cell1.1, sheetName)
    setCellValue(cell1.2, thisDate)
    
    
    # For the dataframe create a list
    dfColIndex = rep(list(cs4), dim(df)[2])
    names(dfColIndex) = seq(1, dim(df)[2], by = 1)
    
    
    # add the data frame
    addDataFrame(df, sheet,
                 row.names = FALSE,
                 startRow = 1,
                 startColumn = 1,
                 colnamesStyle = cs3,
                 rownamesStyle = cs1,
                 colStyle = dfColIndex)
     
    autoSizeColumn(sheet, 1:(dim(df)[2]+13))
    
    setCellStyle(cell1.1, cs3)
    setCellStyle(cell1.2, cs3)
    
  }
  
  getMySheet(dfT1, sheet1, "New")
  
  #fileName = paste("AppOps-Conversion", thisDate, ".xlsx")
  saveWorkbook(wb, fname)
  
  if(!sku %in% c("ICICI Bank Personal Loan")){
  
  xls=xl.get.excel() 
  xl.workbook.open(fname,password = "")
  xl.workbook.save(fname,password = "cm@1234")
  xl.workbook.close()
  
  }
}



################################################################################################
#27-09-2017  - adding a new function to generate files by passing required field names

GenerateFilewithRequiredColumns = function(df,fname,sku,fields){
  print(fname)
  
  
  dffinal=df %>% select_(.dots=fields)
  
  # ##########################################
  
  wb = createWorkbook(type = "xls")
  sheet1 = createSheet(wb, sheetName = "New")
  
  thisDate = as.character(thisDate)
  
  dfT1 = dffinal
  
  
  getMySheet = function(df, sheet, sheetName) {
    myBorder = Border(color = "black", 
                      position = c("TOP","BOTTOM","LEFT","RIGHT"),
                      pen = c("BORDER_THIN"))
    
    # Create various cell styles
    cs1 = CellStyle(wb) + Font(wb, isItalic = TRUE) # rowcolumns
    cs2 = CellStyle(wb) 
    #+ Font(wb, color = "darkgrey")
    cs3 = CellStyle(wb) + Font(wb, isBold = TRUE) + myBorder +
      Fill(foregroundColor = "green")# header
    cs4 = CellStyle(wb) + myBorder
    #+ Font(wb, color = "darkgrey") 
    
    
    # Declare rows
    rows1  = createRow(sheet, rowIndex = 1:4)
    
    # Declare Cells for inserting textinfo and date
    cell1.1 = createCell(rows1, colIndex = 1:5)[[2,3]]
    cell1.2 = createCell(rows1, colIndex = 1:5)[[2,4]]
    
    # Combine cell values and info text
    setCellValue(cell1.1, sheetName)
    setCellValue(cell1.2, thisDate)
    
    
    # For the dataframe create a list
    dfColIndex = rep(list(cs4), dim(df)[2])
    names(dfColIndex) = seq(1, dim(df)[2], by = 1)
    
    
    # add the data frame
    addDataFrame(df, sheet,
                 row.names = FALSE,
                 startRow = 6,
                 startColumn = 2,
                 colnamesStyle = cs3,
                 rownamesStyle = cs1,
                 colStyle = dfColIndex)
    
    autoSizeColumn(sheet, 1:(dim(df)[2]+3))
    
    setCellStyle(cell1.1, cs3)
    setCellStyle(cell1.2, cs3)
    
  }
  
  getMySheet(dfT1, sheet1, "New")
  
  
  saveWorkbook(wb, fname)
  
  
}




####################################################





#Generate ICICIPL SPOCS
# ICICIPLSPOCS=unique(dfICICIPLSpoc$SKU.Category)
# for(i in ICICIPLSPOCS){
#   
#   m=subset(dfICICIPL,  SKU.Category == i)
#   
#   g=paste( "./ICICINew/", i ,".xls",sep = "")
#   
#   if(nrow(m)>0){
#     
#     dffinal1 = GenerateFile(m,g,"ICICI Bank Personal Loan")
#     
#   }
#   
# }  

#Generate Fullerton SPOCS
FullertonSPOCS = unique(dffullertonSpoc$SPOC)
m = subset(dffullertonPL,SPOC == "others")
dffinal1 = GenerateFile(m,paste("./Output/",thisDate,"/FullertonNew/others.xls",sep=""),"Fullerton Personal Loan")
for(i in FullertonSPOCS){
  
  m=subset(dffullertonPL,  SPOC == i)
  
  g=paste("./Output/",thisDate,"/FullertonNew/",i,".xls",sep="")
  
  if(nrow(m)>0){
    
    #dffinal1 = GenerateFile(m,g,"Fullerton Personal Loan")
    
  }
  
}  

#Generate IndusInd SPOCS
if(nrow(dfIndusIndPLFullList)>0){
  dffinal1 = GenerateFile(dfIndusIndPLFullList,paste("./Output/",thisDate,"/IndusIndNew/IndusInd Bank Personal Loan.xls",sep=""),"IndusInd Bank Personal Loan")
}

  #GL
  if(nrow(dfIndusIndGL)>0){
  dffinal1 = GenerateFile(dfIndusIndGL,paste("./Output/",thisDate,"/IndusIndNew/IndusInd Bank Gold Loan.xls",sep=""),"IndusInd Bank Gold Loan")
}

IndusIndSPOCS=unique(dfIndusIndSpoc$SPOC)
for(i in IndusIndSPOCS){
  
  m=subset(dfIndusIndPL,  SPOC == i)
  g=paste( "./Output/",thisDate,"/IndusIndNew/", i ,".xls",sep = "")
  
  if(nrow(m)>0){
    
    dffinal1 = GenerateFile(m,g,"IndusInd Bank Personal Loan")
    
  }
  
}  

#Shriram PL
ShriramSpocList=unique(dfShriramPL$SPOC)
for(i in ShriramSpocList){
  m=subset(dfShriramPL, SPOC == i)
  g=paste( "./Output/",thisDate,"/ShriramNew/", i ,".xls",sep = "")
  if(nrow(m)>0){
    dffinal1=GenerateFile(m,g,"Shriram City Union Personal Loan")
  }
}

#Generate Files
  
#DCB

if(nrow(dfDCBSCC)>0){
  dffinal1 = GenerateFile(dfDCBSCC,paste("./Output/",thisDate,"/OtherNew/DCB PayLess Secured Card.xls",sep=""),"DCB PayLess Secured Card") 
}

if(nrow(dfDCBGL)>0){
  dffinal1 = GenerateFile(dfDCBGL,paste("./Output/",thisDate,"/OtherNew/DCB Bank Gold Loan.xls",sep = ""),"DCB Bank Gold Loan")
}



#Magma
if(nrow(dfMagmaHL)>0){
  dffinal1 = GenerateFile(dfMagmaHL,paste("./Output/",thisDate,"/OtherNew/Magma Fincorp Home Loan.xls",sep =""),"Magma Fincorp Home Loan")
}

if(nrow(dfMagmaLAP)>0){
  dffinal1 = GenerateFile(dfMagmaLAP,paste("./Output/",thisDate,"/OtherNew/Magma Fincorp Loan Against Property.xls",sep = ""),"Magma Fincorp Loan Against Property")
}

#Sundaram

if(nrow(dfSundaramHL)>0){    
  dffinal1 = GenerateFile(dfSundaramHL,paste("./Output/",thisDate,"/OtherNew/Sundaram BNP Paribas Home Finance.xls",sep = ""),"Sundaram BNP Paribas Home Finance")
}
if(nrow(dfSundaramLAP)>0){ 
  dffinal1 = GenerateFile(dfSundaramLAP,paste("./Output/",thisDate,"/OtherNew/Sundaram BNP Paribas Loan Against Property.xls",sep = ""),"Sundaram BNP Paribas Loan Against Property")
}

#Subham
if(nrow(dfSubhamHL)>0){
  dffinal1 = GenerateFile(dfSubhamHL,paste("./Output/",thisDate,"/OtherNew/SHUBHAM Home Loan.xls",sep=""),"SHUBHAM Home Loan")
}

if(nrow(dfSubhamLAP)>0){
  dffinal1 = GenerateFile(dfSubhamLAP,paste("./Output/",thisDate,"/OtherNew/SHUBHAM Loan Against Property.xls",sep = ""),"SHUBHAM Loan Against Property")
}

#Shriram PL
if(nrow(dfShriramPL)>0){
  dffinal1 = GenerateFile(dfShriramPL,paste("./Output/",thisDate,"/OtherNew/Shriram City Union Personal Loan.xls",sep = ""),"Shriram City Union Personal Loan")
}

#Lending Kart
dfLendingKart=subset(dfDownloadMaster,Product_Name == "Lending Kart Business Loan")
#LendingKart=subset(NEWLB,select = c("Date","Current_Residence_City","Application_Number","Customer_Name","Phone_No","Email","Product_Status","Loan_Amount_Required","Product_Name","Tenor_Opted","Annual_Turnover"))
if(nrow(dfLendingKart)>0){
GenerateFile(dfLendingKart,paste("./Output/",thisDate,"/OtherNew/Lending Kart Business Loan.xls",sep = ""),"Lending Kart Business Loan")
}

dfLTHL=subset(dfDownloadMaster,Product_Name == "L&T Finance Home Loan")
#LTNEW=subset(NEWLT, select = c("Date","Application_Number","Product_Name","Customer_Name","DOB","PAN","Email","Phone_No","Current_Residence_City","Annual_Turnover","Loan_Amount_Required","Tenor_Opted","Residence_Pincode"))
if(nrow(dfLTHL)>0){
GenerateFile(dfLTHL,paste("./Output/",thisDate,"/OtherNew/L&T Finance Home Loan.xls",sep = ""),"L&T Finance Home Loan")
}

dfIndiaBullsHL=subset(dfDownloadMaster,Product_Name == "India Bulls Home Loan")
#IndiaHL=subset(Newindia,select = c("Customer_Name","DOB","Email","Phone_No","Current_Residence_City","Loan_Amount_Required","Company_Name","Customer_Segment","Net_Take_Home_Per_Month","Residence_Pincode"))
if(nrow(dfIndiaBullsHL)>0){
GenerateFile(dfIndiaBullsHL,paste("./Output/",thisDate,"/OtherNew/India Bulls Home Loan.xls",sep = ""),"India Bulls Home Loan")
}

dfIndiaBullsLAP=subset(dfDownloadMaster,Product_Name == "India Bulls Loan Against Property")
#IndiaLAP=subset(Newindia1,select = c("Customer_Name","DOB","Email","Phone_No","Current_Residence_City","Loan_Amount_Required","Company_Name","Customer_Segment","Net_Take_Home_Per_Month","Residence_Pincode"))
if(nrow(dfIndiaBullsLAP)>0){
GenerateFile(dfIndiaBullsLAP,paste("./Output/",thisDate,"/OtherNew/India Bulls Loan Against Property.xls",sep = ""),"India Bulls Loan Against Property")
}

dfLTLAP=subset(dfDownloadMaster,Product_Name == "L&T Finance Loan Against Property")
#lapnew=subset(newlap,select=c("Date","Application_Number","Product_Name","Customer_Name","DOB","PAN","Email","Phone_No","Current_Residence_City","Annual_Turnover","Loan_Amount_Required","Tenor_Opted","Residence_Pincode"))
if(nrow(dfLTLAP)>0){
GenerateFile(dfLTLAP,paste("./Output/",thisDate,"/OtherNew/L&T Finance Loan Against Property.xls",sep = ""),"L&T Finance Loan Against Property")
}

dfBajajPL=subset(dfDownloadMaster,Product_Name == "Bajaj Finance Personal Loan")
#Bajanew=subset(Baja,select = c("Date","Application_Number","Customer_Name","Phone_No","Current_Residence_City","LINKUP"))
if(nrow(dfBajajPL)>0){
GenerateFile(dfBajajPL,paste("./Output/",thisDate,"/OtherNew/Bajaj Finance Personal Loan.xls",sep = ""),"Bajaj Finance Personal Loan")
}

dfTataPL=subset(dfDownloadMaster,Product_Name == "Tata Capital Personal Loan")
#Tatapl=subset(Tata,select = c("Date","Application_Number","Customer_Name","Email","Phone_No","Current_Residence_City","Net_Take_Home_Per_Month","Profit_After_Tax","Loan_Amount_Required","LINKUP"))
if(nrow(dfTataPL)>0){
GenerateFile(dfTataPL,paste("./Output/",thisDate,"/OtherNew/Tata Capital Personal Loan.xls",sep = ""),"Tata Capital Personal Loan")
}

dfTataHL=subset(dfDownloadMaster,Product_Name == "Tata Capital Home Loan")
#tatahl2=subset(tatah,select = c("Date","Application_Number","Customer_Name","Email","Phone_No","Current_Residence_City","Net_Take_Home_Per_Month","Profit_After_Tax","Loan_Amount_Required","LINKUP"))
if(nrow(dfTataHL)>0){
GenerateFile(dfTataHL,paste("./Output/",thisDate,"/OtherNew/Tata Capital Home Loan.xls",sep = ""),"Tata Capital Home Loan")
}

dfTataLAP=subset(dfDownloadMaster,Product_Name == "TATA Capital Loan Against Property")
#tatalap=subset(tatahl1,select = c("Date","Application_Number","Customer_Name","Email","Phone_No","Current_Residence_City","Net_Take_Home_Per_Month","Profit_After_Tax","Loan_Amount_Required","LINKUP"))
if(nrow(dfTataLAP)>0){
GenerateFile(dfTataLAP,paste("./Output/",thisDate,"/OtherNew/TATA Capital Loan Against Property.xls",sep = ""),"TATA Capital Loan Against Property")
}

dfEdelweissHL=subset(dfDownloadMaster,Product_Name == "Edelweiss Home Loan")
#EdelweissHL=subset(Edelweiss,select=c("Date","Application_Number","Customer_Name","Phone_No","Current_Residence_City","Loan_Amount_Required"))
if(nrow(dfEdelweissHL)>0){
GenerateFile(dfEdelweissHL,paste("./Output/",thisDate,"/OtherNew/Edelweiss Home Loan.xls",sep = ""),"Edelweiss Home Loan")
}

dfEdelweissLAP=subset(dfDownloadMaster,Product_Name == "Edelweiss Loan Against Property")
#edelweisslap1=subset(edelweisslap,select = c("Date","Application_Number","Customer_Name","Phone_No","Current_Residence_City","Loan_Amount_Required"))
if(nrow(dfEdelweissLAP)>0){
GenerateFile(dfEdelweissLAP,paste("./Output/",thisDate,"/OtherNew/Edelweiss Loan Against Property.xls",sep = ""),"Edelweiss Loan Against Property")
}

dfHDFCGL=subset(dfDownloadMaster,Product_Name == "HDFC Bank Gold Loan")
#hdfcgl1=subset(hdfcgl,select = c("Date","Application_Number","Customer_Name","Phone_No","Current_Residence_City"))
if(nrow(dfHDFCGL)>0){
GenerateFile(dfHDFCGL,paste("./Output/",thisDate,"/OtherNew/HDFC Bank Gold Loan.xls",sep = ""),"HDFC Bank Gold Loan")
}

# 31-08-2017 - chanegs made to convert pnb into pnb hlbt since pnb went like in api and we need hlbt as email
dfPNBHLBT=subset(dfDownloadMaster,Product_Name == "PNB Housing Finance Ltd Home Loan" & Product_Status == "HLBT-250")
#PNBHL1=subset(PNBHL,select =c("Date","Application_Number","Product_Name","Product_Status","Customer_Name","DOB","PAN","Email","Phone_No","Current_Residence_City","Company_Name","Total_Work_Experience","Customer_Segment","Net_Take_Home_Per_Month","Loan_Amount_Required"))
if(nrow(dfPNBHLBT)>0){
GenerateFile(dfPNBHLBT,paste("./Output/",thisDate,"/OtherNew/PNB Housing Finance Ltd Home Loan-HLBT.xls",sep = ""),"PNB Housing Finance Ltd Home Loan-HLBT")
}

dfJanaBL=subset(dfDownloadMaster, Product_Name == "Janalakshmi Financial Services Ltd Business Loan")
#janabl1=subset(JanaBL,select = c("Date","Application_Number","Customer_Name","DOB","PAN","Email","Phone_No","Current_Residence_City","Customer_Segment","Annual_Turnover","Loan_Amount_Required","Tenor_Opted"))
if(nrow(dfJanaBL)>0){
GenerateFile(dfJanaBL,paste("./Output/",thisDate,"/OtherNew/Janalakshmi Financial Services Ltd Business Loan.xls",sep = ""),"Janalakshmi Financial Services Ltd Business Loan")
}

dfArogyaPL=subset(dfDownloadMaster, Product_Name == "Arogya Finance Personal Loan")
#Arogya1=subset(Arogya,select = c("Application_Number","Customer_Name","Email","Phone_No","Current_Residence_City","Company_Name"))
if(nrow(dfArogyaPL)>0){
GenerateFile(dfArogyaPL,paste("./Output/",thisDate,"/OtherNew/Arogya Finance Personal Loan.xls",sep = ""),"Arogya Finance Personal Loan")
}
dfIDFCPL=subset(dfDownloadMaster, Product_Name == "IDFC Personal Loan")
#IDFCPL1=subset(IDFCPL,select = c ("Date","Application_Number","CMOL_Status","Product_Name","Customer_Name","DOB","PAN","Email","Phone_No","Current_Residence_City","Company_Name","Total_Work_Experience","Customer_Segment","Company_Category","Net_Take_Home_Per_Month","Annual_Turnover","Profit_After_Tax","Annual_Commitments","Credit_Score","Total_Accounts","No_Of_Active_Accounts","No_of_Positive_Accounts","No_of_Negative_Accounts","No_Of_Delayed_payments","No_Of_Enquries_Made","Loan_Amount_Required","Tenor_Opted"))
if(nrow(dfIDFCPL)>0){
GenerateFile(dfIDFCPL,paste("./Output/",thisDate,"/OtherNew/IDFC Personal Loan.xls",sep = ""),"IDFC Personal Loan")
}


dfHFFCHL=subset(dfDownloadMaster, Product_Name == "Home First Finance Company Home Loan")
#HFFC1=subset(HFFC, select = c("Date","Application_Number","Product_Name","Customer_Name","DOB","Email","Phone_No","Current_Residence_City","Customer_Segment","Monthly_Commitments","Credit_Score","Loan_Amount_Required"))
if(nrow(dfHFFCHL)>0){
  GenerateFile(dfHFFCHL,paste("./Output/",thisDate,"/OtherNew/Home First Finance Company Home Loan.xls",sep = ""),"Home First Finance Company Home Loan")
}

dfHFFCLAP=subset(dfDownloadMaster, Product_Name == "Home First Finance Company Loan Against Property")
#HFFC1LAP=subset(HFFC2, select = c("Date","Application_Number","Product_Name","Customer_Name","DOB","Email","Phone_No","Current_Residence_City","Customer_Segment","Monthly_Commitments","Credit_Score","Loan_Amount_Required"))
if(nrow(dfHFFCLAP)>0){
  GenerateFile(dfHFFCLAP,paste("./Output/",thisDate,"/OtherNew/Home First Finance Company Loan Against Property.xls",sep = ""),"Home First Finance Company Loan Against Property")
}

dfVistaarBL=subset(dfDownloadMaster, Product_Name =="Vistaar Finance Business Loan")
#vistaarbl=subset(Vistaar, select = c("Date","Application_Number","CMOL_Status","Product_Name","Customer_Name","DOB","PAN","Email","Phone_No","Current_Residence_City","Customer_Segment","Annual_Turnover","Annual_Commitments","Credit_Score","Total_Accounts","No_Of_Active_Accounts","No_of_Positive_Accounts","No_of_Negative_Accounts","No_Of_Delayed_payments","Loan_Amount_Required","Tenor_Opted"))
if(nrow(dfVistaarBL)>0){
  GenerateFile(dfVistaarBL,paste("./Output/",thisDate,"/OtherNew/Vistaar Finance Business Loan.xls",sep = ""),"Vistaar Finance Business Loan")
} 



# yes bank new cases 01-09-2017 change continues 


dfYESCCDM=subset(dfYESCCDM, Product_Name %in% c("Yes First Preferred Credit Card","Yes Prosperity Edge Credit Card","Yes Prosperity Rewards Plus Credit Card"))
YesBankfields=c("RefDate",	"Application_Number",	"Product_Name",	"Customer_Name",	"DOB", "PAN",	"Email",	"Phone_No",	"Current_Residence_City",	"Company_Name",	"Net_Take_Home_Per_Month")
fname=paste("./Output/",thisDate,"/OtherNew/Yes Bank CC.xls",sep = "")

if(nrow(dfYESCCDM)>0)
{
  GenerateFilewithRequiredColumns(dfYESCCDM,fname,"YESBANKCC",YesBankfields)
}

# yes bank new cases change ends 

#26-12-2017 - CASHE required 

dfCasheDM =subset(dfCasheDM, Product_Name %in% c("Cashe Instant Short Term Loan"))

Cashefields=c("Application_Number",	"Customer_Name",	"Company_Name",	"Phone_No",	"Email","Current_Residence_City","PAN",	"Net_Take_Home_Per_Month")
fname=paste("./Output/",thisDate,"/OtherNew/CASHe.xls",sep = "")

if(nrow(dfCasheDM)>0)
{
  GenerateFilewithRequiredColumns(dfCasheDM,fname,"CASHE",Cashefields)
}


# 08-09-2017 adding INDUS IND CC INDUSIND Bank Cards With Score - LTD file creation - change starts here 
# this code will pick the LTD file with score from input folder and addd new cases on top and place it in output folder for the current day

options(xlsx.date.format = "yyyy-mm-dd")


# reading data from LTD dump
dfIndusIndLTD <- openxlsx::read.xlsx("./Input/Indusind_CC_New_Leads-LTD.xlsx",1,)
#colnames(dfIndusIndLTD$New.Old) <- "New/Old" 
#dfIndusIndLTD$Date <- dmy(dfIndusIndLTD$Date)
#dfIndusIndLTD$Pickup_Date <- ymd(dfIndusIndLTD$Pickup_Date)

dfIndusIndLTD$`New.Old` <- "OLD"


# reading data from input file downloaded from Grid
inputfilename=paste("./Input/AppOps_download_",format(as.Date(Sys.Date()), "%d%b%Y"),".xlsx",sep="")
outputpath=paste("./Output/",thisDate,"/OtherNew/Indusind_CC_New_Leads-LTD.xlsx",sep="")



#assing the LTD to final just in case if there is no new cases
dfIndusIndFinal=dfIndusIndLTD

#removing the records from input file if it already exist in the LTD file

if(file.exists(inputfilename)) 
{
  dfIndusIndNew=openxlsx::read.xlsx(inputfilename,1)
  
 
#dfIndusIndNew$Date <- dmy(as.character(dfIndusIndNew$Date))
#dfIndusIndNew$Pickup_Date <- ymd(as.character(dfIndusIndNew$Pickup_Date))
    dfIndusIndNew$`New.Old` = 'NEW'
    
    #Remove Duplicates
    dfFinal1 <- left_join(dfIndusIndLTD,
                           dfIndusIndNew[,c("S.NO","CreditMantri_Application_Reference_Number")],
                           by = c("CreditMantri_Application_Reference_Number" = "CreditMantri_Application_Reference_Number"))
    
    dfIndusIndLTD <- dfFinal1 %>% 
      filter(is.na(S.NO.y)) %>% 
      select(-S.NO.y)

    
    dfname <- names(dfIndusIndLTD)
    x <- str_replace(dfname , "S.NO.x","S.NO")
    names(dfIndusIndLTD) <- x
    
    #dfIndusIndFinal$New/Old[is.na(dfIndusIndFinal$New/Old)] = 'New'
    
    dfIndusIndNew$Applicant_Mobile <- as.character(dfIndusIndNew$Applicant_Mobile)
    
    dfIndusIndFinal <-bind_rows(dfIndusIndLTD,dfIndusIndNew)
    
    dfIndusIndFinal= dfIndusIndFinal %>% arrange(`New.Old`)

    dfIndusIndFinal$S.NO = 1:nrow(dfIndusIndFinal) 
  }
  
  # writing data back to the dump
  write.xlsx(dfIndusIndFinal,"./Input/Indusind_CC_New_Leads-LTD.xlsx",1,sheetName = "LTD",row.names = FALSE) 
  
  # writing data to output folder for mail to pick up
  
  write.xlsx(dfIndusIndFinal,outputpath,1,sheetName = "LTD",row.names = FALSE) 
  

# 08-09-2017 adding INDUS IND CC INDUSIND Bank Cards With Score - LTD file creation - change ends  here 



##25-10-2017 - HDFC PL NEW cases merging the code. changes starts here 



if(!dir.exists(paste("./Output/",thisDate,"/HDFCPLNew",sep=""))){
  dir.create(paste("./Output/",thisDate,"/HDFCPLNew",sep=""))
}



################################################################################################
#this function generates excel file with the required columns

GenerateFile = function(df,fname){
  print(fname)
  
  
  dffinal=df[1:37]
  
  # ##########################################
  
  wb = createWorkbook(type = "xls")
  sheet1 = createSheet(wb, sheetName = "New")
  
  thisDate = as.character(thisDate)
  
  dfT1 = dffinal
  
  
  getMySheet = function(df, sheet, sheetName) {
    myBorder = Border(color = "black", 
                      position = c("TOP","BOTTOM","LEFT","RIGHT"),
                      pen = c("BORDER_THIN"))
    
    # Create various cell styles
    cs1 = CellStyle(wb) + Font(wb, isItalic = TRUE) # rowcolumns
    cs2 = CellStyle(wb) 
    #+ Font(wb, color = "darkgrey")
    cs3 = CellStyle(wb) + Font(wb, isBold = TRUE) + myBorder +
      Fill(foregroundColor = "green")# header
    cs4 = CellStyle(wb) + myBorder
    #+ Font(wb, color = "darkgrey") 
    
    
    # Declare rows
    rows1  = createRow(sheet, rowIndex = 1:4)
    
    # Declare Cells for inserting textinfo and date
    cell1.1 = createCell(rows1, colIndex = 1:5)[[2,3]]
    cell1.2 = createCell(rows1, colIndex = 1:5)[[2,4]]
    
    # Combine cell values and info text
    setCellValue(cell1.1, sheetName)
    setCellValue(cell1.2, thisDate)
    
    
    # For the dataframe create a list
    dfColIndex = rep(list(cs4), dim(df)[2])
    names(dfColIndex) = seq(1, dim(df)[2], by = 1)
    
    
    # add the data frame
    addDataFrame(df, sheet,
                 row.names = FALSE,
                 startRow = 6,
                 startColumn = 2,
                 colnamesStyle = cs3,
                 rownamesStyle = cs1,
                 colStyle = dfColIndex)
    
    autoSizeColumn(sheet, 1:(dim(df)[2]+3))
    
    setCellStyle(cell1.1, cs3)
    setCellStyle(cell1.2, cs3)
    
  }
  
  getMySheet(dfT1, sheet1, "New")
  
  
  saveWorkbook(wb, fname)
  
  
}
####################################################################################



inputfilename=paste("./Input/HDFCPL/AppOps_download_",format(as.Date(Sys.Date()), "%d%b%Y"),".xlsx",sep="")


dfHDFCPLNEW= read.xlsx(inputfilename,1)

dfHDFCPLNEW$Branch.Code=as.character(dfHDFCPLNEW$Branch.Code)

dfHDFCPLNEW$spoc=unlist(lapply(strsplit(dfHDFCPLNEW$Branch.Code," "), `[[`, 1))


## Generating the master file to be sent 

g=paste( "./Output/",thisDate,"/HDFCPLNew/","HDFC-consolidated-NewLeads",".xls",sep = "")
if(nrow(dfHDFCPLNEW)>0){
  
  dfHDFCPLNEW$S.NO=1:nrow(dfHDFCPLNEW)
  GenerateFile(dfHDFCPLNEW,g)
}


## splitting the data frame into two based on the column Product ENW-PL / ENW-BL 

dfHDFCPL=subset(dfHDFCPLNEW,Product=="ENW-PL")
dfHDFCBL=subset(dfHDFCPLNEW,Product=="ENW-BL")

HDFCPLSpoclist=unique(dfHDFCPL$spoc)
HDFCBLSpoclist=unique(dfHDFCBL$spoc) 


## creating files for PL spoc

for(i in HDFCPLSpoclist){
  m=subset(dfHDFCPL, spoc == i)
  
  g=paste( "./Output/",thisDate,"/HDFCPLNew/","HDFCENWPL-", i ,".xls",sep = "")
  if(nrow(m)>0){
    m$S.NO=1:nrow(m)
    GenerateFile(m,g)
  }
  
}


## creating files for BL spoc

for(i in HDFCBLSpoclist){
  m=subset(dfHDFCBL, spoc == i)
  g=paste( "./Output/",thisDate,"/HDFCPLNew/","HDFCENWBL-", i,".xls",sep = "")
  if(nrow(m)>0){
    m$S.NO=1:nrow(m)
    GenerateFile(m,g)
  }
  
}


# 29-01-2017  -  HDB lead generation city wise automation starts here 
library(openxlsx)
Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")


dfHDBLeads=xlsx::read.xlsx("./Input/AppOps_download_HDB.xlsx",sheetIndex = 1)

if (nrow(dfHDBLeads) >0)
{
dfHDBLeads$Residence_Pincode = as.numeric(as.character(dfHDBLeads$Residence_Pincode ))
dfHDBPincodecitymapper = read.csv("./Input/HDBPincodecitymapper.csv")

dfHDBPincodecitymapper = dfHDBPincodecitymapper %>% group_by(Pincode) %>% slice(1) %>% ungroup()

dfworking=inner_join(dfHDBLeads,dfHDBPincodecitymapper,by = c('Residence_Pincode'='Pincode'))

dfworking$Current_Residence_City = dfworking$City

dfworking$City  = NULL

veccity=  unique(dfworking$Current_Residence_City)

hs <- openxlsx::createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize=12,fontName="Arial Narrow", fgFill = "#4F80BD",border = "TopBottomLeftRight",borderColour="black")

for (i in veccity)
{
  print (i) 
  m=subset(dfworking,Current_Residence_City==i)
  m$S.NO=1:nrow(m)
  
  wb<-createWorkbook()
  
  addWorksheet(wb,"NewLeads")
  setColWidths(wb,sheet=1,cols =1:ncol(m), widths = "auto")
  writeData(wb,"NewLeads",m,headerStyle = hs,borderColour = "black",borders = "all")
  
  fname=paste("./Output/",thisDate,"/OtherNew/HDBNew-",i,".xlsx",sep="")
  
  saveWorkbook(wb, file =fname , overwrite = TRUE)
  
}

if (nrow(dfworking)<nrow(dfHDBLeads)){
  print("Attention, HDB lead not mapped to city - Please check")
  print("Attention, HDB lead not mapped to city - Please check")
  print("Attention, HDB lead not mapped to city - Please check")
  print("Attention, HDB lead not mapped to city - Please check")
  print("Attention, HDB lead not mapped to city - Please check")
  print("Attention, HDB lead not mapped to city - Please check")
}

}
# HDB automation ends here 

# HDFC HL LAP HLBT code starts here - 21/03/2018
hs <- openxlsx::createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize=12,fontName="Arial Narrow", fgFill = "#4F80BD",border = "TopBottomLeftRight",borderColour="black")
dfhdfchl_branch <- read.xlsx("./Input/HDFC_HL_master.xlsx",1,check.names = F)

#filter only HDFC HL, LAP, HLBT Leads
dfdm <- dfdm %>% 
  filter(Product_Name %in% 
           c("HDFC Bank Home Loan",
             "HDFC Bank Loan Against Property",
             "HDFC Housing Finance Ltd Home Loan"),
         grepl("-250",Product_Status))


#Fetch required fileds in download master
hdfchl_leads <- dfdm %>% 
  select(Application_Number,
         Customer_Segment,
         Customer_Name,
         Phone_No,
         Product_Name,
         Profit_After_Tax,
         Current_Residence_City,
         Loan_Amount_Required,
         Monthly_Commitments)

hdfchl_leads <- hdfchl_leads %>% 
  mutate(Profit_After_Tax = ifelse(grepl("Self",Customer_Segment) & is.na(Profit_After_Tax),
                                   Monthly_Commitments,
                                   Profit_After_Tax))

hdfchl_leads <- hdfchl_leads %>% 
  select(Application_Number,
         Customer_Segment,
         Customer_Name,
         Phone_No,
         Product_Name,
         Profit_After_Tax,
         Current_Residence_City,
         Loan_Amount_Required)


colnames(hdfchl_leads) <- c("Application Number",
                            "Profile",
                            "Customer Name",
                            "Phone",
                            "Product",
                            "Net Salary_Net Profit",
                            "Branch Name",
                            "Estimated Amount")




hdfchl_leads <- hdfchl_leads %>% 
  mutate('Promo Code'="credit Mantri- HL",
         'Shadow LG Code'="",
         'Shadow LG Name'="",
         'Shadow LC Code'="",
         'Customer ID'="",
         'Ref. Account No.'="",
         'Customer Band'="",
         'Lead Source'="Marketing",
         'Lead Priority'="",
         'LC Code'="",
         'Customer Category'="",
         'No. of years in Employment_business'="",
         'Company'="",
         'Documents Collected'="",
         'Rate'="",
         'Alternate address'="",
         'Alternate Phone 1'="",
         'Alternate Phone 2'="",
         'Alternate phone 3'="",
         'Alternate email id'="",
         'Comments'="",
         'Remarks 1'= `Application Number`,
         'Remarks 2'="",
         'Remarks 3'="",
         'Remarks 4'="",
         'Remarks 5'=""
  )

#Lower case 
hdfchl_leads$`Branch Name` <- tolower(hdfchl_leads$`Branch Name`)

#left_Join
hdfchl_leads <- left_join(hdfchl_leads,
                          dfhdfchl_branch[,c("Location","CRM_Branch_Code")],
                          by = c("Branch Name" = "Location"))


names(hdfchl_leads)[names(hdfchl_leads) == "CRM_Branch_Code"] <- "Branch Code"

hdfchl_leads <- hdfchl_leads %>% 
  select(`Promo Code`,
         `Branch Code`,
         `Branch Name`,
         `Shadow LG Code`,
         `Shadow LG Name`,
         `Shadow LC Code`,
         `Profile`,
         `Customer Name`,
         `Customer ID`,
         `Ref. Account No.`,
         `Customer Band`,
         `Phone`,
         `Lead Source`,
         `Lead Priority`,
         `LC Code`,
         `Product`,
         `Estimated Amount`,
         `Customer Category`,
         `Net Salary_Net Profit`,
         `No. of years in Employment_business`,
         `Company`,
         `Documents Collected`,
         `Rate`,
         `Alternate address`,
         `Alternate Phone 1`,
         `Alternate Phone 2`,
         `Alternate phone 3`,
         `Alternate email id`,
         `Comments`,
         `Remarks 1`,
         `Remarks 2`,
         `Remarks 3`,
         `Remarks 4`,
         `Remarks 5`
  )


# write.csv(hdfchl_leads,"C:/Users/Balamurugan/Desktop/Referals/R/HDFC HL_LA_BL/New HDFC HL.csv")

fname=paste("./Output/",Sys.Date(),"/OtherNew/HDFCHLLAPHLBT.xlsx",sep="")
wb <- openxlsx::createWorkbook()
addWorksheet(wb,"HDFC_Leads")
setColWidths(wb,"HDFC_Leads",cols =1:ncol(hdfchl_leads),widths = "auto")
writeData(wb,"HDFC_Leads",hdfchl_leads,headerStyle = hs,borderColour = "black",borders = "all",keepNA = FALSE)
openxlsx::saveWorkbook(wb,fname,overwrite = T)



# HDFC HLBT code ends here 21/02/2018



# 09-01-2017 - adding code to send - send to lender email to sudarshan


dfDownloadMastergroup = sqldf("select Product_name,count(1) as Count_of_Application_number  from dfDownloadMaster group by Product_name")

total =  sqldf("select count(1)   from dfDownloadMaster ")

dftotal= data.frame("Product_name",total)

temp = data.frame("Product_Name"="TOTAL","Count_of_Application_number"=as.integer(total))

dfDownloadMastergroup= rbind(dfDownloadMastergroup,temp)

SendMail=function(df1,subj,recepients,cclist){
  
  myMessage = paste0(subj,"-", thisDate)
  
  sender <- "referrals@creditmantri.com"
  
  msg<- (paste('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
               <html xmlns="http://www.w3.org/1999/xhtml">
               <head>
               <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
               <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
               <style>
               p {
               font-family:  Verdana, Geneva, sans-serif;
               font-size: 13px
               }
               
               table {
               font-family:  Verdana, Geneva, sans-serif;
               font-size: 12px;
               width = 100%;
               border: 1px solid black;
               border-collapse: collapse;
               }
               
               th {
               
               padding: 5px;
               text-align: left;
               background-color: #4CAF50;
               color: white;
               width:150px;
               }
               
               td {
               background-color: #FFF;
               padding: 5px;
               text-align: left;
               width:100%;
               white-space:nowrap !important;
               }
               
               </style>
               
               </head> <body> <p> Sir,</p>
               <p> Please find the send to lender status below.<br>
               </p>',print(xtable(df1), include.rownames = FALSE,type = 'html'), '
               </p><br/>Regards<br/>Referrals Team</body> 
               </html>'))
  
  
  send.mail(from = sender,
            to = c(recepients),
            subject = myMessage,
            html = TRUE,
            body = msg,
            cc = cclist,
            bcc = c("Parivel.R@creditmantri.com"),
            smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                        user.name = "AKIAI7T5HYFCTUZMOV3Q",
                        passwd = "AtHel2jMbKGwbGlQjalkTZxEW144VM+LmgfLpNINg07E" , ssl = TRUE),
            authenticate = TRUE,
            send = TRUE)
}

####### CALLING THE EMAIL FUNCTION ########

SendMail(
  dfDownloadMastergroup,
  "Send to lender",
  c("r.sudarshan@creditmantri.com"),
  c("ranjit.punja@creditmantri.com","sarvesh.s@creditmantri.com","rupa@creditmantri.com","Bhalakumaran@creditmantri.com","upasna.batra@creditmantri.com","harish.r@creditmantri.com","Steffi@creditmantri.com","Kanchi@creditmantri.com","Parivel.R@creditmantri.com","Manpreet@creditmantri.com","Bhuvanesh.v@creditmantri.com","Pavan.vikas@creditmantri.com","samyuktha.g@creditmantri.com","Saurabh.Hirway@creditmantri.com","Abhinav.Priyadarshi@creditmantri.com","referrals@creditmantri.com")
)




## 25-10-2017 - HDFC PL code merge change ends 






#Bajaj PL

# if(nrow(dfBajajPL)>0){
#   dffinal1 = generatefile(dfBajajPL,Paste("./",thisdate,"/","OtherNew/Bajaj Finance Personal Loan.xls",sep =""),"Bajaj Finance Personal Loan")
#}

##############################################################################################################
#Generate files with predefined format#
#######################################

# #Axis HL
# dfAxisHL=subset(dfDownloadMaster,Product_Name == "Axis Bank Home Loan")
# dfAxisHL$record_id=""
# dfAxisHL$contact_info=dfAxisHL$Phone_No
# dfAxisHL$contact_info_type=""
# dfAxisHL$record_type=""
# dfAxisHL$record_status=""
# dfAxisHL$call_result=""
# dfAxisHL$attempt=""
# dfAxisHL$dial_sched_time=""
# dfAxisHL$call_time=""
# dfAxisHL$daily_from=""
# dfAxisHL$daily_till=""
# dfAxisHL$tz_dbid=""
# dfAxisHL$campaign_id=""
# dfAxisHL$agent_id=""
# dfAxisHL$chain_id=""
# dfAxisHL$chain_n=""
# dfAxisHL$group_id=""
# dfAxisHL$app_id=""
# dfAxisHL$treatments=""
# dfAxisHL$media_ref=""
# dfAxisHL$email_subject=""
# dfAxisHL$email_template_id=""
# dfAxisHL$switch_id=""
# dfAxisHL$LEAD_ID=""
# dfAxisHL$CUST_TYPE=""
# dfAxisHL$PRODUCT_NAME="Home Loan"
# dfAxisHL$CAMPAIGN_NAME="Credit Mantri"
# dfAxisHL$FREE1=dfAxisHL$Customer_Name
# dfAxisHL$FREE2=dfAxisHL$Email
# dfAxisHL$FREE3=dfAxisHL$Phone_No
# dfAxisHL$FREE4=dfAxisHL$Current_Residence_City
# dfAxisHL$FREE5=""
# dfAxisHL$FREE6="Credit Mantri"
# dfAxisHL$FREE7=""
# dfAxisHL$FREE8=""
# dfAxisHL$FREE9=""
# dfAxisHL$FREE10=""
# dfAxisHL$UPLOAD_DATE=""
# 
# dfAxisHL$record_id=1:nrow(dfAxisHL)
# 
# dfAxisHL1=subset(dfAxisHL,select = c("record_id","contact_info","contact_info_type","record_type","record_status","call_result","attempt","dial_sched_time","call_time","daily_from","daily_till","tz_dbid","campaign_id","agent_id","chain_id","chain_n","group_id","app_id","treatments","media_ref","email_subject","email_template_id","switch_id","LEAD_ID","CUST_TYPE","PRODUCT_NAME","CAMPAIGN_NAME","FREE1","FREE2","FREE3","FREE4","FREE5","FREE6","FREE7","FREE8","FREE9","FREE10","UPLOAD_DATE"))
# 
# if(nrow(dfAxisHL1)>0){
#   #dffinal1 = GenerateFile2(dfAxisHL1,paste("./",thisDate,"/","OtherNew/Axis Bank Home Loan.xls",sep = ""),"Axis Bank Home Loan")
# }

######################
#ICICI PL
#################


# dfICICIPL=subset(dfDownloadMaster,Product_Name == "ICICI Bank Personal Loan")
# dfICICIPL$State = ""
# 
#   dfICICIPL = subset(dfICICIPL,select=c("RefDate","Customer_Name","Email","Loan_Amount_Required","CMOL_Status","Customer_Segment","Net_Take_Home_Per_Month","Phone_No","What_date_do_you_want_to_give_an_appointment_to_the_lender","What_time","Do_you_want_the_appointment_in_office_or_residence","Appointment_Address_for_doc_pick_up","Current_Residence_City","State","Company_Name","Tenor_Opted","EMI_Eligible","Customer_docs_ready_for_pick_up","Application_Number"))
#     names(dfICICIPL)=c("Date","CustomerName","Email_ID","Loan_Amount","CMOL","EmploymentType","NetSalary","Mobile","DateOfAppontment","Time","PlaceOfAppointment","Address","City","State","Company Name","Tenor_Months","EMI","DocsToBePickedUp","Application_Number")
# 
#     if(nrow(dfICICIPL)>0){
#       dffinal1 = GenerateFile2(dfICICIPL,paste("./",thisDate,"/","OtherNew/ICICI Bank Personal Loan.xls",sep = ""),"ICICI Bank Personal Loan") 
#     }
  
