#Set Basic setup
rm(list=ls())
# setwd('C:\\Users\\Balamurugan\\OneDrive - M s. CreditMantri Finserve Private Limited\\CM-Referrals\\1_AppOppsNew\\Source')
setwd(Sys.getenv('AppOpsNew'))
today = Sys.Date()

#Create Today Date Folder
dir.name <- paste('.\\Output\\',today,sep = '')
if(!dir.exists(dir.name)) dir.create(dir.name)
if(!dir.exists(paste(dir.name,"\\Excel_File",sep =''))) dir.create(paste(dir.name,"\\Excel_File",sep =''))
if(!dir.exists(paste(dir.name,"\\Zip",sep =''))) dir.create(paste(dir.name,"\\Zip",sep =''))

#Load Libraries
library(data.table)
library(readxl)
library(openxlsx)
library(sqldf)
library(stringi)
library(stringr)
library(dplyr)

source('.\\Source\\Fuction_Files.R')

#Read Required files
dm <-  fread('.\\Input\\downloadmaster.csv', na.strings = c(''))
dm <- dm[!is.na(dm$Application_Number),]
dm <- dm[grepl(250,dm$Product_Status),]
dm.lender_details <- read_excel('.\\Input\\Lender_Details.xlsx')

#Grid Files
dm.hdb <- read_excel('.\\Input\\HDB\\AppOps_download_HDB.xlsx')
dm.hdfc_pl <- read_excel('.\\Input\\HDFC_PL\\AppOps_download_HDFC_PL.xlsx')


#Create Standard Temp Files ----------------------------------------------------
Temp.lenders <- dm.lender_details[grepl('Temp',dm.lender_details$Templete),]

if(nrow(Temp.lenders)>0){
  
dir.others <- paste(dir.name,'\\Excel_File\\Others',sep ='')
if(!dir.exists(dir.others)) dir.create(dir.others)
 

lapply(Temp.lenders$Product, function(x){
  
  x.product_status <- dm[grepl(x, dm$Product_Status),]
  
  
  # browser()
  lapply(Temp.lenders$Lender, function(y){

    y.product <- x.product_status[grepl(y,x.product_status$Product_Name),]
    
    lender_name <- unique(y.product$Product_Name)
    
    NewLeads <- Temp(y.product)
    FilePath <- paste(dir.name,'\\Excel_File\\Others\\',lender_name,'.xlsx',sep = '')
    
    if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)
    
    
  })  
  
})

}

#IndusInd CC -------------------------------------------------------------------
dm.indusind_cc <- read_excel('.\\Input\\INDUSIND_CC\\AppOps_download_IndusInd_CC.xlsx')
ltd.indusind_cc <- read_excel('.\\Input\\INDUSIND_CC\\Indusind_CC_New_Leads-LTD.xlsx')

if(nrow(dm.indusind_cc)>0){

  dm.indusind_cc$Date <- as.character(dm.indusind_cc$Date)  
  ltd.indusind_cc$Date <- as.character(format(ltd.indusind_cc$Date, '%d-%m-%Y'))  
  
dm.indusind_cc <- anti_join(dm.indusind_cc,ltd.indusind_cc,"CreditMantri_Application_Reference_Number")

if(nrow(dm.indusind_cc)>0){  
  
    dm.indusind_cc$`New/Old` <- 'NEW'
    ltd.indusind_cc$`New/Old` <- 'OLD'
  
    indusind_cc.final <- bind_rows(dm.indusind_cc,ltd.indusind_cc)
  
    NewLeads <- indusind_cc.final
    NewLeads$S.NO <- 1:nrow(NewLeads)
    
    #save it in Output
    FilePath <- paste(dir.name,'\\Excel_File\\Others\\Indusind_CC_New_Leads-LTD.xlsx',sep = '')
    if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)
    
    #save it in Input
    FilePath <- '.\\Input\\INDUSIND_CC\\Indusind_CC_New_Leads-LTD.xlsx'
    if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)
}

}

#HDFC HL/HLBT/LAP ---------------------------------------------------------------
#load download master file
DM <- dm
DM$Phone_No <- as.character(DM$Phone_No)
HL_Master <- read_excel('.\\Input\\HDFC_HL_master.xlsx')

#Query for required template
HDFC_HL_Query <- "SELECT 
                        'credit Mantri- HL' AS 'Promo Code',
                        HL_Master.CRM_Branch_Code as 'Branch Code',
                        lower(DM.Current_Residence_City) AS 'Branch Name',
                        NULL AS 'Shadow LG Code',
                        NULL AS 'Shadow LG Name',
                        NULL AS 'Shadow LC Code',
                        DM.Customer_Segment AS 'Profile',
                        DM.Customer_Name AS 'Customer Name',
                        NULL AS 'Customer ID',
                        NULL AS 'Ref. Account No.',
                        NULL AS 'Customer Band',
                        DM.Phone_No AS 'PhoneNo',
                        'Marketing' AS 'Lead Source',
                        NULL AS 'Lead Priority',
                        NULL AS 'LC Code',
                        DM.Product_Name AS 'Product',
                        DM.Loan_Amount_Required AS 'Estimated Amount',
                        NULL AS 'Customer Category',
                        (CASE
                            WHEN DM.Customer_Segment = 'Salaried' THEN DM.Net_Take_Home_Per_Month
                            WHEN DM.Customer_Segment = 'Self employed business' THEN DM.Monthly_Commitments
                            ELSE 0 END) AS 'Net Salary_Net Profit',
                        NULL AS 'No. of years in Employment_business',
                        NULL AS 'Company',
                        NULL AS 'Documents Collected',
                        NULL AS 'Rate',
                        NULL AS 'Alternate address',
                        NULL AS 'Alternate Phone 1',
                        NULL AS 'Alternate Phone 2',
                        NULL AS 'Alternate phone 3',
                        NULL AS 'Alternate email id',
                        NULL AS 'Comments',
                        DM.Application_Number AS 'Remarks 1',
                        NULL AS 'Remarks 2',
                        NULL AS 'Remarks 3',
                        NULL AS 'Remarks 4',
                        NULL AS 'Remarks 5'
                  FROM DM LEFT JOIN HL_Master ON 
                        lower(DM.Current_Residence_City) = HL_Master.Location
                  WHERE Product_Name == 'HDFC Bank Home Loan' OR
                        Product_Name == 'HDFC Bank Loan Against Property' OR
                        Product_Name == 'HDFC Housing Finance Ltd Home Loan'"

#Fetch data from DM by using Query
HDFC_HL <- sqldf(HDFC_HL_Query)

if(nrow(HDFC_HL)>0){
  
NewLeads = HDFC_HL %>% distinct(PhoneNo,.keep_all = T)
FilePath <- paste(dir.name,'\\Excel_File\\Others\\HDFC_HomeLoan.xlsx',sep = '')
if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)

}

#HDB Leads ---------------------------------------------------------------------
dm.hdb <- read_excel('.\\Input\\HDB\\AppOps_download_HDB.xlsx')
pincode.hdb <- fread('.\\Input\\HDB\\HDBPincodecitymapper.csv',na.strings = c(''))

if(nrow(dm.hdb)>0){
  
dm.hdb$Residence_Pincode <- as.character(dm.hdb$Residence_Pincode)
pincode.hdb$Pincode <- as.character(pincode.hdb$Pincode)
pincode.hdb <- pincode.hdb %>% group_by(Pincode) %>% slice(1) %>% ungroup()

hdb <- inner_join(dm.hdb,pincode.hdb,by = c('Residence_Pincode' = 'Pincode'))

hdb$Current_Residence_City <- hdb$City
hdb$City = NULL

dir.hdb <- paste(dir.name,'\\Excel_File\\HDB',sep = '')
if(!dir.exists(dir.hdb)) dir.create(dir.hdb)

hdb_city <- unique(hdb$Current_Residence_City)

lapply(hdb_city, function(x){
  
  NewLeads <- hdb[grepl(x, hdb$Current_Residence_City),]
  NewLeads$S.NO <- 1:nrow(NewLeads)
  FilePath <- paste(dir.hdb,'\\HDBNew-',x,'.xlsx',sep ='')
  if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)
  
  
})

}

#IndusInd PL -------------------------------------------------------------------
spoc.indusind_pl <- fread('.\\Input\\IndusIndSPOCcityMaster.csv')
dm.indusind_pl <- dm[grepl('IndusInd Bank Personal Loan',dm$Product_Name),]

if(nrow(dm.indusind_pl)>0){
  
spoc.indusind_pl$Location <- tolower(spoc.indusind_pl$Location)
dm.indusind_pl$Current_Residence_City <- tolower(dm.indusind_pl$Current_Residence_City)

indusind_pl <- left_join(dm.indusind_pl,spoc.indusind_pl,
                         by = c('Current_Residence_City' = 'Location')) 
indusind_pl <- indusind_pl %>% distinct(Application_Number,.keep_all = T)

#Save all Leads
dir.indusind_pl <- paste(dir.name,'\\Excel_File\\IndusInd_PL',sep='')
if(!dir.exists(dir.indusind_pl)) dir.create(dir.indusind_pl)

NewLeads <- Temp(indusind_pl)
FilePath <- paste(dir.indusind_pl,'\\IndusInd Bank Personal Loan.xlsx',sep='')
if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)

#Save spocwise leads
spoc <- unique(indusind_pl$SPOC)
spoc <- spoc[!is.na(spoc)]
lapply(spoc, function(x){
  
  Leads <- indusind_pl[grepl(x,indusind_pl$SPOC),]
  NewLeads <- Temp(Leads)
  FilePath <- paste(dir.indusind_pl,'\\',x,'.xlsx',sep='')
  if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)
  
  
})

}

#HDFC PL -----------------------------------------------------------------------
dm.hdfc_pl <- read_excel('.\\Input\\HDFC_PL\\AppOps_download_HDFC_PL.xlsx')

if(nrow(dm.hdfc_pl)>0){

dir.hdfc_pl <- paste(dir.name,'\\Excel_File\\HDFC PL',sep = '')
if(!dir.exists(dir.hdfc_pl)) dir.create(dir.hdfc_pl)

dm.hdfc_pl$BranchName <- tolower(dm.hdfc_pl$BranchName)

#Save all leads
NewLeads <- dm.hdfc_pl
FilePath <- paste(dir.hdfc_pl,'\\HDFC-consolidated-NewLeads.xlsx',sep = '')
if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)

#Split Spocwise Leads
spoc.hdfc <- unique(dm.hdfc_pl$`Branch Code`)
product.hdfc <- unique(dm.hdfc_pl$Product)


lapply(product.hdfc, function(y){
  
  NewLeads <- dm.hdfc_pl[grepl(y,dm.hdfc_pl$Product),]
  spoc.hdfc <- unique(NewLeads$`Branch Code`)
  
  lapply(spoc.hdfc, function(x){
    
    NewLeads <- NewLeads[grepl(x,NewLeads$`Branch Code`),]
    NewLeads$S.NO <- 1:nrow(NewLeads)
    
    z <- trimws(str_replace(x, 'Open-Beu',''))
    product <- unique(NewLeads$Product)
    FilePath <- paste(dir.hdfc_pl,"\\HDFC",product,'-',z,'.xlsx',sep = '')
    if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)
    
  })
  
})

}

#Yes Bank CC -------------------------------------------------------------------
dm.yes_bank <- dm[grepl('CC',dm$Product_Status),]
dm.yes_bank <- dm.yes_bank[grepl('Yes',dm.yes_bank$Product_Name),]

if(nrow(dm.yes_bank )>0){

col <- c('Application_Number','Customer_Name','Phone_No','Your_residence_landline','What_Is_Your_Office_Landline_Number',
         'Net_Take_Home_Per_Month','ITR_Take_Home_Per_Month','Product_Name','Current_Residence_City','What_Is_Your_Office_Address',
         'What_Is_Your_Residence_Address','Documents','What_date_do_you_want_to_give_an_appointment_to_the_lender',
         'What_time','Company_Name','Residence_Pincode','CC_Bank_Name','CC_Cut_Off_Limit')

col.yes_bank <- dm.yes_bank %>% select(col)

col.name <- c('CRM Lead No.','Customer Name','Mobile Number 1','Home Landline','Official landline',
              'Monthly Salary','ITR(Yearly)','Card Type','City','Office Address','Home Adress','Documents',
              'Appointment Date','Appointment details','Company Name','Pin Code',
              'CC_Bank_Name','CC_Cut_Off_Limit')

names(col.yes_bank) <- col.name

col.yes_bank$`SR No` <- 1:nrow(col.yes_bank)
col.yes_bank$`Lead Type` <- 'NA'
col.yes_bank$`Calling Date` <- Sys.Date()
col.yes_bank$`Contract No` <- 'NA'
col.yes_bank$`Mobile Number 2` <- NA
col.yes_bank$`YBL Customer` <- 'NA'
col.yes_bank$`Sourcing Type` <- 'CreditMantri'
col.yes_bank$Status <- 'NA'
col.yes_bank$`Source Code` <- 'DACM'
col.yes_bank$`Communication Address` <- col.yes_bank$`Home Adress`
col.yes_bank$`ITR(Yearly)` <- col.yes_bank$`ITR(Yearly)` * 12 

req_col <- c('SR No','Lead Type','Calling Date','CRM Lead No.','Contract No','Customer Name',
             'Mobile Number 1','Mobile Number 2','Home Landline','Official landline',
             'Monthly Salary','ITR(Yearly)','Card Type','YBL Customer','City','Office Address','Home Adress',
             'Sourcing Type','Documents','Appointment Date','Appointment details',
             'Company Name','Communication Address','Pin Code','Status','Source Code',
             'CC_Bank_Name','CC_Cut_Off_Limit')

NewLeads <- col.yes_bank %>% select(req_col)
FilePath <- paste(dir.name,'\\Excel_File\\Others\\Yes bank CC.xlsx',sep = '')

if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)

}

#HDFC EM -----------------------------------------------------------------------
dm.hdfc_em <- read_excel('.\\Input\\HDFC_EM\\AppOps_download_HDFC_EM.xlsx')
NewLeads <- dm.hdfc_em
FilePath <- paste(dir.name,'\\Excel_File\\Others\\HDFC_EM_NewLeads.xlsx',sep = '')
if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)
#Ujjivan------------------------------------------------------------------------
dm.ujjivan <- dm %>% 
  filter(grepl('Ujjivan',Product_Name, ignore.case =TRUE)) %>% 
  filter(grepl('PL',Product_Status, ignore.case =TRUE)) %>% 
  select("S.NO","Date","Application_Number","Product_Name","Customer_Name","DOB",
         "PAN","Email","Phone_No","Gender","Marital_Status",
         "Current_Residence_City","Company_Name","Total_Work_Experience",
         "Customer_Segment","Net_Take_Home_Per_Month",
         "Salary_Deposited_To_Which_Bank","Date_Of_Joining_Current_Employer",
         "Monthly_Commitments","Loan_Amount_Required","Loan_Type",
         "What_Is_Your_Residence_Address","Residence_Pincode",
         "What_Is_Your_Office_Address",
         "What_date_do_you_want_to_give_an_appointment_to_the_lender",
         "What_time","LeadId")

NewLeads <- dm.ujjivan
FilePath <- paste(dir.name,'\\Excel_File\\Others\\Ujjivan_NewLeads.xlsx',sep = '')
if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)

#BOB ---------------------------------------------------------------------------
dm.bob_query <- dm %>% 
  filter(grepl('BOB', Product_Name,ignore.case = TRUE)) %>% 
  select("S.NO","Date","Application_Number",
         "Product_Name","Customer_Name","DOB","Age","PAN","Address","Email",
         "Phone_No","Current_Residence_City","Company_Name","Total_Work_Experience",
         "Net_Take_Home_Per_Month","Annual_Turnover","Profit_After_Tax",
         "What_Is_Your_Residence_Address","Residence_Pincode",
         "What_date_do_you_want_to_give_an_appointment_to_the_lender",
         "What_time","Documents","LeadId")

dm.city <- read_excel('./Input/BOB_CityList.xlsx')



if(nrow(dm.bob_query) > 0){
  
  dir.bob <- paste(dir.name,'\\Excel_File\\BOB',sep = '')
  if(!dir.exists(dir.bob)) {dir.create(dir.bob)}
  

  
  bob_city <- dm.city$SMS
  
  lapply(bob_city, function(x){
        
        #browser()
        dm.bob <- dm.bob_query[grepl(x,dm.bob_query$Current_Residence_City, 
                                     ignore.case = T),]
        if(nrow(dm.bob)>0) dm.bob$S.NO <- 1:nrow(dm.bob)
        
        NewLeads <- dm.bob
        FilePath <- paste(dir.name,'\\Excel_File\\BOB\\BOB - ',x,'.xlsx',sep = '')
        if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)
    
    
  })
  
}

#KreditBee----------------------------------------------------------------------
dm.kreditbee <- dm %>% 
  filter(grepl('Kredit Bee Short Term Loan',Product_Name)) %>% 
  select('S.NO',"Phone_No","Email","Customer_Name","Gender","PAN",
         "DOB","Residence_Pincode","Customer_Segment",
         "Net_Take_Home_Per_Month","LeadId","Loan_Amount_Required",
         "Company_Name","Total_Work_Experience")

col_name <- c('s.no',"mobile","email","Customer_Name","gender","pan","dob",
              "pincode","profession","salary","referenceid",
              "loan amount ","Company","totalworkexp")

names(dm.kreditbee) <- col_name
if(nrow(dm.kreditbee)>0) dm.kreditbee$s.no <- 1:nrow(dm.kreditbee)

NewLeads <- dm.kreditbee
FilePath <- paste(dir.name,'\\Excel_File\\Others\\kreditbee_NewLeads.xlsx',sep = '')
if(nrow(NewLeads)>0) FileCreate <- FileCreation(NewLeads,FilePath)


#If you wanna add code, write your code code above zip file --------------------






#Zip File ----------------------------------------------------------------------
source_path <- paste('.\\Output\\',today,'\\Excel_File',sep='')
zip.file(source_path)


