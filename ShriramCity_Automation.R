#Remove existing List
rm(list=ls())

#Load libraries
library(dplyr)
library(openxlsx)
require(readxl)
library(data.table)
library(sqldf)
library(nanotime)


setwd("E:\\Automation\\ShriramCity")
#File Path
#Get Dump
# dump_name <- paste(format(Sys.Date(),"%d-%m-%y"),".xlsx",sep= "")
# #dfdump <- readxl::read_excel(dump_name)
# # dfdump <- read.xlsx("28-03-18.xlsx",1)

y_date <- Sys.Date()-1
DMPath <- paste("E:\\Automation\\AppOpsDump\\appopsdump_",y_date,".csv",sep = '')
dm <- fread(DMPath, sep = ",",na.strings = NULL )

Ops <- read.xlsx(paste(".\\AppOpsCode\\AppOpsCode.xlsx"),1)
names(Ops) <- c("Ops_Status_Code","Ops_Status")

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

dfdump <- DMFinal

#Get SBPL Input File
dfin <- read_excel("SBPL_Input_File.xlsx",1)

#Data Manipulation
dfsbpl <- data.frame(dfin,stringsAsFactors = F)
dfsbpl <- dfsbpl %>%
  select(Cmidentifier,
         Remarks,
         UserRemarks) %>%
  filter(!Remarks == "",
         !Remarks =="Lead Disbursed")


dfsbpl <- inner_join(dfsbpl,
                     dfdump[,c("offer_reference_number",
                                "offer_application_number",
                                "App_Ops_State")],
                     by=c("Cmidentifier"="offer_reference_number"))

dfsbpl <- dfsbpl %>%
  filter(!offer_application_number == "")

dfsbpl <- dfsbpl %>%
  filter(App_Ops_State %in%
           c("Initial FB - Contact successful",
             "Docs in process - Security Not identified/Not Applicable",
             "Docs in process - Security identified",
             "Application Forwarded to Lender",
             "LOS pending decision", NA))


#---------------------------------------------------------------------------------------
#2018-06-13 Los Pending
dfsbpl_P0 <- dfsbpl %>%
  filter(Remarks %in%
           c("Proposal Entered",
             "Application Entry Updation",
             "PDC ACH Collected",
             "Application PDF Print Taken"),
         App_Ops_State %in%
           c("Initial FB - Contact successful",
             "Docs in process - Security Not identified/Not Applicable",
             "Docs in process - Security identified",
             "Application Forwarded to Lender",
             NA))

dfsbpl_P0$post <- "LOS pending decision"
dfsbpl_P0$Rejection_Tag <- "Nil"
dfsbpl_P0$Rejection_Category <- "Nil"


#Process1------------------------------------------------------------------------------
dfsbpl_P1 <- dfsbpl %>%
  filter(Remarks %in%
           c("Assigned to Business Executive",
             "Fixed Appointment with Customer",
             "Tele Call Made"),
         App_Ops_State %in%
           c("Application Forwarded to Lender",
             NA))

dfsbpl_P1$post <- "Initial FB - Contact successful"
dfsbpl_P1$Rejection_Tag <- "Nil"
dfsbpl_P1$Rejection_Category <- "Nil"

#Process2------------------------------------------------------------------------------

dfsbpl_P2 <- dfsbpl %>%
  filter(Remarks == "Documents Collected",
         App_Ops_State %in%
           c("Application Forwarded to Lender",
             "Initial FB - Contact successful",
             NA))

dfsbpl_P2$post <- "Docs in process - Security Not identified/Not Applicable"
dfsbpl_P2$Rejection_Tag <- "Nil"
dfsbpl_P2$Rejection_Category <- "Nil"


#---------------------------------------------------------------------------------------

#Not Interest in stage 1 Cases --------------------------------------------------------

dfsbpl_NI1 <- dfsbpl %>%
  filter(Remarks %in%
           c("Customer Not interested",
             "Loan Not Required Now"),
         App_Ops_State %in%
           c("Application Forwarded to Lender",
             "Initial FB - Contact successful",
             NA))

dfsbpl_NI1$post <- "Initial FB - NI"
dfsbpl_NI1$Rejection_Tag <- "NI"
dfsbpl_NI1$Rejection_Category <- "Rate Shopping / Enquiries in case of future need"


#Not Interest in stage 2 Cases --------------------------------------------------------

dfsbpl_NI2 <- dfsbpl %>%
  filter(Remarks %in%
           c("Customer Not interested",
             "Loan Not Required Now"),
         grepl("Docs in process",App_Ops_State))

if (nrow(dfsbpl_NI2)>0)

  {
    dfsbpl_NI2$post <- "Docs NI"
    dfsbpl_NI2$Rejection_Tag <- "NI"
    dfsbpl_NI2$Rejection_Category <-"Rate Shopping / Enquiries in case of future need"
  }

#---------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------
#Not Contactable in Stage1 Cases-------------------------------------------------------

dfsbpl_NC1 <- dfsbpl %>%
  filter(Remarks == "Customer Not Reachable",
         App_Ops_State %in%
           c("Application Forwarded to Lender",
             "Initial FB - Contact successful",
             NA))

if (nrow(dfsbpl_NC1)>0)
  {
    dfsbpl_NC1$post <- "Initial FB - NC"
    dfsbpl_NC1$Rejection_Tag <- "NC"
    dfsbpl_NC1$Rejection_Category <- "Not Contactable"
  }
#Not Contactable in Stage2 Cases-------------------------------------------------------

dfsbpl_NC2 <- dfsbpl %>%
  filter(Remarks == "Customer Not Reachable",
         grepl("Docs in process",App_Ops_State))

if (nrow(dfsbpl_NC2)>0)
  {

  dfsbpl_NC2$post <- "Docs NC"
  dfsbpl_NC2$Rejection_Tag <- "NC"
  dfsbpl_NC2$Rejection_Category <-  "Not Contactable"

  }


#----------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------
dfsbpl_r1 <- dfsbpl %>%
  filter(Remarks == "Lead Rejected",
         App_Ops_State %in%
           c("Application Forwarded to Lender",
             "Initial FB - Contact successful",
             NA))

if(nrow(dfsbpl_r1)>0)
  {
    dfsbpl_r1$post <- "AIP rejected"
    dfsbpl_r1$Rejection_Tag <- ""
    dfsbpl_r1$Rejection_Category <- ""
  }

#Rejected cases on stage 2--------------------------------------------------------------
dfsbpl_r2 <- dfsbpl %>%
  filter(Remarks == "Lead Rejected",
         grepl("Docs in process",App_Ops_State))


if(nrow(dfsbpl_r2)>0)
  {
    dfsbpl_r2$post <- "Docs stage - Rejected"
    dfsbpl_r2$Rejection_Tag <- ""
    dfsbpl_r2$Rejection_Category <- ""
}


#Rejected cases on stage 3--------------------------------------------------------------
dfsbpl_r3 <- dfsbpl %>%
  filter(Remarks == "Lead Rejected",
         grepl("LOS pending decision",App_Ops_State))


if(nrow(dfsbpl_r3)>0)
{
  dfsbpl_r3$post <- "LOS rejected"
  dfsbpl_r3$Rejection_Tag <- ""
  dfsbpl_r3$Rejection_Category <- ""
}

#Rejected cases Feedback----------------------------------------------------------------
dfsbpl_reject <- bind_rows(dfsbpl_r1,
                           dfsbpl_r2,
                           dfsbpl_r3)

#Convert into Lower Case
dfsbpl_reject$UserRemarks <- tolower(dfsbpl_reject$UserRemarks)

#Regex creation-----------------------------------------
docs_reject_regex <- "(add prf|kyc|docs|proof|proff|insufficient doc|no income doc|document|no cheque|no pay slip|no bank statement|applicant have not any chq|no address prove|cm have no op|customer not living)"
geo_location_regex <- "(geo|out of g|jio limit|ogl |out of sta|out of jeo| area|belong|not eligible living in|nearest branch|resident of|area negative|lead aasign to|lead transfer to|transfer to|customer dont|he is working)"
rented_regex <- "(rent case|surety|rented|ranted|staying|surity|residance|resi |guarantor|not residing|rent hourse|rent house|customer address|not address)"
not_contact_regex <- "(call| contact|switched off|respon|reachable|phone|ph |fake ptp|switch |wrong num|connect|not answer|not lift|not rechible)"
not_intrest_regex <- "(customer delay|not int|no need|not req|dont want loan|no requirement|not come|just info|no requirenment|customer not  interst|rate of intrest high|tenure|no  requirement|need|loan req|no  requirement|interest rate)"
cash_regex <- "(cash|hand|customer salary not deposited)"
income_regex <- "(no income source|low income|income low|low sal| ltr| iir|income below|end use of money|repayment|emi|salary low|low  income|low incom|customer not able to pay)"
default_regex <- "(default|arrear|negative pro|profile|legal case|broker|police|navy|army|risk |legal|house wife|bachlar accommodation)"
existing_regex <- "(alredy|already|dedupe|loan live|scuf|customer allready existing loan|log decctention|allready apporved loan|exsting|multiple loan|existing customer)"
foir_dbr_regex <- "(poor bank)"
duplicate_regex <- "(duplicate|double)"
cibil_regex <- "(fi |cibil|f.i|customer very low banking|customer working in|no banking|low banking|co applicant|tym to tym|not eligible|rejected|negative|many bouncing)"
policy_regex <- "(norms|noms|out of policy|lead rejected|rejected by|not available|job stability|loan forwad)"
Loan_amount <- "(overliverage|log decctention)"


#Rejection Tag-------------------------------
dfsbpl_reject <- dfsbpl_reject %>%
  mutate(
    Rejection_Tag = ifelse((grepl(docs_reject_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(geo_location_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(rented_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(not_contact_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NC",
                    ifelse((grepl(not_intrest_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NI",
                    ifelse((grepl(cash_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(income_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(default_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(existing_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(foir_dbr_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(duplicate_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(cibil_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NE",
                    ifelse((grepl(policy_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"NE",
                    ifelse( is.na(.$UserRemarks),"NE", "")))))))))))))))




#Rejection_Category--------------------------
dfsbpl_reject <- dfsbpl_reject %>%
  mutate(
    Rejection_Category = ifelse((grepl(docs_reject_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"No documents",
                         ifelse((grepl(geo_location_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"GeoLocation",
                         ifelse((grepl(rented_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"Residence Type",
                         ifelse((grepl(not_contact_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"Not Contactable",
                         ifelse((grepl(not_intrest_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"Rate Shopping / Enquiries in case of future need",
                         ifelse((grepl(cash_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"Salary Type",
                         ifelse((grepl(income_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"Income Slab",
                         ifelse((grepl(default_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"Negative customer profiles",
                         ifelse((grepl(existing_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"Existing Product Holder",
                         ifelse((grepl(foir_dbr_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"FOIR / DBR Reasons",
                         ifelse((grepl(duplicate_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"Duplicate lead",
                         ifelse((grepl(cibil_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"CIBIL Reject - Negative record in CIBIL",
                         ifelse((grepl(policy_regex,.$UserRemarks)&(!is.na(.$Rejection_Tag))),"Miscellaneous Policy",
                         ifelse( is.na(.$UserRemarks),"Miscellaneous Policy", "")))))))))))))))

#----------------------------------------------------------------------------------------

df <- bind_rows(dfsbpl_P0,
                dfsbpl_P1,
                dfsbpl_P2,
                dfsbpl_NI1,
                dfsbpl_NI2,
                dfsbpl_NC1,
                dfsbpl_NC2,
                dfsbpl_reject)


df$Concatenate <- paste(df$Cmidentifier,
                        "/",
                        df$Remarks,
                        "/",
                        df$UserRemarks)
df$Concatenate <- sub("/ NA","",df$Concatenate)

#Remove Duplicates
df <- df %>%
  distinct(offer_application_number,
           .keep_all = T)

#Arrange Columns
df_op <- df %>%
  select(Cmidentifier,
         offer_application_number,
         App_Ops_State,
         post,
         Remarks,
         UserRemarks,
         Rejection_Tag,
         Rejection_Category,
         Concatenate)


# --------------------------------------------------------------------------------------------
#Write Data in Excel
wb <- createWorkbook()
addWorksheet(wb,"SBPL_Remarks")
setColWidths(wb,"SBPL_Remarks",cols = 1:100,widths = 18)
writeData(wb,"SBPL_Remarks",df_op,rowNames = F,borders = "all")
path <- paste("SBPL_Remarks",Sys.Date(),".xlsx", sep = '')
saveWorkbook(wb,path,overwrite = T)
openXL(path)

print(dim(dfin))



