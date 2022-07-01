rm(list=ls())
# setwd('Z:\\Amrish\\1_AppOps_Ref\\1_AppOppsNew')
setwd(Sys.getenv('AppOpsNew'))
today <- Sys.Date()
source('.\\Source\\Fuction_Files.R')

#load Library
library(data.table)
library(readxl)
library(mailR)
library(xtable)
library(dplyr)
library(stringi)
library(stringr)

#Send Sudharsan Mail
dm <- fread('.\\Input\\DownloadMaster.csv')
dm.product <- as.data.frame(dm %>% group_by(Product_Name) %>% summarise( count = n()))
dm.product[nrow(dm.product)+1,2] = sum(dm.product$count)
dm.product[nrow(dm.product),1] = 'Total'

SendMail(
  dm.product,
  "Send to lender",
  c("r.sudarshan@creditmantri.com"),
  c("ranjit.punja@creditmantri.com","rupa@creditmantri.com","Bhalakumaran@creditmantri.com","upasna.batra@creditmantri.com","Manpreet@creditmantri.com","Bhuvanesh.v@creditmantri.com","Pavan.vikas@creditmantri.com","samyuktha.g@creditmantri.com","Abhinav.Priyadarshi@creditmantri.com", "referrals@creditmantri.com")
)

#Sending New Leads to all lender -----------------------------------------------
#Load Mail list
mail_list <- '.\\Input\\Mail List\\MailMasterNew.xlsx'
mail_Others <- read_excel(mail_list, sheet = "Others")
mail_indusind_pl <- read_excel(mail_list, sheet = 'indusind_pl')
mail_hdb <- read_excel(mail_list, sheet = 'HDB')
mail_hdfc_bl <- read_excel(mail_list, sheet = 'HDFCBLSPOC')
mail_hdfc_pl <- read_excel(mail_list, sheet = 'HDFCPLSPOC')
mail_bob <- read_excel(mail_list, sheet = 'BOB')

mail <- list(mail_Others,
             mail_indusind_pl,
             mail_hdb,
             mail_hdfc_bl,
             mail_hdfc_pl,
             mail_bob)

# mail <- mail_hdfc_pl

lapply(mail, function(x){
  
  x <- data.frame(x)
  lender <- x$File
  
  #browser()
  lapply(lender, function(y){
    
    #browser()
    row_set <- x[grepl(y, x$File),]
    sender <- 'referrals@creditmantri.com'
    # recepients <- row_set$SendId
    recepients <- stringr::str_split(row_set$SendId, ",\\s*")[[1]]
    myMessage <- paste(row_set$Subject,today)
    msg <- paste('Hi Team,<br><br>Please find attached the list of new referrals for ',today,'.<br><br>Regards<br>CreditMantri Referral Team',sep='')
    # cclist <- row_set$CC 
    cclist <- stringr::str_split(row_set$CC, ",\\s*")[[1]]
    filename <- paste('.\\Output\\',today,'\\Zip\\',row_set$File,sep='')
    
    # browser()
    if(file.exists(filename)){ 
      
      
      send.mail(from = sender,
                to = c(recepients),
                subject = myMessage,
                html = TRUE,
                body = msg,
                cc = c(cclist),
                smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                            user.name = "AKIAI7T5HYFCTUZMOV3Q",
                            passwd = "AtHel2jMbKGwbGlQjalkTZxEW144VM+LmgfLpNINg07E" , ssl = TRUE),
                authenticate = TRUE,
                attach.files = filename,
                send = TRUE)
    }
    
  })
  
  
  
})


