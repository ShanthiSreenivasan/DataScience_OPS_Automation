rm(list=ls())

library(lubridate)  #date manipulation
library(data.table) #High value Excel Manipulation
library(readxl)     #excel manipulation
library(openxlsx)   #excel manipulation   
library(stringr)    #string amnipulation
library(dplyr)
library(mailR)
library(htmlTable)

ydate <- format(Sys.Date(),"%Y%m%d")
fname <- paste('E:\\Leave daetails\\Leave Data 2019.xlsx',sep ='')
Leave <- read_excel(fname)

x <-Leave %>% names()

tempDF <- x
tempDF[] <- lapply(x, as.character)



#mail ---------------------------------------------------------------------------------

mail_function <- function(sender, recipients, cc_recipients, message, email_body){
  
  send.mail(from = sender,
            to = recipients,
            bcc = cc_recipients,
            subject = message,
            html = TRUE,
            inline = T,
            body = str_interp(email_body),
            smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                        user.name = "AKIAI7T5HYFCTUZMOV3Q",
                        passwd = "AtHel2jMbKGwbGlQjalkTZxEW144VM+LmgfLpNINg07E" , ssl = TRUE),
            authenticate = TRUE,
            send = TRUE)
}


#select dates-------------------------------------------

Business_Development <- Leave %>% filter(Department =="Business Development") 
Business_Development <- Business_Development %>% select(-"S.No",-"Department",-"Status")

CIS <- Leave %>% filter(Department =="CIS")
CIS <- CIS %>% select(-"S.No",-"Department",-"Status")

Data_Science <- Leave %>% filter(Department =="Data Science")
Data_Science <- Data_Science %>% select(-"S.No",-"Department",-"Status")

Engineering <- Leave %>% filter(Department =="Engineering")
Engineering <- Engineering %>% select(-"S.No",-"Department",-"Status")

Finance <- Leave %>% filter(Department =="Finance")
Finance <- Finance %>% select(-"S.No",-"Department",-"Status")

Marketing <- Leave %>% filter(Department =="Marketing")
Marketing <- Marketing %>% select(-"S.No",-"Department",-"Status")

Need_Help <- Leave %>% filter(Department =="Need Help")
Need_Help <- Need_Help %>% select(-"S.No",-"Department",-"Status")

Operation <- Leave %>% filter(Department =="Operation")
Operation <- Operation %>% select(-"S.No",-"Department",-"Status")

Product_Management <- Leave %>% filter(Department =="Product Management")
Product_Management <- Product_Management %>% select(-"S.No",-"Department",-"Status")

Referral <- Leave %>% filter(Department =="Referral")
Referral <- Referral %>% select(-"S.No",-"Department",-"Status")

Training_Quality <- Leave %>% filter(Department =="Training & Quality")
Training_Quality <- Training_Quality %>% select(-"S.No",-"Department",-"Status")

Unit_Head <- Leave %>% filter(Department =="Unit Head")
Unit_Head <- Unit_Head %>% select(-"S.No",-"Department",-"Status")

HR <- Leave %>% filter(Department =="HR")
HR <- HR %>% select(-"S.No",-"Department",-"Status")







#Email code-------------------------------------------------------------------

mail_list <- 'E:\\Leave daetails\\MailMasterNew.xlsx'

mail_Others <- read_excel(mail_list, sheet = "Others")



#Mail sent to lender -----------------------------------------------------------
File_list <- list(CIS,HR)

lapply(File_list, function(x){
  file[] <- htmlTable(x,rnames = FALSE)})

mail <- list(mail_Others)

lapply(mail, function(x){
  
  #browser()
  
  x <- data.frame(x)
  lender <- x$File


  lapply(lender, function(y){  
  
  #browser()
    
    row_set <- x[grepl(y, x$File),]
    sender <-  'care@creditmantri.com'
    cc_recipients <-  stringr::str_split(row_set$SendId, ",\\s*")[[1]]
    message  <- str_interp('Weekly Attendance Tracker')
    body1 <- paste0('Dear Team,<br><br>Please see below  the leave details for members on your team from.<br><br>
                        Please check and confirm whether the entries are correct. Wherever a team member is marked as <font color="red">Not Present</font>, please indicate whether this is CL, SL, PL, Travelling or Agency . Please respond by tomorrow eod.<br><br>
                       ',
                    '<style>
                      table {font-family:  Verdana, Geneva, sans-serif; font-size: 11px;
                        width = 100%; border: 1px solid black; border-collapse: collapse;
                        text-align: center; padding: 5px;}
                    th {height = 12px;background-color: #4CAF50;color: white;}
                      td {background-color: #FFF;}
                          </style>',
                    '<table>',File_list,'</table>')
    email_body <- body1
    recipients <- stringr::str_split(row_set$CC, ",\\s*")[[1]]
    mail_function(sender, recipients, cc_recipients, message, email_body)
    
})
  
})


