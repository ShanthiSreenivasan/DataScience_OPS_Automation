#Temp1 Format
Temp <- function(NewLeads){
  
  col <- c('Application_Number','Customer_Name','Email','Phone_No','Current_Residence_City',
           'Loan_Amount_Required','What_date_do_you_want_to_give_an_appointment_to_the_lender',
           'Net_Take_Home_Per_Month','Profit_After_Tax')
  
  NewLeads <- NewLeads %>% select(col)
  NewLeads$Date <- Sys.Date()
  NewLeads <- NewLeads %>% select(Date, 1:9)
  
  names <- c('Date','ApplicationNumber','CustomerName','Email','Phone','Location','LoanAmountRequired',
             'AppointmentDate','NetSalary','ProfitAfterTax')
  names(NewLeads) <- names
  
  return(NewLeads)
  
}


#Create File
FileCreation <- function(NewLeads, FilePath){
  
  wb <- createWorkbook()
  addWorksheet(wb,'Sheet1')
  hs1 <- createStyle(fgFill = "#4F81BD", 
                     halign = "CENTER", 
                     textDecoration = "Bold",
                     border = "Bottom", 
                     fontColour = "white")
  setColWidths(wb,'Sheet1',cols = 1:ncol(NewLeads), widths = 'auto')
  setColWidths(wb,'Sheet1',cols = 1, widths = 10)
  writeData(wb, "Sheet1", x = NewLeads ,headerStyle = hs1, borders = "all")
  saveWorkbook(wb,FilePath,overwrite = T)

}


#Zip File ----------------------------------------------------------------------
Sys.setenv('R_ZIPCMD' = 'C:/Rtools/bin/zip.exe')
zip.file <- function(source_path){

dir_list <- list.dirs(source_path)
dir_list <- dir_list[ dir_list != source_path]

get_password <- function(lender) {
  case_when(
    grepl('HDB', lender, ignore.case = T) ~ 'Cmhdb!123',
    grepl('HDFC', lender, ignore.case = T) ~ 'Cmhdfc!123',
    grepl('Indusind|IndusInd', lender, ignore.case = T) ~ 'Cmindusind!123',
    grepl('DCB', lender, ignore.case = T) ~ 'Cmdcb!123',
    grepl('SHUBHAM', lender, ignore.case = T) ~ 'Cmshubham!123',
    grepl('Sundaram', lender, ignore.case = T) ~ 'Cmsundaram!123',
    grepl('Yes', lender, ignore.case = T) ~ 'Cmyes!123',
    grepl('BOB', lender, ignore.case = T) ~ 'Cmbob!123',
    grepl('Ujjivan',lender, ignore.case = T) ~ 'Cmujjivan!123',
    grepl('Kredit',lender, ignore.case = T) ~ 'Cmkredit!123',
    T ~ 'cm@1234'
  )
}

lapply(dir_list, function(x){
  
  # browser()
  zip_path <- stringr::str_replace(x, 'Excel_File','Zip')
  if(!dir.exists(zip_path)) dir.create(zip_path)
  
  files <- list.files(x)
  
  lapply(files, function(y){
    
   # browser()
    source_file <- paste(x,'\\',y,sep='')
    zip_path <- paste(zip_path,'\\',y,sep='')
    zip_file <- stringr::str_replace(zip_path,'.xlsx','.zip')
  
    zip(zip_file,source_file, flags = paste("-j -r9Xj -P", get_password(y)))
    
    #zip(zip_file,source_file, flags = paste("-j -r9Xj -P", 'Cm@1234'))

    
  })
  
  
})

}
#Sudharsan Mail ----------------------------------------------------------------
SendMail=function(df1,subj,recepients,cclist){
  
  myMessage = paste0(subj,"-", Sys.Date())
  
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













