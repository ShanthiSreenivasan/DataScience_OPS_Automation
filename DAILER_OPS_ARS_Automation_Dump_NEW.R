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
library(bit64)



setwd("C:\\R\\Dialer OPS Automation")

getwd()

thisdate<-format(Sys.Date(),'%Y-%m-%d')

source('.\\Function file.R')

if(!dir.exists(paste("./Output/",thisdate,sep="")))
{
  dir.create(paste("./Output/",thisdate,sep=""))
} 

Arsdump <- fread('./Input/CM_REGISTER_BASE.csv')
Arsdump$lead_id <- as.numeric(Arsdump$lead_id)
Arsdump$total_amount_overdue <- round(Arsdump$total_amount_overdue +1,0)



Output_fields <- read.xlsx(".\\Input\\Automation_ARS_Dialer.xlsx")

Output_fields <- Output_fields %>% filter(files =='1')
  
'%nin%' <- Negate(`%in%`)

for(i in Output_fields$lender){
  
  lender_list <- Output_fields %>% filter(lender %in% c(i)) 

  base_type_1 <- lender_list$base_type
  
  base_1 <- lender_list$base
  
  allocation_vintage_1 <- lapply(strsplit(as.character(lender_list$allocation_vintage),","),as.character)
  
  allocation_vintage_1 <- unlist(allocation_vintage_1)
  
  mo_allocation_vintage_1 <- lender_list$mo_allocation_vintage
  
  product_1 <- lapply(strsplit(as.character(lender_list$product),","),as.character)
  
  product_1 <- unlist(product_1)
  
  is_allocated_1 <- lapply(strsplit(as.character(lender_list$is_allocated),","),as.character)
  
  is_allocated_1 <- unlist(is_allocated_1)
  
  negative_status_accounts_1 <- lapply(strsplit(as.character(lender_list$negative_status_accounts),","),as.character)
  
  negative_status_accounts_1 <- unlist(negative_status_accounts_1)
  
  dialer_base_1 <- lapply(strsplit(lender_list$dialer_base,","),as.character)
  
  dialer_base_1 <- unlist(dialer_base_1)
  
  decile_1 <- lapply(strsplit(as.character(lender_list$decile),","),as.character)
  
  decile_1 <- unlist(decile_1)
  
  m0_attempts_1 <- lapply(strsplit(as.character(lender_list$m0_attempts),","),as.character)
  
  m0_attempts_1 <- unlist(m0_attempts_1)
  
  m0_eff_attempts_1 <- lapply(strsplit(as.character(lender_list$m0_eff_attempts),","),as.character)
  
  m0_eff_attempts_1 <- unlist(m0_eff_attempts_1)
  
  m1_attempts_1 <- lapply(strsplit(as.character(lender_list$m1_attempts),","),as.character)
  
  m1_attempts_1 <- unlist(m1_attempts_1)
  
  m1_eff_attempts_1 <- lapply(strsplit(as.character(lender_list$m1_eff_attempts),","),as.character)
  
  m1_eff_attempts_1 <- unlist(m1_eff_attempts_1)
  
  dispo_1 <- lapply(strsplit(lender_list$DispoNew,","),as.character)
  
  dispo_1 <- unlist(dispo_1)
  
  subscription_vintage1 <- lapply(strsplit(as.character(lender_list$subscription_vintage),","),as.character)
  
  subscription_vintage1 <- unlist(subscription_vintage1)
  
  first_profile_vintage1<- lapply(strsplit(as.character(lender_list$first_profile_vintage),","),as.character)
  
  first_profile_vintage1 <- unlist(first_profile_vintage1)
  
  language_text1<- lapply(strsplit(as.character(lender_list$language_text),","),as.character)
  
  language_text1<- unlist(language_text1)
  
  boa_flag1<- lapply(strsplit(as.character(lender_list$boa_flag),","),as.character)
  
  boa_flag1<- unlist(boa_flag1)
  
  latest_oic1<- lapply(strsplit(as.character(lender_list$latest_oic),","),as.character)
  
  latest_oic1<- unlist(latest_oic1)
  
  product_status1<- lapply(strsplit(as.character(lender_list$product_status),","),as.character)
  
  product_status1<- unlist(product_status1)
  
  ltd_attempts1<- lapply(strsplit(as.character(lender_list$ltd_attempts),","),as.character)
  
  ltd_attempts1<- unlist(ltd_attempts1)
  
  ltd_eff_attempts1<- lapply(strsplit(as.character(lender_list$ltd_eff_attempts),","),as.character)
  
  ltd_eff_attempts1<- unlist(ltd_eff_attempts1)
  
  latest_dispo_ivr1<- lapply(strsplit(as.character(lender_list$latest_dispo_ivr),","),as.character)
  
  latest_dispo_ivr1<- unlist(latest_dispo_ivr1)
  
  latest_oic_ivr1<- lapply(strsplit(as.character(lender_list$latest_oic_ivr),","),as.character)
  
  latest_oic_ivr1<- unlist(latest_oic_ivr1)
  
  ever_ivr_keypress1<- lapply(strsplit(as.character(lender_list$ever_ivr_keypress),","),as.character)
  
  ever_ivr_keypress1<- unlist(ever_ivr_keypress1)
  
  ptp_generated1<- lapply(strsplit(as.character(lender_list$ptp_generated),","),as.character)
  
  ptp_generated1<- unlist(ptp_generated1)
  
  itp_generated1<- lapply(strsplit(as.character(lender_list$itp_generated),","),as.character)
  
  itp_generated1<- unlist(itp_generated1)
  
  m1_ptp_generated1<- lapply(strsplit(as.character(lender_list$m1_ptp_generated),","),as.character)
  
  m1_ptp_generated1<- unlist(m1_ptp_generated1)
  
  m1_itp_generated1<- lapply(strsplit(as.character(lender_list$m1_itp_generated),","),as.character)
  
  m1_itp_generated1<- unlist(m1_itp_generated1)
  
  account_status1<- lapply(strsplit(as.character(lender_list$account_status),","),as.character)
  
  account_status1<- unlist(account_status1)
  
  movement_341_1<- lapply(strsplit(as.character(lender_list$movement_341),","),as.character)
  
  movement_341_1<- unlist(movement_341_1)
  
  movement_349_1<- lapply(strsplit(as.character(lender_list$movement_349),","),as.character)
  
  movement_349_1<- unlist(movement_349_1)
  
  vintage_last_login1<- lapply(strsplit(as.character(lender_list$vintage_last_login),","),as.character)
  
  vintage_last_login1<- unlist(vintage_last_login1)
  
  lender_confirmed_account_status1<- lapply(strsplit(as.character(lender_list$lender_confirmed_account_status),","),as.character)
  
  lender_confirmed_account_status1<- unlist(lender_confirmed_account_status1)
  
  lender_confirmed_account_status1<- lapply(strsplit(as.character(lender_list$lender_confirmed_account_status),","),as.character)
  
  lender_confirmed_account_status1<- unlist(lender_confirmed_account_status1)
  
  bd_decile1<- lapply(strsplit(as.character(lender_list$bd_decile),","),as.character)
  
  bd_decile1<- unlist(bd_decile1)
 
  status_text1<- lapply(strsplit(as.character(lender_list$status_text),","),as.character)
  
  status_text1<- unlist(status_text1)
  
  chatbot_login_count1<- lapply(strsplit(as.character(lender_list$chatbot_login_count),","),as.character)
  
  chatbot_login_count1<- unlist(chatbot_login_count1)
  
  c2c_login_count1<- lapply(strsplit(as.character(lender_list$c2c_login_count),","),as.character)
  
  c2c_login_count1<- unlist(c2c_login_count1)
  
  monthly_income_split1<- lapply(strsplit(as.character(lender_list$monthly_income_split),","),as.character)
  
  monthly_income_split1<- unlist(monthly_income_split1)
  
  team1<- lapply(strsplit(as.character(lender_list$team),","),as.character)
  
  team1<- unlist(team1)
  
  assigned_team1<- lapply(strsplit(as.character(lender_list$assigned_team),","),as.character)
  
  assigned_team1<- unlist(assigned_team1)
  
  no_of_payments1<- lapply(strsplit(as.character(lender_list$no_of_payments),","),as.character)
  
  no_of_payments1<- unlist(no_of_payments1)
  
  tos_flag_1  <- lapply(strsplit(lender_list$tos_flag,","),as.character)
  
  tos_flag_1 <- unlist(tos_flag_1)
  
  ever_ptp_1 <- lender_list$ever_ptp
  
  ever_itp_1 <- lender_list$ever_itp
  
  flows_flag_1 <- lender_list$flows_flag
  
  camp <- lender_list$campaign_name
  
  lender_dump <- Arsdump %>% filter(lender %in% c(i))
  
  lender_dump_1 <- lender_dump %>% filter(base_type %in% base_type_1)
  
  lender_dump_2 <- lender_dump_1 %>% filter(is_allocated %in% is_allocated_1)
  
  lender_dump_3 <- lender_dump_2 %>% filter(`Dispo New` %nin% dispo_1)
  
  lender_dump_4 <- lender_dump_3 %>% filter(tos_flag %nin% tos_flag_1)
    if(!is.na(allocation_vintage_1[1]))
  {
    
  lender_dump_5 <- lender_dump_4 %>% filter(allocation_vintage %in% allocation_vintage_1)
  }else{
  lender_dump_5 <- lender_dump_4
  }
  if(!is.na(mo_allocation_vintage_1[1]))
  {
    
  lender_dump_6 <- lender_dump_5 %>% filter(mo_allocation_vintage %in% mo_allocation_vintage_1)
  }else{
  lender_dump_6 <- lender_dump_5
  }
  if(!is.na(product_1[1]))
  {
    
  lender_dump_7 <- lender_dump_6 %>% filter(product %in% product_1)
  }else{
  lender_dump_7 <- lender_dump_6
  }
  if(!is.na(negative_status_accounts_1[1]))
  {
    
  lender_dump_8 <- lender_dump_7 %>% filter(negative_status_accounts %in% negative_status_accounts_1)
  }else{
  lender_dump_8 <- lender_dump_7
  }
  
    if(!is.na(m0_attempts_1[1]))
  {
    
  lender_dump_9 <- lender_dump_8 %>% filter(m0_attempts %in% m0_attempts_1)
  }else{
  lender_dump_9 <- lender_dump_8
  }
  if(!is.na(m0_eff_attempts_1[1]))
  {
    
  lender_dump_10 <- lender_dump_9 %>% filter(m0_eff_attempts %in% m0_eff_attempts_1)
  }else{
  lender_dump_10 <- lender_dump_9
  }
  if(!is.na(m1_attempts_1[1]))
  {
    
  lender_dump_11 <- lender_dump_10 %>% filter(m1_attempts %in% m1_attempts_1)
  }else{
  lender_dump_11 <- lender_dump_10
  }
  if(!is.na(m1_eff_attempts_1))
  {
    
  lender_dump_12 <- lender_dump_11 %>% filter(m1_eff_attempts %in% m1_eff_attempts_1)
  }else{
  lender_dump_12 <- lender_dump_11
  }
  if(!is.na(ever_ptp_1))
  {
    
  lender_dump_13 <- lender_dump_12 %>% filter(ever_ptp %in% ever_ptp_1)
  }else{
  lender_dump_13 <- lender_dump_12
  }
  if(!is.na(ever_itp_1))
  {
    
  lender_dump_14 <- lender_dump_13 %>% filter(ever_itp %in% ever_itp_1)
  }else{
  lender_dump_14 <- lender_dump_13
  }
  if(!is.na(flows_flag_1))
  {
    
  lender_dump_15 <- lender_dump_14 %>% filter(flows_flag %in% flows_flag_1)
  }else{
  lender_dump_15 <- lender_dump_14
  }
  if(!is.na(dialer_base_1[1]))
  {
    
  lender_dump_16 <- lender_dump_15 %>% filter(dialer_base %in% dialer_base_1)
  }else{
  lender_dump_16 <- lender_dump_15
  }
  if(!is.na(decile_1[1]))
  {
    
  lender_dump_17 <- lender_dump_16 %>% filter(decile %in% decile_1)
  }else{
  lender_dump_17 <- lender_dump_16 
  }
  if(!is.na(subscription_vintage1[1]))
  {
    
    lender_dump_19 <- lender_dump_17 %>% filter(subscription_vintage %in% subscription_vintage1)
  }else{
    lender_dump_19 <- lender_dump_17 
  }
  if(!is.na(first_profile_vintage1[1]))
  {
    
    lender_dump_20 <- lender_dump_19 %>% filter(first_profile_vintage %in% first_profile_vintage1)
  }else{
    lender_dump_20 <- lender_dump_19 
  }
  
  if(!is.na(language_text1[1]))
  {
    
    lender_dump_21 <- lender_dump_20 %>% filter(language_text %in% language_text1)
  }else{
    lender_dump_21 <- lender_dump_20 
  }
  
  if(!is.na(boa_flag1[1]))
  {
    
    lender_dump_22 <- lender_dump_21 %>% filter(boa_flag %in% boa_flag1)
  }else{
    lender_dump_22 <- lender_dump_21 
  }
  
  if(!is.na(no_of_payments1[1]))
  {
    
    lender_dump_23 <- lender_dump_22 %>% filter(no_of_payments %in% no_of_payments1)
  }else{
    lender_dump_23 <- lender_dump_22 
  }
  
  if(!is.na(latest_oic1[1]))
  {
    
    lender_dump_24 <- lender_dump_23 %>% filter(latest_oic %in% latest_oic1)
  }else{
    lender_dump_24 <- lender_dump_23 
  }
  if(!is.na(product_status1[1]))
  {
    
    lender_dump_25 <- lender_dump_24 %>% filter(product_status %in% product_status1)
  }else{
    lender_dump_25 <- lender_dump_24 
  }
  
  if(!is.na(ltd_attempts1[1]))
  {
    
    lender_dump_26 <- lender_dump_25 %>% filter(ltd_attempts %in% ltd_attempts1)
  }else{
    lender_dump_26 <- lender_dump_25 
  }
  if(!is.na(ltd_eff_attempts1[1]))
  {
    
    lender_dump_27 <- lender_dump_26 %>% filter(ltd_eff_attempts %in% ltd_eff_attempts1)
  }else{
    lender_dump_27 <- lender_dump_26 
  }
  
  if(!is.na(latest_dispo_ivr1[1]))
  {
    
    lender_dump_28 <- lender_dump_27 %>% filter(latest_dispo_ivr %in% latest_dispo_ivr1)
  }else{
    lender_dump_28 <- lender_dump_27 
  }
  
  if(!is.na(latest_oic_ivr1[1]))
  {
    
    lender_dump_29 <- lender_dump_28 %>% filter(latest_oic_ivr %in% latest_oic_ivr1)
  }else{
    lender_dump_29 <- lender_dump_28 
  }
  
  if(!is.na(ever_ivr_keypress1[1]))
  {
    
    lender_dump_30 <- lender_dump_29 %>% filter(ever_ivr_keypress %in% ever_ivr_keypress1)
  }else{
    lender_dump_30 <- lender_dump_29 
  }
  
  
  if(!is.na(ptp_generated1[1]))
  {
    
    lender_dump_31 <- lender_dump_30 %>% filter(ptp_generated %in% ptp_generated1)
  }else{
    lender_dump_31 <- lender_dump_30 
  }
  
  if(!is.na(itp_generated1[1]))
  {
    
    lender_dump_32 <- lender_dump_31 %>% filter(itp_generated %in% itp_generated1)
  }else{
    lender_dump_32 <- lender_dump_31 
  }
  
  if(!is.na(m1_ptp_generated1[1]))
  {
    
    lender_dump_33 <- lender_dump_32 %>% filter(m1_ptp_generated %in% m1_ptp_generated1)
  }else{
    lender_dump_33 <- lender_dump_32 
  }
  
  
  if(!is.na(account_status1[1]))
  {
    
    lender_dump_35 <- lender_dump_33 %>% filter(account_status %in% account_status1)
  }else{
    lender_dump_35 <- lender_dump_33 
  }
  
  if(!is.na(movement_341_1[1]))
  {
    
    lender_dump_36 <- lender_dump_35 %>% filter(movement_341 %in% movement_341_1)
  }else{
    lender_dump_36 <- lender_dump_35 
  }
  
  if(!is.na(movement_349_1[1]))
  {
    
    lender_dump_37 <- lender_dump_36 %>% filter(movement_349 %in% movement_349_1)
  }else{
    lender_dump_37 <- lender_dump_36 
  }
  
  if(!is.na(vintage_last_login1[1]))
  {
    
    lender_dump_38 <- lender_dump_37 %>% filter(vintage_last_login %in% vintage_last_login1)
  }else{
    lender_dump_38 <- lender_dump_37 
  }
 
  if(!is.na(lender_confirmed_account_status1[1]))
  {
    
    lender_dump_39 <- lender_dump_38 %>% filter(lender_confirmed_account_status %in% lender_confirmed_account_status1)
  }else{
    lender_dump_39 <- lender_dump_38 
  }
  
  if(!is.na(bd_decile1[1]))
  {
    
    lender_dump_40 <- lender_dump_39 %>% filter(bd_decile %in% bd_decile1)
  }else{
    lender_dump_40 <- lender_dump_39 
  }
  
  if(!is.na(status_text1[1]))
  {
    
    lender_dump_41 <- lender_dump_40 %>% filter(status_text %in% status_text1)
  }else{
    lender_dump_41 <- lender_dump_40 
  }
  
  if(!is.na(chatbot_login_count1[1]))
  {
    
    lender_dump_42 <- lender_dump_41 %>% filter(chatbot_login_count %in% chatbot_login_count1)
  }else{
    lender_dump_42 <- lender_dump_41 
  }
 
  if(!is.na(c2c_login_count1[1]))
  {
    
    lender_dump_43 <- lender_dump_42 %>% filter(c2c_login_count %in% c2c_login_count1)
  }else{
    lender_dump_43 <- lender_dump_42 
  }
  
  if(!is.na(monthly_income_split1[1]))
  {
    
    lender_dump_44 <- lender_dump_43 %>% filter(monthly_income_split %in% monthly_income_split1)
  }else{
    lender_dump_44 <- lender_dump_43 
  }
  
  if(!is.na(team1[1]))
  {
    
    lender_dump_45 <- lender_dump_44 %>% filter(team %in% team1)
  }else{
    lender_dump_45 <- lender_dump_44 
  }
  
  if(!is.na(assigned_team1[1]))
  {
    
    lender_dump_46 <- lender_dump_45 %>% filter(assigned_team %in% assigned_team1)
  }else{
    lender_dump_46 <- lender_dump_45 
  }
  
  
  lender_dump_46$campaign_name <- camp
  
  lender_dump_46$base_type <- base_1
  
  lender_dump_47 <- lender_dump_46 %>% select(lead_id,user_id,customer_name,phone_home,campaign_name,base_type) 
  
  Filename <- paste('.\\Output\\',thisdate,'\\',i,"_",camp,".csv",sep = '')
  FileCreate(lender_dump_47,sheet_name="Sheet1",Filename)
  
}
 
