library(tidyverse)
library(readxl)

#Read in Previous Data
oldvalues = read_csv("P:/jyoung/Accreditation/SPOC_SPAC/COPS Data/PAP Data.csv")


#Place the month names for the previous Quarter in here
current_month <- lubridate::month(as.Date(Sys.Date(), format = "%Y/%m/%d"))
current_year <- lubridate::year(as.Date(Sys.Date(), format = "%Y/%m/%d"))
Q1 <- c("January", "February", "March")
Q2 <- c("April", "May", "June")
Q3 <- c("July", "August", "September")
Q4 <- c("October", "November", "December")


quarter <- switch(current_month, 
                  "1" = Q4, "2" = Q4, "3" = Q4,
                  "4" = Q1, "5" = Q1, "6" = Q1,
                  "7" = Q2, "8" = Q2, "9" = Q2,
                  "10" = Q3, "11" = Q3, "12" = Q3)

#lex_months <- c("January", "February")
#oral_onc_months <- c("July", "August", "September","October", "November", "December")
#all_months <- c("January", "February", "March", "April", "May", "June", "July", "August", "September","October", "November", "December")

year <- ifelse(current_month %in% c(1, 2, 3), current_year - 1, current_year)

names = c("Medication", "Total WAC") # need clarification

#These are the paths for reading in the new data. Each year, the path will need to change
MCPath22 = "P:/Medication Access Center/PAP-(Patient Assistance program)/PAP Values/2022 PAP Values.xlsx"
DHPPath22 = "P:/Medication Access Center/PAP-DHP/DHP-PAP Values/2022 DHP-PAP Values.xlsx"
LexPath22 = "P:/Medication Access Center/Lexington/PAP-Lexington/2022 Lexington PAP Values.xlsx"
DermPath22 = "P:/Medication Access Center/PAP-CCD/PAP CCD Values/2022 CCD PAP Values.xlsx"
OralOncPath22 = "P:/CCC Rx/ONC PAP/ONC Values/2022 ONC PAP Values.xlsx"

#This will read in the data for the months of the last quarter. Each year,
  #the year column in each function will need to be updated
MClist22 = lapply(quarter, function(x){
      dat = read_excel(MCPath22, sheet = x, skip = 1)[c(1,2)] # have to skip all merged lines; looks like 3 for 2023
      names(dat) = names
      dat$Month = x
      dat$Year = year
      dat$Location = "Winston"
      return(dat)
      })


DHPlist22 = lapply(quarter, function(x){
      dat = read_excel(DHPPath22, sheet = x, skip = 1)[c(1,2)]
      names(dat) = names
      dat$Month = x
      dat$Year = year
      dat$Location = "DHP"
      return(dat)
      })
Lexlist22 = lapply(quarter, function(x){
      dat = read_excel(LexPath22, sheet = x, skip = 1)[c(1,2)]
      names(dat) = names
      dat$Month = x
      dat$Year = year
      dat$Location = "Lexington"
      return(dat)
      })

Dermlist22 = lapply(quarter, function(x){
  dat = read_excel(DermPath22, sheet = x, skip = 1)[c(1,2)]
  names(dat) = names
  dat$Month = x
  dat$Year = year
  dat$Location = "Dermatology"
  return(dat)
})

Onclist22 = lapply(quarter, function(x){
  dat = read_excel(OralOncPath22, sheet = x, skip = 1)[c(1,2)]
  names(dat) = names
  dat$Month = x
  dat$Year = year
  dat$Location = "Oral Oncology"
  return(dat)
})

file_names = c("MC_newdata", "DHP_newdata", "Lex_newdata", "Derm_newdata", "Onc_newdata")


MC_newdata = do.call(rbind, MClist22) %>% 
  select( Medication, `Total WAC`, Month, Year, Location) %>%
  mutate(Date.Added = Sys.time())

DHP_newdata = do.call(rbind, DHPlist22)%>% 
  select( Medication, `Total WAC`, Month, Year, Location) %>%
  mutate(Date.Added = Sys.time())

Lex_newdata = do.call(rbind, Lexlist22)%>% 
  select( Medication, `Total WAC`, Month, Year, Location) %>%
  mutate(Date.Added = Sys.time())

Derm_newdata = do.call(rbind, Dermlist22)%>% 
  select( Medication, `Total WAC`, Month, Year, Location) %>%
  mutate(Date.Added = Sys.time())

Onc_newdata = do.call(rbind, Onclist22)%>% 
  select( Medication, `Total WAC`, Month, Year, Location) %>%
  mutate(Date.Added = Sys.time())

data = rbind(MC_newdata, DHP_newdata, Lex_newdata, Derm_newdata, Onc_newdata, oldvalues) %>% 
   filter(!is.na(Medication))

write_csv(data,"H:/R Scripts/Data/PAP Data.csv")
