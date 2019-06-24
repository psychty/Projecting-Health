
# Recreating POPPI

library(easypackages)
# This uses the easypackages package to load several libraries at once. Note: it should only be used when you are confident that all packages are installed as it will be more difficult to spot load errors compared to loading each one individually.
libraries(c("readxl", "readr", "plyr", "dplyr", "ggplot2", "png", "tidyverse", "reshape2", "scales", "viridis", "rgdal", "officer", "flextable", "tmaptools", "lemon", "fingertipsR", "PHEindicatormethods", "xlsx"))

# You must specify a set of chosen areas.
# Areas_to_include <- c("Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "NHS Coastal West Sussex CCG","West Sussex", "NHS Crawley CCG", "NHS Horsham and Mid Sussex CCG")

#Areas_to_include <- c("Basingstoke and Deane","East Hampshire","Eastleigh","Fareham","Gosport","Hart","Havant","New Forest","Rushmoor","Test Valley","Hampshire", "Winchester", "NHS North East Hampshire and Farnham CCG", "NHS North Hampshire CCG", "NHS South Eastern Hampshire CCG", "NHS West Hampshire CCG", "NHS Fareham and Gosport CCG", "Portsmouth")

# If you run the source("~/Projecting-Health/Get data - mye and projections.R") without a variable called Areas_to_include then it will fail. 
# source("~/Projecting-Health/Get data - mye and projections.R")

# I have built in an if statement so that we do not build population files unecessarily (e.g if we already have the chosen data). The programme tries to load the Area_population.df from the root folder. If the file does not exist then the ("~/Projecting-Health/Get data - mye and projections.R")

if(!(exists("Areas_to_include"))){
  print("There are no areas defined. Please create or load 'Areas_to_include' which is a character string of chosen areas.")
}

# Areas_to_include <- c("Basingstoke and Deane","East Hampshire","Eastleigh","Fareham","Gosport","Hart","Havant","New Forest","Rushmoor","Test Valley","Hampshire", "Winchester", "NHS North East Hampshire and Farnham CCG", "NHS North Hampshire CCG", "NHS South Eastern Hampshire CCG", "NHS West Hampshire CCG", "NHS Fareham and Gosport CCG", "Portsmouth")

# Areas_to_include <- c("Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "NHS Coastal West Sussex CCG","West Sussex", "NHS Crawley CCG", "NHS Horsham and Mid Sussex CCG")

Areas_to_include <- c("Eastbourne", "Hastings", "Lewes","Rother", "Wealden","Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "Brighton and Hove", "NHS Brighton and Hove CCG", "NHS Coastal West Sussex CCG", "NHS Crawley CCG","NHS Eastbourne, Hailsham and Seaford CCG", "NHS Hastings and Rother CCG","NHS High Weald Lewes Havens CCG", "NHS Horsham and Mid Sussex CCG", "West Sussex", "East Sussex")

if(!(file.exists("./Projecting-Health/Area_population_df.csv"))){
  print("Area_population_df is not available, it will be built using the 'Areas_to_include' object")
  source("~/Documents/Repositories/Projecting-Health/Get data - mye and projections.R")
  
  if(file.exists("./Projecting-Health/Area_population_df.csv")){
    print("Area_population_df is now available.")
  }
}

if(exists("Areas_to_include") & file.exists("./Projecting-Health/Area_population_df.csv")){
  print("Both objects are available")
  Area_population_df <- read_csv("./Projecting-Health/Area_population_df.csv", col_types = cols(Area_Name = col_character(),Area_Code = col_character(),Area_Type = col_character(),  Sex = col_character(),Age_group = col_character(),Age_band_type = col_character(),Year = col_double(),Population = col_double(),Data_type = col_character()))  
  
  Areas <- read_csv("./Projecting-Health/Area_lookup_table.csv", col_types = cols(LTLA17CD = col_character(),LTLA17NM = col_character(),UTLA17CD = col_character(),UTLA17NM = col_character(),FID = col_double()))
  Lookup <- read_csv("./Projecting-Health/Area_types_table.csv", col_types = cols(Area_Code = col_character(),Area_Name = col_character(),Area_Type = col_character()))
  
  if(length(setdiff(Areas_to_include, Area_population_df$Area_Name))>0){
    print("There are some areas chosen that are not in the Area_population_df. The 'Get data - mye and projections' script will now run and will overwrite the Area_population_df.")
    source("~/Documents/Repositories/Projecting-Health/Get data - mye and projections.R")
  }
  
  
  
  if(length(setdiff(Areas_to_include, Area_population_df$Area_Name))>0){
    print("There are still some areas chosen that are not in the Area_population_df. Check the Areas_to_include object.")
  }
  
  if(length(setdiff(Areas_to_include, Area_population_df$Area_Name))==0){
    print("The Area_population_df matches the Areas_to_include list.")
  }
}

# Assumption changes ####

# We can use this script to also get census tables to use in the poppi calculation (where national assumptions are not used).

# Living alone #

# Figures are taken from the General Household Survey 2007, table 3.4 Percentage of men and women living alone by age, ONS. The General Household Survey is a continuous survey which has been running since 1971, and is based each year on a sample of the general population resident in private households in Great Britain

# paste0("The POPPI assumptions for the proportion of males and females living alone are based on the 2007 General Household Survey by the Office for National Statistiscs. However, there are more recent data available from the 2011 version of this now titled General Lifestyle Survey. The updated data shows an increase in females aged 65-74 and males aged 75 and over living alone.")
# 
# # https://www.ons.gov.uk/peoplepopulationandcommunity/personalandhouseholdfinances/incomeandwealth/compendium/generallifestylesurvey/2013-03-07/chapter3householdsfamilesandpeoplegenerallifestylesurveyoverviewareportonthe2011generallifestylesurvey
# 
# GLS_3.4_2007 <- data.frame(Age_band = c("65-74 years", "75 and over"), Males = c(.20,.34), Females = c(.30,.61))
# GLS_3.4_2011 <- data.frame(Age_band = c("16-24 years", "25-44 years", "45-64 years", "65-74 years", "75 and over"), Males = c(.02,.13,.16,.20,.35), Females = c(.05,.07,.15,.32,.61))
# 
# GLS_3.4_2011 <- GLS_3.4_2011 %>% 
#   gather(key = Sex, value = Percentage, Males:Females)


# Residence type by sex and age ####

# POPPI based estimates of population age 65 and over living in a care home with or without nursing

# This dataset provides 2011 Census estimates that classify usual residents in communal establishments in England and Wales by communal establishment management and type, by sex and by age

# CCGs_residence_type_census <- read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_1086_1.data.csv?geography=TYPE266&c_age=0,17...21&c_residence_type=0...2&c_sex=1,2&measures=20100&select=geography_name,geography_code,c_age_name,c_residence_type_name,c_sex_name,obs_value"), col_types = cols(GEOGRAPHY_NAME = col_character(),GEOGRAPHY_CODE = col_character(),C_AGE_NAME = col_character(),C_RESIDENCE_TYPE_NAME = col_character(),C_SEX_NAME = col_character(),OBS_VALUE = col_integer())) %>% 
  # spread(C_RESIDENCE_TYPE_NAME, OBS_VALUE) %>% 
  # mutate(Perc_living_communal = `Lives in a communal establishment`/`All categories: Residence type`,
  #        GEOGRAPHY_NAME = paste(GEOGRAPHY_NAME, "CCG", sep = " ")) %>% 
  # filter(GEOGRAPHY_NAME %in% Chosen_area_codes$Area_Name)

# District_UA_residence_type_census <- read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_1086_1.data.csv?geography=TYPE464&c_age=0,17...21&c_residence_type=0...2&c_sex=1,2&measures=20100&select=geography_name,geography_code,c_age_name,c_residence_type_name,c_sex_name,obs_value"), col_types = cols(GEOGRAPHY_NAME = col_character(),GEOGRAPHY_CODE = col_character(),C_AGE_NAME = col_character(),C_RESIDENCE_TYPE_NAME = col_character(),C_SEX_NAME = col_character(),OBS_VALUE = col_integer()))%>% 
#   spread(C_RESIDENCE_TYPE_NAME, OBS_VALUE) %>% 
#   mutate(Perc_living_communal = `Lives in a communal establishment`/`All categories: Residence type`) %>% 
#   filter(GEOGRAPHY_CODE %in% Chosen_area_codes$Area_Code)

# County_UA_residence_type_census <- read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_1086_1.data.csv?geography=TYPE463&c_age=0,17...21&c_residence_type=0...2&c_sex=1,2&measures=20100&select=geography_name,geography_code,c_age_name,c_residence_type_name,c_sex_name,obs_value"), col_types = cols(GEOGRAPHY_NAME = col_character(),GEOGRAPHY_CODE = col_character(),C_AGE_NAME = col_character(),C_RESIDENCE_TYPE_NAME = col_character(),C_SEX_NAME = col_character(),OBS_VALUE = col_integer()))%>% 
#   spread(C_RESIDENCE_TYPE_NAME, OBS_VALUE) %>% 
#   mutate(Perc_living_communal = `Lives in a communal establishment`/`All categories: Residence type`) %>% 
#   filter(GEOGRAPHY_CODE %in% Chosen_area_codes$Area_Code)
# 
# Chosen_area_residence_type_census <- County_UA_residence_type_census %>% 
#   bind_rows(District_UA_residence_type_census) %>% 
#   bind_rows(CCGs_residence_type_census) %>% 
#   distinct() %>% 
#   mutate(C_AGE_NAME = gsub("Age ","", C_AGE_NAME)) %>% 
#   mutate(C_AGE_NAME = gsub(" to ", "-", C_AGE_NAME)) %>% 
#   mutate(C_AGE_NAME = gsub(" and over", "+", C_AGE_NAME))
# 
# rm(County_UA_residence_type_census, District_UA_residence_type_census, CCGs_residence_type_census)
# 
# LTLI by general health, age and sex ####

# LTLI_census_health <- read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_829_1.data.csv?date=latest&geography=TYPE266&c_sex=1,2&c_age=0,6...8&c_disability=0...3&c_health=1...3&measures=20100&select=geography_name,geography_code,c_sex_name,c_age_name,c_disability_name,c_health_name,obs_value"), col_types = cols(GEOGRAPHY_NAME = col_character(),GEOGRAPHY_CODE = col_character(),C_SEX_NAME = col_character(),C_AGE_NAME = col_character(),C_DISABILITY_NAME = col_character(),C_HEALTH_NAME = col_character(),OBS_VALUE = col_integer())) %>% 
#   mutate(GEOGRAPHY_NAME = paste(GEOGRAPHY_NAME, "CCG", sep = " ")) %>% 
#   bind_rows(read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_829_1.data.csv?date=latest&geography=TYPE463&c_sex=1,2&c_age=0,6...8&c_disability=0...3&c_health=1...3&measures=20100&select=geography_name,geography_code,c_sex_name,c_age_name,c_disability_name,c_health_name,obs_value"), col_types = cols( GEOGRAPHY_NAME = col_character(),GEOGRAPHY_CODE = col_character(),C_SEX_NAME = col_character(),C_AGE_NAME = col_character(),C_DISABILITY_NAME = col_character(),C_HEALTH_NAME = col_character(),OBS_VALUE = col_integer()))) %>% 
#   bind_rows(read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_829_1.data.csv?date=latest&geography=TYPE464&c_sex=1,2&c_age=0,6...8&c_disability=0...3&c_health=1...3&measures=20100&select=geography_name,geography_code,c_sex_name,c_age_name,c_disability_name,c_health_name,obs_value"), col_types = cols(GEOGRAPHY_NAME = col_character(),GEOGRAPHY_CODE = col_character(),C_SEX_NAME = col_character(),C_AGE_NAME = col_character(),C_DISABILITY_NAME = col_character(),C_HEALTH_NAME = col_character(),OBS_VALUE = col_integer()))) %>% 
#   bind_rows(read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_829_1.data.csv?date=latest&geography=TYPE464&c_sex=1,2&c_age=0,6...8&c_disability=0...3&c_health=1...3&measures=20100&select=geography_name,geography_code,c_sex_name,c_age_name,c_disability_name,c_health_name,obs_value&recordoffset=25000"), col_types = cols(GEOGRAPHY_NAME = col_character(),GEOGRAPHY_CODE = col_character(),C_SEX_NAME = col_character(),C_AGE_NAME = col_character(),C_DISABILITY_NAME = col_character(),C_HEALTH_NAME = col_character(),OBS_VALUE = col_integer()))) %>% 
#   distinct() %>%  
#   spread(C_DISABILITY_NAME, OBS_VALUE) %>% 
#   mutate(Percentage_limited_little = `Day-to-day activities limited a little` / `All categories: Long-term health problem or disability`,
#          Percentage_limited_lot = `Day-to-day activities limited a lot` / `All categories: Long-term health problem or disability`) %>% 
#   mutate(C_AGE_NAME = gsub("Age ","", C_AGE_NAME)) %>% 
#   mutate(C_AGE_NAME = gsub(" to ", "-", C_AGE_NAME)) %>% 
#   mutate(C_AGE_NAME = gsub(" and over", "+", C_AGE_NAME)) %>% 
#   group_by(GEOGRAPHY_NAME, C_SEX_NAME, C_AGE_NAME) %>% 
#   mutate(Percentage_health_status = `All categories: Long-term health problem or disability`/sum(`All categories: Long-term health problem or disability`)) %>% 
#   ungroup()
# 
# LTLI_census_health_2 <- read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_829_1.data.csv?date=latest&geography=TYPE266&c_sex=1,2&c_age=0,6...8&c_disability=0...3&c_health=0&measures=20100&select=geography_name,geography_code,c_sex_name,c_age_name,c_disability_name,c_health_name,obs_value"), col_types = cols(GEOGRAPHY_NAME = col_character(),GEOGRAPHY_CODE = col_character(),C_SEX_NAME = col_character(),C_AGE_NAME = col_character(),C_DISABILITY_NAME = col_character(),C_HEALTH_NAME = col_character(),OBS_VALUE = col_integer())) %>% 
#   mutate(GEOGRAPHY_NAME = paste(GEOGRAPHY_NAME, "CCG", sep = " ")) %>% 
#   bind_rows(read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_829_1.data.csv?date=latest&geography=TYPE463&c_sex=1,2&c_age=0,6...8&c_disability=0...3&c_health=0&measures=20100&select=geography_name,geography_code,c_sex_name,c_age_name,c_disability_name,c_health_name,obs_value"), col_types = cols( GEOGRAPHY_NAME = col_character(),GEOGRAPHY_CODE = col_character(),C_SEX_NAME = col_character(),C_AGE_NAME = col_character(),C_DISABILITY_NAME = col_character(),C_HEALTH_NAME = col_character(),OBS_VALUE = col_integer()))) %>% 
#   bind_rows(read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_829_1.data.csv?date=latest&geography=TYPE464&c_sex=1,2&c_age=0,6...8&c_disability=0...3&c_health=0&measures=20100&select=geography_name,geography_code,c_sex_name,c_age_name,c_disability_name,c_health_name,obs_value"), col_types = cols(GEOGRAPHY_NAME = col_character(),GEOGRAPHY_CODE = col_character(),C_SEX_NAME = col_character(),C_AGE_NAME = col_character(),C_DISABILITY_NAME = col_character(),C_HEALTH_NAME = col_character(),OBS_VALUE = col_integer()))) %>% 
#   bind_rows(read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_829_1.data.csv?date=latest&geography=TYPE464&c_sex=1,2&c_age=0,6...8&c_disability=0...3&c_health=0&measures=20100&select=geography_name,geography_code,c_sex_name,c_age_name,c_disability_name,c_health_name,obs_value&recordoffset=25000"), col_types = cols(GEOGRAPHY_NAME = col_character(),GEOGRAPHY_CODE = col_character(),C_SEX_NAME = col_character(),C_AGE_NAME = col_character(),C_DISABILITY_NAME = col_character(),C_HEALTH_NAME = col_character(),OBS_VALUE = col_integer()))) %>% 
#   distinct() %>%  
#   spread(C_DISABILITY_NAME, OBS_VALUE) %>% 
#   mutate(Percentage_limited_little = `Day-to-day activities limited a little` / `All categories: Long-term health problem or disability`,
#          Percentage_limited_lot = `Day-to-day activities limited a lot` / `All categories: Long-term health problem or disability`) %>% 
#   mutate(C_AGE_NAME = gsub("Age ","", C_AGE_NAME)) %>% 
#   mutate(C_AGE_NAME = gsub(" to ", "-", C_AGE_NAME)) %>% 
#   mutate(C_AGE_NAME = gsub(" and over", "+", C_AGE_NAME)) %>% 
#   group_by(GEOGRAPHY_NAME, C_SEX_NAME, C_AGE_NAME) %>% 
#   mutate(Percentage_health_status = `All categories: Long-term health problem or disability`/sum(`All categories: Long-term health problem or disability`)) %>% 
#   ungroup()
# 
# LTLI_census_health <- LTLI_census_health %>% 
#   bind_rows(LTLI_census_health_2)
# 
# chosen_LTLI_census_health <- LTLI_census_health %>% 
#   filter(GEOGRAPHY_NAME %in% Chosen_area_codes$Area_Name)
# 
# LTLI_age_LA <- read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_1092_1.data.csv?date=latest&geography=TYPE464&c_sex=1,2&c_age=0,14...18&c_disability=1...2&measures=20100&select=geography_name,geography_code,c_sex_name,c_age_name,c_disability_name,obs_value")) %>% 
#   bind_rows(read_csv(url("https://www.nomisweb.co.uk/api/v01/dataset/NM_1092_1.data.csv?date=latest&geography=TYPE463&c_sex=1,2&c_age=0,14...18&c_disability=1...2&measures=20100&select=geography_name,geography_code,c_sex_name,c_age_name,c_disability_name,obs_value"))) %>% 
#   distinct() %>%  
#   spread(C_DISABILITY_NAME, OBS_VALUE) %>% 
#   mutate(`All usual residents` = `Day-to-day activities limited` + `Day-to-day activities not limited`) %>% 
#   mutate(Percentage_limited = `Day-to-day activities limited` / `All usual residents`) %>% 
#   mutate(C_AGE_NAME = gsub("Age ","", C_AGE_NAME)) %>% 
#   mutate(C_AGE_NAME = gsub(" to ", "-", C_AGE_NAME)) %>% 
#   mutate(C_AGE_NAME = gsub(" and over", "+", C_AGE_NAME))
# 
# chosen_LTLI_age_LA <- LTLI_age_LA %>% 
#   filter(GEOGRAPHY_NAME %in% Chosen_area_codes$Area_Name)

# Usual_resident_sex_age_census_WSx <- read_csv(url("http://www.nomisweb.co.uk/api/v01/dataset/NM_1086_1.data.csv?date=&geography=1941962888&c_age=&c_residence_type=&c_sex=&measures=20100&select=c_age_name,c_residence_type_name,c_sex_name,obs_value")) #%>% 
#   filter(C_RESIDENCE_TYPE_NAME == "All categories: Residence type" & C_AGE_NAME %in% c("Age 65 to 69","Age 70 to 74","Age 75 to 79","Age 80 to 84","Age 85 and over","All categories: Age")) %>% 
#   group_by(C_AGE_NAME, C_SEX_NAME) %>% 
#   summarise(Population = sum(OBS_VALUE, na.rm = TRUE)) %>% 
#   rename(Sex = C_SEX_NAME, 
#          Age_band = C_AGE_NAME) %>% 
#   ungroup() %>% 
#   mutate(Age_band = ifelse(Age_band %in% c("Age 65 to 69", "Age 70 to 74"), "65-74 years", ifelse(Age_band %in% c("Age 75 to 79", "Age 80 to 84"), "75-84 years", ifelse(Age_band == "Age 85 and over", "85+ years", "All ages")))) %>% 
#   group_by(Age_band, Sex) %>% 
#   summarise(Population = sum(Population, na.rm = TRUE))
# 
# Care_home_residents_census <- read_csv(url("http://www.nomisweb.co.uk/api/v01/dataset/NM_784_1.data.csv?date=latest&geography=1132462324...1132462330&c_sex=1,2&c_cectmcews11=8,9,15,16&position_age=7...9&measures=20100&select=date_name,geography_name,geography_code,c_sex_name,c_cectmcews11_name,position_age_name,measures_name,obs_value,obs_status_name")) %>% 
#   select(GEOGRAPHY_NAME, C_SEX_NAME, C_CECTMCEWS11_NAME, POSITION_AGE_NAME, OBS_VALUE) %>% 
#   group_by(C_SEX_NAME, C_CECTMCEWS11_NAME, POSITION_AGE_NAME) %>% 
#   summarise(Number = sum(OBS_VALUE, na.rm = TRUE)) %>% 
#   mutate(Establishment_type = ifelse(C_CECTMCEWS11_NAME %in% c("Medical and care establishment: Local Authority: Care home with nursing", "Medical and care establishment: Local Authority: Care home without nursing"), "LA care home with or without nursing", ifelse(C_CECTMCEWS11_NAME %in% c("Medical and care establishment: Other: Care home with nursing", "Medical and care establishment: Other: Care home without nursing"), "Non LA care home with or without nursing", NA))) %>% 
#   group_by(C_SEX_NAME, Establishment_type, POSITION_AGE_NAME) %>% 
#   summarise(Number = sum(Number, na.rm = TRUE)) %>% 
#   rename(Sex = C_SEX_NAME, 
#          Age_band = POSITION_AGE_NAME) %>% 
#   ungroup() %>% 
#   mutate(Age_band = ifelse(Age_band == "Resident: Age 65 to 74", "65-74 years", ifelse(Age_band == "Resident: Age 75 to 84", "75-84 years", ifelse(Age_band == "Resident: Age 85 and over", "85+ years", NA)))) # %>% 
# left_join(Usual_resident_sex_age_census_WSx, by = c("Sex" , "Age_band")) %>% 
#   mutate(Percentage = Number / Population)

# build excel file ####

wb <- loadWorkbook("~/Projecting-Health/POPPI Build.xlsx")

alphabet<- data.frame(num = seq(1,26), let = LETTERS[seq( from = 1, to = 26 )])

# Title and sub title styles
cs_wb_title <- CellStyle(wb) +
  Font(wb,  heightInPoints = 12, isBold = TRUE, underline = 1, name = "Verdana")

# Styles for the data table row/column names
cs_title <- CellStyle(wb) +
  Font(wb, heightInPoints = 12, isBold = TRUE, isItalic = FALSE, name="Verdana") +   Alignment(horizontal = "ALIGN_LEFT", wrapText = FALSE)
cs_left <- CellStyle(wb) +
  Font(wb, heightInPoints = 8, isBold = TRUE, isItalic = FALSE, name="Verdana") +   Alignment(horizontal = "ALIGN_LEFT", vertical = "VERTICAL_TOP", wrapText = TRUE) +
  Border(color = "black", position = c("TOP", "BOTTOM"), pen = c("BORDER_THIN", "BORDER_THIN"))
cs_right <- CellStyle(wb) +
  Font(wb, heightInPoints = 11, isBold = TRUE, isItalic = FALSE, name="Verdana") +   Alignment(horizontal = "ALIGN_RIGHT", vertical = "VERTICAL_TOP", wrapText = TRUE) +
  Border(color = "black", position = c("TOP", "BOTTOM"), pen = c("BORDER_THIN", "BORDER_THIN"))
cs_thousand_sep <- CellStyle(wb) +
  DataFormat("#,##0")
cs_monies <- CellStyle(wb) +
  DataFormat("?#,##0.00")
cs_top_border <- CellStyle(wb) +
  Border(color = "black", position = "TOP", pen = "BORDER_THIN")
cs_bottom_border <- CellStyle(wb) +
  Border(color = "black", position = "BOTTOM", pen = "BORDER_THIN")

cs_top_bottom_border <- CellStyle(wb) +
  Border(color = "black", position = c("TOP","BOTTOM"), pen = "BORDER_THIN")

cs_monies_coloured <- CellStyle(wb) +
  DataFormat("?#,##0.00") + 
  Fill(foregroundColor = "#f1eae1")

cs_thousand_sep_coloured <- CellStyle(wb) +
  DataFormat("#,##0")+ 
  Fill(foregroundColor = "#f1eae1")

cs_right_coloured <- CellStyle(wb) +
  Font(wb, heightInPoints = 11, isBold = TRUE, isItalic = FALSE, name="Verdana") +  
  Alignment(horizontal = "ALIGN_RIGHT", vertical = "VERTICAL_TOP", wrapText = TRUE) +
  Border(color = "black", position = c("TOP", "BOTTOM"), pen = c("BORDER_THIN", "BORDER_THIN"))+ 
  Fill(foregroundColor = "#f1eae1")

removeSheet(wb, sheetName = "List")
sheet <- createSheet(wb, "List")
rows <- createRow(sheet, rowIndex = 1:max(length(unique(Area_population_df$Area_Name)), length(unique(Area_population_df$Data_type)), length(unique(Area_population_df$Year)), na.rm = TRUE))
cells <- createCell(rows, colIndex = 1:3)

setCellValue(cells[[1,1]], "Area")
setCellStyle(cells[[1,1]], cs_left)
addDataFrame(unique(Area_population_df$Area_Name), sheet, 
             startRow = 2, 
             startColumn = 1, 
             col.names = FALSE,
             row.names = FALSE)
setColumnWidth(sheet, colIndex = 1, colWidth = max(nchar(Area_population_df$Area_Name)))

setCellValue(cells[[1,2]], "Type")
setCellStyle(cells[[1,2]], cs_left)
addDataFrame(unique(Area_population_df$Data_type), sheet, 
             startRow = 2, 
             startColumn = 2, 
             col.names = FALSE,
             row.names = FALSE)
setColumnWidth(sheet, colIndex = 2, colWidth = max(nchar(Area_population_df$Data_type)))

setCellValue(cells[[1,3]], "Years")
setCellStyle(cells[[1,3]], cs_left)
addDataFrame(unique(Area_population_df$Year), sheet, 
             startRow = 2, 
             startColumn = 3, 
             col.names = FALSE,
             row.names = FALSE)
setColumnWidth(sheet, colIndex = 3, colWidth = 6)

createRange("Area_list", cells[[2,1]], cells[[length(unique(Area_population_df$Area_Name)),1]])

# For some reason saveWorkbook does not like the tilde (~) so we have to set the wd 
setwd("~/Projecting-Health")

wb$setActiveSheet(0L)
wb$setSheetHidden(10L, 1L) # This assumes List is sheet number 10.

removeSheet(wb, sheetName = "Raw Data")
sheet <- createSheet(wb, "Raw Data")
rows <- createRow(sheet, rowIndex = 1:nrow(Area_population_df))
cells <- createCell(rows, colIndex = 1:9)

setCellValue(cells[[1,1]], "Area_name")
setCellStyle(cells[[1,1]], cs_left)
setCellValue(cells[[1,2]], "Area_code")
setCellStyle(cells[[1,2]], cs_left)
setCellValue(cells[[1,3]], "Area_type")
setCellStyle(cells[[1,3]], cs_left)
setCellValue(cells[[1,4]], "Sex")
setCellStyle(cells[[1,4]], cs_left)
setCellValue(cells[[1,5]], "Age_group")
setCellStyle(cells[[1,5]], cs_left)
setCellValue(cells[[1,6]], "Age_band_type")
setCellStyle(cells[[1,6]], cs_left)
setCellValue(cells[[1,7]], "Year")
setCellStyle(cells[[1,7]], cs_left)
setCellValue(cells[[1,8]], "Population")
setCellStyle(cells[[1,8]], cs_left)
setCellValue(cells[[1,9]], "Data_type")
setCellStyle(cells[[1,9]], cs_left)

addDataFrame(as.data.frame(Area_population_df), sheet, 
             startRow = 2, 
             startColumn = 1, 
             col.names = FALSE,
             row.names = FALSE)

wb$setSheetHidden(10L, 1L) # This assumes Raw Data is sheet number 10 (which it should be now, as index starts at 0)

autoSizeColumn(sheet, colIndex = 1:9)

# to do - Load parameters sheet (or rebuild) and specify data validation range ####

saveWorkbook(wb, file = "./POPPI 2019 Update v1.xlsx")
