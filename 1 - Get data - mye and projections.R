
# Use this r script to download local authority and ccg level population estimates and projections from ONS for chosen areas.

# We need to create a script that outputs the spreadsheet files (csvs)

# Notes
# 1.  Long-term subnational population projections are an indication of the future trends in population by age and sex over the next 25 years. They are trend-based projections, which means assumptions for future levels of births, deaths and migration are based on observed levels mainly over the previous 5 years. They show what the population will be if recent trends continue.

# 2.  The projected resident population of an area includes all people who usually live there, whatever their nationality.  People moving into or out of the country are only included in the resident population if their total stay in that area is for 12 months or more, thus visitors and short-term migrants are not included. Armed forces stationed abroad are not included, but armed forces stationed within an area are included. Students are taken to be resident at their term-time address.

# 3.  The projections generally do not take into account any policy changes that have not yet occurred, nor those that have not yet had an impact on observed trends. They are constrained at the national level to the 2014-based national projections published on 29 October 2015.

# 4.  These projections published on 25 May 2016 are based on the 2014 mid-year population estimates published on 25 June 2015.

# 5.  These data are based on administrative geographic boundaries in place on the reference date of 30 June 2014.
# 6.  Further advice on the appropriate use of these data can be obtained by emailing projections@ons.gsi.gov.uk or phoning +44 (0) 1329 44 4652.
# 
# Figures may not add exactly due to rounding.

# Create a project folder if it does not exist in your working directory
if(file.exists("./Projecting-Health") ==  FALSE){
  dir.create("./Projecting-Health")}

github_repo_dir <- "~/Documents/Repositories/Projecting-Health"

# Choose areas ####

# Areas_to_include <- c("Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "NHS Coastal West Sussex CCG","West Sussex", "NHS Crawley CCG", "NHS Horsham and Mid Sussex CCG")

#Areas_to_include <- c("Basingstoke and Deane","East Hampshire","Eastleigh","Fareham","Gosport","Hart","Havant","New Forest","Rushmoor","Test Valley","Hampshire", "Winchester", "NHS North East Hampshire and Farnham CCG", "NHS North Hampshire CCG", "NHS South Eastern Hampshire CCG", "NHS West Hampshire CCG", "NHS Fareham and Gosport CCG", "Portsmouth")

Areas_to_include <- c("Eastbourne", "Hastings", "Lewes","Rother", "Wealden","Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "Brighton and Hove", "NHS Brighton and Hove CCG", "NHS Coastal West Sussex CCG", "NHS Crawley CCG","NHS Eastbourne, Hailsham and Seaford CCG", "NHS Hastings and Rother CCG","NHS High Weald Lewes Havens CCG", "NHS Horsham and Mid Sussex CCG", "West Sussex", "East Sussex", "England")

# This uses the easypackages package to load several libraries at once. Note: it should only be used when you are confident that all packages are installed as it will be more difficult to spot load errors compared to loading each one individually.
library(easypackages)

libraries(c("readxl", "readr", "plyr", "dplyr", "ggplot2", "png", "tidyverse", "reshape2", "scales", "viridis", "rgdal", "officer", "flextable", "tmaptools", "lemon", "fingertipsR", "PHEindicatormethods"))

if(!(exists("Areas_to_include"))){
 print("There are no areas defined. Please create or load 'Areas_to_include' which is a character string of chosen areas.")
}

if(exists("Areas_to_include")){

capwords = function(s, strict = FALSE) {
  cap = function(s) paste(toupper(substring(s, 1, 1)),
                          {s = substring(s, 2); if(strict) tolower(s) else s},sep = "", collapse = " " )
  sapply(strsplit(s, split = " "), cap, USE.NAMES = !is.null(names(s)))}

# Create a lookup for area codes ####

if(!(file.exists("./Projecting-Health/Area_lookup_table.csv") & file.exists("~/Projecting-Health/Area_types_table.csv"))){
LAD <- read_csv(url("https://opendata.arcgis.com/datasets/a267b55f601a4319a9955b0197e3cb81_0.csv"), col_types = cols(LAD17CD = col_character(),LAD17NM = col_character(),  LAD17NMW = col_character(),  FID = col_integer()))

Counties <- read_csv(url("https://opendata.arcgis.com/datasets/7e6bfb3858454ba79f5ab3c7b9162ee7_0.csv"), col_types = cols(CTY17CD = col_character(),  CTY17NM = col_character(),  Column2 = col_character(),  Column3 = col_character(),  FID = col_integer()))

lookup <- read_csv(url("https://opendata.arcgis.com/datasets/41828627a5ae4f65961b0e741258d210_0.csv"), col_types = cols(LTLA17CD = col_character(),  LTLA17NM = col_character(),  UTLA17CD = col_character(),  UTLA17NM = col_character(),  FID = col_integer()))
# This is a lower tier LA to upper tier LA lookup
UA <- subset(lookup, LTLA17NM == UTLA17NM)

CCG <- read_csv(url("https://opendata.arcgis.com/datasets/4010cd6fc6ce42c29581c4654618e294_0.csv"), col_types = cols(CCG18CD = col_character(),CCG18CDH = col_skip(),CCG18NM = col_character(), FID = col_skip())) %>% 
  rename(Area_Name = CCG18NM,
         Area_Code = CCG18CD) %>% 
  mutate(Area_Type = "Clinical Commissioning Group (2018)")

Region <- read_csv(url("https://opendata.arcgis.com/datasets/cec20f3a9a644a0fb40fbf0c70c3be5c_0.csv"), col_types = cols(RGN17CD = col_character(),  RGN17NM = col_character(),  RGN17NMW = col_character(),  FID = col_integer()))
colnames(Region) <- c("Area_Code", "Area_Name", "Area_Name_Welsh", "FID")

Region$Area_Type <- "Region"
Region <- Region[c("Area_Code", "Area_Name", "Area_Type")]

LAD <- subset(LAD, substr(LAD$LAD17CD, 1, 1) == "E")
LAD$Area_Type <- ifelse(LAD$LAD17NM %in% UA$LTLA17NM, "Unitary Authority", "District")
colnames(LAD) <- c("Area_Code", "Area_Name", "Area_Name_Welsh", "FID", "Area_Type")
LAD <- LAD[c("Area_Code", "Area_Name", "Area_Type")]

Counties$Area_type <- "County"
colnames(Counties) <- c("Area_Code", "Area_Name", "Col2", "Col3", "FID", "Area_Type")
Counties <- Counties[c("Area_Code", "Area_Name", "Area_Type")]

England <- data.frame(Area_Code = "E92000001", Area_Name = "England", Area_Type = "Country")

Areas <- rbind(LAD, CCG, Counties, Region, England)
rm(LAD, CCG, Counties, Region, England, UA)

write.csv(lookup, "./Projecting-Health/Area_lookup_table.csv", row.names = FALSE)
write.csv(Areas, "./Projecting-Health/Area_types_table.csv", row.names = FALSE)
}

Lookup <- read_csv("./Projecting-Health/Area_lookup_table.csv", col_types = cols(LTLA17CD = col_character(),LTLA17NM = col_character(), UTLA17CD = col_character(),  UTLA17NM = col_character(), FID = col_character()))
Areas <- read_csv("./Projecting-Health/Area_types_table.csv", col_types = cols(Area_Code = col_character(), Area_Name = col_character(), Area_Type = col_character()))
}

if(length(setdiff(Areas_to_include, Areas$Area_Name))>0){
  print(paste0("Some areas in the Areas_to_include list are not likely to be available. Check ", setdiff(Areas_to_include, Areas$Area_Name)))
}

if(length(setdiff(Areas_to_include, Areas$Area_Name)) == 0){
# NOMIS geographies ####

NOMIS_codes <- read_csv(paste0(github_repo_dir,"/NOMIS_area_codes.csv"), col_types = cols(GEOGRAPHY_CODE = col_character(),GEOGRAPHY_NAME = col_character(),GEOGRAPHY = col_double(),  Area_Type = col_character()))

Chosen_area_codes <- subset(Areas, Area_Name %in% Areas_to_include) %>% 
  left_join(NOMIS_codes[c("GEOGRAPHY", "GEOGRAPHY_CODE")], by = c("Area_Code" = "GEOGRAPHY_CODE"))

# Projections ####

if(!(file.exists("./Projecting-Health/2016 SNPP CCG pop females.csv"))){
download.file("https://www.ons.gov.uk/file?uri=/peoplepopulationandcommunity/populationandmigration/populationprojections/datasets/clinicalcommissioninggroupsinenglandz2/2016based/snppz2ccgpop.zip", "./Projecting-Health/CCG_projections_2016_based.zip", mode = "wb")
unzip("./Projecting-Health/CCG_projections_2016_based.zip", exdir = "./Projecting-Health")
file.remove("./Projecting-Health/CCG_projections_2016_based.zip")
}

ONS_projections_SYOA <- data.frame(GEOGRAPHY = double(),GEOGRAPHY_NAME = character(), GEOGRAPHY_CODE = character(),PROJECTED_YEAR_NAME = double(),GENDER_NAME = character(),C_AGE_NAME = character(),  MEASURES_NAME = character(), OBS_VALUE = double(), OBS_STATUS_NAME = character(), RECORD_COUNT = double())

for(i in 0:floor(as.numeric(read_csv(paste0("http://www.nomisweb.co.uk/api/v01/dataset/NM_2006_1.data.csv?geography=",paste(as.numeric(Chosen_area_codes$GEOGRAPHY), collapse = ","),"&projected_year=2016...2041&gender=1,2&c_age=101...191&measures=20100&select=record_count&recordlimit=1"), col_types = cols(RECORD_COUNT = col_double())))/25000)){
  df <- read_csv(url(paste0("http://www.nomisweb.co.uk/api/v01/dataset/NM_2006_1.data.csv?geography=",paste(as.numeric(Chosen_area_codes$GEOGRAPHY), collapse = ","),"&projected_year=2016...2041&gender=1,2&c_age=101...191&measures=20100&select=geography,geography_name,geography_code,projected_year_name,gender_name,c_age_name,measures_name,obs_value,obs_status_name,record_count&recordoffset=", 25000 * i)), col_types = cols(GEOGRAPHY = col_double(),GEOGRAPHY_NAME = col_character(),  GEOGRAPHY_CODE = col_character(),PROJECTED_YEAR_NAME = col_double(), GENDER_NAME = col_character(),C_AGE_NAME = col_character(), MEASURES_NAME = col_character(),OBS_VALUE = col_double(), OBS_STATUS_NAME = col_character(),RECORD_COUNT = col_double()))
  
  ONS_projections_SYOA <- ONS_projections_SYOA %>% 
    bind_rows(df)
}

ONS_projections_SYOA <- ONS_projections_SYOA %>% 
  rename(AREA_CODE = GEOGRAPHY_CODE,
         AREA_NAME = GEOGRAPHY_NAME,
         SEX = GENDER_NAME) %>% 
  mutate(SEX = ifelse(SEX == "Male", "males", ifelse(SEX == "Female", "females", NA))) %>% 
  mutate(AGE_GROUP = gsub("Age ", "", C_AGE_NAME)) %>% 
  select(AREA_CODE, AREA_NAME, SEX, AGE_GROUP, PROJECTED_YEAR_NAME, OBS_VALUE) %>% 
  spread(PROJECTED_YEAR_NAME, OBS_VALUE) %>% 
  mutate(AGE_GROUP  = ifelse(AGE_GROUP == "Aged 90+", "90 and over", AGE_GROUP))

# We will need to add any CCG projections to this.
setdiff(Chosen_area_codes$Area_Name, ONS_projections_SYOA$AREA_NAME) # This shows that they were not included

ONS_ccg_projections_df <- read_csv("./Projecting-Health/2016 SNPP CCG pop females.csv", col_types = cols(.default = col_double(), AREA_CODE = col_character(),  AREA_NAME = col_character(),  COMPONENT = col_character(),  SEX = col_character(),  AGE_GROUP = col_character())) %>% 
  bind_rows(read_csv("./Projecting-Health/2016 SNPP CCG pop males.csv",col_types = cols(.default = col_double(), AREA_CODE = col_character(),  AREA_NAME = col_character(),  COMPONENT = col_character(),  SEX = col_character(),  AGE_GROUP = col_character()))) %>% 
  filter(AREA_CODE %in% c(Chosen_area_codes$Area_Code)) %>% 
  select(-c(COMPONENT)) %>% 
  mutate_if(is.numeric, round,0)

# We have rounded the numbers because NOMIS rounds the projected values for LAs

# Now we have a file containing our areas
ONS_projections_SYOA <- ONS_projections_SYOA %>% 
  bind_rows(ONS_ccg_projections_df)

# Areas all included
setdiff(Chosen_area_codes$Area_Name, ONS_projections_SYOA$AREA_NAME) 

ONS_projection_1941_quinary <- ONS_projections_SYOA %>% 
  filter(AGE_GROUP != "All ages") %>% 
  mutate(Age = as.numeric(gsub(" and over", "", AGE_GROUP))) %>% 
  mutate(`Age group` = ifelse(Age <= 4, "0-4 years", ifelse(Age <= 9, "5-9 years", ifelse(Age <= 14, "10-14 years", ifelse(Age <= 19, "15-19 years", ifelse(Age <= 24, "20-24 years", ifelse(Age <= 29, "25-29 years",ifelse(Age <= 34, "30-34 years", ifelse(Age <= 39, "35-39 years",ifelse(Age <= 44, "40-44 years", ifelse(Age <= 49, "45-49 years",ifelse(Age <= 54, "50-54 years", ifelse(Age <= 59, "55-59 years",ifelse(Age <= 64, "60-64 years", ifelse(Age <= 69, "65-69 years",ifelse(Age <= 74, "70-74 years", ifelse(Age <= 79, "75-79 years",ifelse(Age <= 84, "80-84 years", ifelse(Age <= 89, "85-89 years", "90+ years"))))))))))))))))))) %>% 
  group_by(AREA_NAME, AREA_CODE, SEX, `Age group`) %>% 
  summarise(`2018` = sum(`2018`, na.rm = TRUE),
            `2019` = sum(`2019`, na.rm = TRUE),
            `2020` = sum(`2020`, na.rm = TRUE),
            `2021` = sum(`2021`, na.rm = TRUE),
            `2022` = sum(`2022`, na.rm = TRUE),
            `2023` = sum(`2023`, na.rm = TRUE),
            `2024` = sum(`2024`, na.rm = TRUE),
            `2025` = sum(`2025`, na.rm = TRUE),
            `2026` = sum(`2026`, na.rm = TRUE),
            `2027` = sum(`2027`, na.rm = TRUE),
            `2028` = sum(`2028`, na.rm = TRUE),
            `2029` = sum(`2029`, na.rm = TRUE),
            `2030` = sum(`2030`, na.rm = TRUE),
            `2031` = sum(`2031`, na.rm = TRUE),
            `2032` = sum(`2032`, na.rm = TRUE),
            `2033` = sum(`2033`, na.rm = TRUE),
            `2034` = sum(`2034`, na.rm = TRUE),
            `2035` = sum(`2035`, na.rm = TRUE),
            `2036` = sum(`2036`, na.rm = TRUE),
            `2037` = sum(`2037`, na.rm = TRUE),
            `2038` = sum(`2038`, na.rm = TRUE),
            `2039` = sum(`2039`, na.rm = TRUE),
            `2040` = sum(`2040`, na.rm = TRUE),
            `2041` = sum(`2041`, na.rm = TRUE)) %>% 
  ungroup() %>% 
  mutate(SEX = capwords(SEX)) %>% 
  rename(Area_name = AREA_NAME,
         Area_code = AREA_CODE,
         Sex = SEX) %>% 
  mutate(Sex = ifelse(Sex == "Females", "Female", ifelse(Sex == "Males", "Male", Sex))) %>% 
  gather(Year, Population, `2018`:`2041`, factor_key = TRUE) %>% 
  mutate(Data_type = "Projected - ONS",
         Age_band_type = "5 years")

ONS_projection_1941_10_year <- ONS_projections_SYOA %>% 
  filter(AGE_GROUP != "All ages") %>% 
  mutate(Age = as.numeric(gsub(" and over", "", AGE_GROUP))) %>% 
  mutate(`Age group` = ifelse(Age <= 9, "0-9 years", ifelse(Age <= 19, "10-19 years", ifelse(Age <= 29, "20-29 years", ifelse(Age <= 39, "30-39 years", ifelse(Age <= 49, "40-49 years", ifelse(Age <= 59, "50-59 years",ifelse(Age <= 69, "60-69 years", ifelse(Age <= 79, "70-79 years",ifelse(Age <= 89, "80-89 years", "90+ years")))))))))) %>% 
  group_by(AREA_NAME, AREA_CODE, SEX, `Age group`) %>% 
  summarise(`2018` = sum(`2018`, na.rm = TRUE),
            `2019` = sum(`2019`, na.rm = TRUE),
            `2020` = sum(`2020`, na.rm = TRUE),
            `2021` = sum(`2021`, na.rm = TRUE),
            `2022` = sum(`2022`, na.rm = TRUE),
            `2023` = sum(`2023`, na.rm = TRUE),
            `2024` = sum(`2024`, na.rm = TRUE),
            `2025` = sum(`2025`, na.rm = TRUE),
            `2026` = sum(`2026`, na.rm = TRUE),
            `2027` = sum(`2027`, na.rm = TRUE),
            `2028` = sum(`2028`, na.rm = TRUE),
            `2029` = sum(`2029`, na.rm = TRUE),
            `2030` = sum(`2030`, na.rm = TRUE),
            `2031` = sum(`2031`, na.rm = TRUE),
            `2032` = sum(`2032`, na.rm = TRUE),
            `2033` = sum(`2033`, na.rm = TRUE),
            `2034` = sum(`2034`, na.rm = TRUE),
            `2035` = sum(`2035`, na.rm = TRUE),
            `2036` = sum(`2036`, na.rm = TRUE),
            `2037` = sum(`2037`, na.rm = TRUE),
            `2038` = sum(`2038`, na.rm = TRUE),
            `2039` = sum(`2039`, na.rm = TRUE),
            `2040` = sum(`2040`, na.rm = TRUE),
            `2041` = sum(`2041`, na.rm = TRUE)) %>% 
  ungroup() %>% 
  mutate(SEX = capwords(SEX)) %>% 
  rename(Area_name = AREA_NAME,
         Area_code = AREA_CODE,
         Sex = SEX) %>% 
  mutate(Sex = ifelse(Sex == "Females", "Female", ifelse(Sex == "Males", "Male", Sex))) %>% 
  gather(Year, Population, `2018`:`2041`, factor_key = TRUE) %>% 
  mutate(Data_type = "Projected - ONS",
         Age_band_type = "10 years")

ONS_projection_1941_broad <- ONS_projections_SYOA %>% 
  filter(AGE_GROUP != "All ages") %>% 
  mutate(Age = as.numeric(gsub(" and over", "", AGE_GROUP))) %>% 
  mutate(`Age group` = ifelse(Age <= 17, "0-17 years", ifelse(Age <= 44, "18-44 years", ifelse(Age <= 64, "45-64 years", "65+ years")))) %>% 
  group_by(AREA_NAME, AREA_CODE, SEX, `Age group`) %>% 
  summarise(`2018` = sum(`2018`, na.rm = TRUE),
            `2019` = sum(`2019`, na.rm = TRUE),
            `2020` = sum(`2020`, na.rm = TRUE),
            `2021` = sum(`2021`, na.rm = TRUE),
            `2022` = sum(`2022`, na.rm = TRUE),
            `2023` = sum(`2023`, na.rm = TRUE),
            `2024` = sum(`2024`, na.rm = TRUE),
            `2025` = sum(`2025`, na.rm = TRUE),
            `2026` = sum(`2026`, na.rm = TRUE),
            `2027` = sum(`2027`, na.rm = TRUE),
            `2028` = sum(`2028`, na.rm = TRUE),
            `2029` = sum(`2029`, na.rm = TRUE),
            `2030` = sum(`2030`, na.rm = TRUE),
            `2031` = sum(`2031`, na.rm = TRUE),
            `2032` = sum(`2032`, na.rm = TRUE),
            `2033` = sum(`2033`, na.rm = TRUE),
            `2034` = sum(`2034`, na.rm = TRUE),
            `2035` = sum(`2035`, na.rm = TRUE),
            `2036` = sum(`2036`, na.rm = TRUE),
            `2037` = sum(`2037`, na.rm = TRUE),
            `2038` = sum(`2038`, na.rm = TRUE),
            `2039` = sum(`2039`, na.rm = TRUE),
            `2040` = sum(`2040`, na.rm = TRUE),
            `2041` = sum(`2041`, na.rm = TRUE)) %>% 
  ungroup() %>% 
  mutate(SEX = capwords(SEX)) %>% 
  rename(Area_name = AREA_NAME,
         Area_code = AREA_CODE,
         Sex = SEX) %>% 
  mutate(Sex = ifelse(Sex == "Females", "Female", ifelse(Sex == "Males", "Male", Sex))) %>% 
  gather(Year, Population, `2018`:`2041`, factor_key = TRUE) %>% 
  mutate(Data_type = "Projected - ONS",
         Age_band_type = "broad years")

Areas_projections_file <- ONS_projection_1941_quinary %>% 
  bind_rows(ONS_projection_1941_10_year) %>% 
  bind_rows(ONS_projection_1941_broad) %>% 
  left_join(Areas, by = c("Area_name" = "Area_Name")) %>% 
  select(Area_name, Area_Code, Area_Type, Sex, `Age group`,Age_band_type, Year, Population, Data_type) %>% 
  rename(Age_group = `Age group`,
         Area_Name = Area_name)

# This is the WSCC housing stock weighted projections. This factors in projected housing growth as a constraint. 

# WSCC_HH_projection_1832_quinary <- data.frame(Area_name = "Adur", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Adur", skip = 4, n_max = 186), check.names = FALSE) %>%   bind_rows(data.frame(Area_name = "Arun", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Arun", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Chichester", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Chichester", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Crawley", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Crawley", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Horsham", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Horsham", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Mid Sussex", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Mid_Sussex", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Worthing", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Worthing", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "West Sussex", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "WSx", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   filter(Sex %in% c("male", "female")) %>% 
#   mutate(Age = as.numeric(ifelse(Age == "90+", "90", Age)),
#          Sex = capwords(Sex, strict = TRUE)) %>% 
#   mutate(`Age group` = factor(ifelse(Age <= 4, "0-4 years", ifelse(Age <= 9, "5-9 years", ifelse(Age <= 14, "10-14 years", ifelse(Age <= 19, "15-19 years", ifelse(Age <= 24, "20-24 years", ifelse(Age <= 29, "25-29 years",ifelse(Age <= 34, "30-34 years", ifelse(Age <= 39, "35-39 years",ifelse(Age <= 44, "40-44 years", ifelse(Age <= 49, "45-49 years",ifelse(Age <= 54, "50-54 years", ifelse(Age <= 59, "55-59 years",ifelse(Age <= 64, "60-64 years", ifelse(Age <= 69, "65-69 years",ifelse(Age <= 74, "70-74 years", ifelse(Age <= 79, "75-79 years",ifelse(Age <= 84, "80-84 years", ifelse(Age <= 89, "85-89 years", "90+ years")))))))))))))))))), levels = c("0-4 years", "5-9 years", "10-14 years", "15-19 years", "20-24 years", "25-29 years", "30-34 years", "35-39 years", "40-44 years", "45-49 years", "50-54 years", "55-59 years", "60-64 years", "65-69 years", "70-74 years", "75-79 years", "80-84 years", "85-89 years", "90+ years"))) %>% 
#   group_by(Area_name, Sex, `Age group`) %>% 
#   summarise(`2018` = sum(`2018`, na.rm = TRUE),
#             `2019` = sum(`2019`, na.rm = TRUE),
#             `2020` = sum(`2020`, na.rm = TRUE),
#             `2021` = sum(`2021`, na.rm = TRUE),
#             `2022` = sum(`2022`, na.rm = TRUE),
#             `2023` = sum(`2023`, na.rm = TRUE),
#             `2024` = sum(`2024`, na.rm = TRUE),
#             `2025` = sum(`2025`, na.rm = TRUE),
#             `2026` = sum(`2026`, na.rm = TRUE),
#             `2027` = sum(`2027`, na.rm = TRUE),
#             `2028` = sum(`2028`, na.rm = TRUE),
#             `2029` = sum(`2029`, na.rm = TRUE),
#             `2030` = sum(`2030`, na.rm = TRUE),
#             `2031` = sum(`2031`, na.rm = TRUE),
#             `2032` = sum(`2032`, na.rm = TRUE)) %>% 
#   gather(Year, Population, `2018`:`2032`, factor_key = TRUE) %>% 
#   mutate(Data_type = "Projected - WSCC Housing",
#          Age_band_type = "5 years") %>% 
#   ungroup()
# 
# WSCC_HH_projection_1832_10_years <- data.frame(Area_name = "Adur", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Adur", skip = 4, n_max = 186), check.names = FALSE) %>%   bind_rows(data.frame(Area_name = "Arun", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Arun", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Chichester", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Chichester", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Crawley", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Crawley", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Horsham", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Horsham", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Mid Sussex", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Mid_Sussex", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Worthing", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Worthing", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "West Sussex", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "WSx", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   filter(Sex %in% c("male", "female")) %>% 
#   mutate(Age = as.numeric(ifelse(Age == "90+", "90", Age)),
#          Sex = capwords(Sex, strict = TRUE)) %>% 
#   mutate(`Age group` = factor(ifelse(Age <= 9, "0-9 years", ifelse(Age <= 19, "10-19 years", ifelse(Age <= 29, "20-29 years", ifelse(Age <= 39, "30-39 years", ifelse(Age <= 49, "40-49 years", ifelse(Age <= 59, "50-59 years",ifelse(Age <= 69, "60-69 years", ifelse(Age <= 79, "70-79 years",ifelse(Age <= 89, "80-89 years", "90+ years"))))))))), levels = c("0-9 years", "10-19 years", "20-29 years", "30-39 years", "40-49 years", "50-59 years", "60-69 years", "70-79 years", "80-89 years", "90+ years"))) %>% 
#   group_by(Area_name, Sex, `Age group`) %>% 
#   summarise(`2018` = sum(`2018`, na.rm = TRUE),
#             `2019` = sum(`2019`, na.rm = TRUE),
#             `2020` = sum(`2020`, na.rm = TRUE),
#             `2021` = sum(`2021`, na.rm = TRUE),
#             `2022` = sum(`2022`, na.rm = TRUE),
#             `2023` = sum(`2023`, na.rm = TRUE),
#             `2024` = sum(`2024`, na.rm = TRUE),
#             `2025` = sum(`2025`, na.rm = TRUE),
#             `2026` = sum(`2026`, na.rm = TRUE),
#             `2027` = sum(`2027`, na.rm = TRUE),
#             `2028` = sum(`2028`, na.rm = TRUE),
#             `2029` = sum(`2029`, na.rm = TRUE),
#             `2030` = sum(`2030`, na.rm = TRUE),
#             `2031` = sum(`2031`, na.rm = TRUE),
#             `2032` = sum(`2032`, na.rm = TRUE)) %>% 
#   gather(Year, Population, `2018`:`2032`, factor_key = TRUE) %>% 
#   mutate(Data_type = "Projected - WSCC Housing",
#          Age_band_type = "10 years") %>% 
#   ungroup()
# 
# WSCC_HH_projection_1832_broad <- data.frame(Area_name = "Adur", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Adur", skip = 4, n_max = 186), check.names = FALSE) %>%   bind_rows(data.frame(Area_name = "Arun", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Arun", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Chichester", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Chichester", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Crawley", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Crawley", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Horsham", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Horsham", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Mid Sussex", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Mid_Sussex", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "Worthing", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Worthing", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   bind_rows(data.frame(Area_name = "West Sussex", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "WSx", skip = 4, n_max = 186), check.names = FALSE)) %>% 
#   filter(Sex %in% c("male", "female")) %>% 
#   mutate(Age = as.numeric(ifelse(Age == "90+", "90", Age)),
#          Sex = capwords(Sex, strict = TRUE)) %>% 
#   mutate(`Age group` = factor(ifelse(Age <= 17, "0-17 years", ifelse(Age <= 44, "18-44 years", ifelse(Age <= 64, "45-64 years", "65+ years"))), levels = c("0-17 years", "18-44 years", "45-64 years", "65+ years"))) %>% 
#   group_by(Area_name, Sex, `Age group`) %>% 
#   summarise(`2018` = sum(`2018`, na.rm = TRUE),
#             `2019` = sum(`2019`, na.rm = TRUE),
#             `2020` = sum(`2020`, na.rm = TRUE),
#             `2021` = sum(`2021`, na.rm = TRUE),
#             `2022` = sum(`2022`, na.rm = TRUE),
#             `2023` = sum(`2023`, na.rm = TRUE),
#             `2024` = sum(`2024`, na.rm = TRUE),
#             `2025` = sum(`2025`, na.rm = TRUE),
#             `2026` = sum(`2026`, na.rm = TRUE),
#             `2027` = sum(`2027`, na.rm = TRUE),
#             `2028` = sum(`2028`, na.rm = TRUE),
#             `2029` = sum(`2029`, na.rm = TRUE),
#             `2030` = sum(`2030`, na.rm = TRUE),
#             `2031` = sum(`2031`, na.rm = TRUE),
#             `2032` = sum(`2032`, na.rm = TRUE)) %>% 
#   gather(Year, Population, `2018`:`2032`, factor_key = TRUE) %>% 
#   mutate(Data_type = "Projected - WSCC Housing",
#          Age_band_type = "broad years") %>% 
#   ungroup()

# If you have local projections to add you can do

# WSCC_projections_housing_file <- WSCC_HH_projection_1832_quinary %>% 
#   bind_rows(WSCC_HH_projection_1832_10_years) %>% 
#   bind_rows(WSCC_HH_projection_1832_broad) %>% 
#   left_join(Areas, by = c("Area_name" = "Area_Name")) %>% 
#   select(Area_name, Area_Code, Area_Type, Sex, `Age group`,Age_band_type, Year, Population, Data_type) %>% 
#   rename(Age_group = `Age group`,
#          Area_Name = Area_name) %>% 
#   bind_rows(Areas_projections_file)

# Current MYE Population ####

ONS_mye_SYOA <- data.frame(DATE_NAME = double(), GEOGRAPHY = double(),GEOGRAPHY_NAME = character(), GEOGRAPHY_CODE = character(), GENDER_NAME = character(),C_AGE_NAME = character(),  MEASURES_NAME = character(), OBS_VALUE = double(), OBS_STATUS_NAME = character(), RECORD_COUNT = double())

for(i in 0:floor(as.numeric(read_csv(url(paste0("http://www.nomisweb.co.uk/api/v01/dataset/NM_2002_1.data.csv?geography=",paste(as.numeric(Chosen_area_codes$GEOGRAPHY), collapse = ","),"&date=latestMINUS6-latest&gender=1,2&c_age=101...191&measures=20100&select=record_count&recordlimit=1")), col_types = cols(RECORD_COUNT = col_double())))/25000)){
  df <- read_csv(url(paste0("http://www.nomisweb.co.uk/api/v01/dataset/NM_2002_1.data.csv?geography=",paste(as.numeric(Chosen_area_codes$GEOGRAPHY), collapse = ","),"&date=latestMINUS7-latest&gender=1,2&c_age=101...191&measures=20100&select=date_name,geography,geography_name,geography_code,gender_name,c_age_name,measures_name,obs_value,obs_status_name,record_count&recordoffset=", 25000 * i)), col_types = cols(DATE_NAME = col_double(), GEOGRAPHY = col_double(),GEOGRAPHY_NAME = col_character(),  GEOGRAPHY_CODE = col_character(), GENDER_NAME = col_character(),C_AGE_NAME = col_character(), MEASURES_NAME = col_character(),OBS_VALUE = col_double(), RECORD_COUNT = col_double()))
  
  ONS_mye_SYOA <- ONS_mye_SYOA %>% 
    bind_rows(df)
  }

ONS_mye_ccg_SYOA <- data.frame(DATE_NAME = double(), GEOGRAPHY = double(),GEOGRAPHY_NAME = character(), GEOGRAPHY_CODE = character(), GENDER_NAME = character(),C_AGE_NAME = character(),  MEASURES_NAME = character(), OBS_VALUE = double(), OBS_STATUS_NAME = character(), RECORD_COUNT = double())

for(i in 0:floor(as.numeric(read_csv(url(paste0("http://www.nomisweb.co.uk/api/v01/dataset/NM_2010_1.data.csv?geography=",paste(as.numeric(Chosen_area_codes$GEOGRAPHY), collapse = ","),"&date=latestMINUS6-latest&gender=1,2&c_age=101...191&measures=20100&select=record_count&recordlimit=1")), col_types = cols(RECORD_COUNT = col_double())))/25000)){
  df <- read_csv(url(paste0("http://www.nomisweb.co.uk/api/v01/dataset/NM_2010_1.data.csv?geography=",paste(as.numeric(Chosen_area_codes$GEOGRAPHY), collapse = ","),"&date=latestMINUS7-latest&gender=1,2&c_age=101...191&measures=20100&select=date_name,geography,geography_name,geography_code,gender_name,c_age_name,measures_name,obs_value,obs_status_name,record_count&recordoffset=", 25000 * i)), col_types = cols(DATE_NAME = col_double(), GEOGRAPHY = col_double(),GEOGRAPHY_NAME = col_character(),  GEOGRAPHY_CODE = col_character(), GENDER_NAME = col_character(),C_AGE_NAME = col_character(), MEASURES_NAME = col_character(),OBS_VALUE = col_double(), RECORD_COUNT = col_double()))
  
  ONS_mye_ccg_SYOA <- ONS_mye_ccg_SYOA %>% 
    bind_rows(df)
  
}


NOMIS_mye_df <- ONS_mye_SYOA %>% 
  bind_rows(ONS_mye_ccg_SYOA) %>% 
  rename(AREA_CODE = GEOGRAPHY_CODE,
         AREA_NAME = GEOGRAPHY_NAME,
         SEX = GENDER_NAME) %>% 
  mutate(SEX = ifelse(SEX == "Male", "males", ifelse(SEX == "Female", "females", NA))) %>% 
  mutate(AGE_GROUP = gsub("Age ", "", C_AGE_NAME)) %>% 
  select(AREA_CODE, AREA_NAME, SEX, AGE_GROUP, DATE_NAME, OBS_VALUE) %>% 
  spread(DATE_NAME, OBS_VALUE) %>% 
  mutate(AGE_GROUP  = ifelse(AGE_GROUP == "Aged 90+", "90 and over", AGE_GROUP))

setdiff(Chosen_area_codes$Area_Name, NOMIS_mye_df$AREA_NAME)

ONS_MYE_quinary <- NOMIS_mye_df %>% 
  filter(AGE_GROUP != "All ages") %>% 
  mutate(Age = as.numeric(gsub(" and over", "", AGE_GROUP))) %>% 
  mutate(`Age group` = ifelse(Age <= 4, "0-4 years", ifelse(Age <= 9, "5-9 years", ifelse(Age <= 14, "10-14 years", ifelse(Age <= 19, "15-19 years", ifelse(Age <= 24, "20-24 years", ifelse(Age <= 29, "25-29 years",ifelse(Age <= 34, "30-34 years", ifelse(Age <= 39, "35-39 years",ifelse(Age <= 44, "40-44 years", ifelse(Age <= 49, "45-49 years",ifelse(Age <= 54, "50-54 years", ifelse(Age <= 59, "55-59 years",ifelse(Age <= 64, "60-64 years", ifelse(Age <= 69, "65-69 years",ifelse(Age <= 74, "70-74 years", ifelse(Age <= 79, "75-79 years",ifelse(Age <= 84, "80-84 years", ifelse(Age <= 89, "85-89 years", "90+ years"))))))))))))))))))) %>% 
  group_by(AREA_NAME, AREA_CODE, SEX, `Age group`) %>% 
  summarise(`2011` = sum(`2011`, na.rm = TRUE),
            `2012` = sum(`2012`, na.rm = TRUE),
            `2013` = sum(`2013`, na.rm = TRUE),
            `2014` = sum(`2014`, na.rm = TRUE),
            `2015` = sum(`2015`, na.rm = TRUE),
            `2016` = sum(`2016`, na.rm = TRUE),
            `2017` = sum(`2017`, na.rm = TRUE),
            `2018` = sum(`2018`, na.rm = TRUE)) %>% 
  ungroup() %>% 
  mutate(SEX = capwords(SEX)) %>% 
  rename(Area_name = AREA_NAME,
         Area_code = AREA_CODE,
         Sex = SEX) %>% 
  mutate(Sex = ifelse(Sex == "Females", "Female", ifelse(Sex == "Males", "Male", Sex))) %>% 
  gather(Year, Population, `2011`:`2018`, factor_key = TRUE) %>% 
  mutate(Data_type = "Estimates - ONS",
         Age_band_type = "5 years")

ONS_MYE_10_year <- NOMIS_mye_df %>% 
  filter(AGE_GROUP != "All ages") %>% 
  mutate(Age = as.numeric(gsub(" and over", "", AGE_GROUP))) %>% 
  mutate(`Age group` = ifelse(Age <= 9, "0-9 years", ifelse(Age <= 19, "10-19 years", ifelse(Age <= 29, "20-29 years", ifelse(Age <= 39, "30-39 years", ifelse(Age <= 49, "40-49 years", ifelse(Age <= 59, "50-59 years",ifelse(Age <= 69, "60-69 years", ifelse(Age <= 79, "70-79 years",ifelse(Age <= 89, "80-89 years", "90+ years")))))))))) %>% 
  group_by(AREA_NAME, AREA_CODE, SEX, `Age group`) %>% 
  summarise(`2011` = sum(`2011`, na.rm = TRUE),
            `2012` = sum(`2012`, na.rm = TRUE),
            `2013` = sum(`2013`, na.rm = TRUE),
            `2014` = sum(`2014`, na.rm = TRUE),
            `2015` = sum(`2015`, na.rm = TRUE),
            `2016` = sum(`2016`, na.rm = TRUE),
            `2017` = sum(`2017`, na.rm = TRUE),
            `2018` = sum(`2018`, na.rm = TRUE)) %>% 
  ungroup() %>% 
  mutate(SEX = capwords(SEX)) %>% 
  rename(Area_name = AREA_NAME,
         Area_code = AREA_CODE,
         Sex = SEX) %>% 
  mutate(Sex = ifelse(Sex == "Females", "Female", ifelse(Sex == "Males", "Male", Sex))) %>% 
  gather(Year, Population, `2011`:`2018`, factor_key = TRUE) %>% 
  mutate(Data_type = "Estimates - ONS",
         Age_band_type = "10 years")

ONS_MYE_broad <- NOMIS_mye_df %>% 
  filter(AGE_GROUP != "All ages") %>% 
  mutate(Age = as.numeric(gsub(" and over", "", AGE_GROUP))) %>% 
  mutate(`Age group` = ifelse(Age <= 17, "0-17 years", ifelse(Age <= 44, "18-44 years", ifelse(Age <= 64, "45-64 years", "65+ years")))) %>% 
  group_by(AREA_NAME, AREA_CODE, SEX, `Age group`) %>% 
  summarise(`2011` = sum(`2011`, na.rm = TRUE),
            `2012` = sum(`2012`, na.rm = TRUE),
            `2013` = sum(`2013`, na.rm = TRUE),
            `2014` = sum(`2014`, na.rm = TRUE),
            `2015` = sum(`2015`, na.rm = TRUE),
            `2016` = sum(`2016`, na.rm = TRUE),
            `2017` = sum(`2017`, na.rm = TRUE),
            `2018` = sum(`2018`, na.rm = TRUE)) %>% 
  ungroup() %>% 
  mutate(SEX = capwords(SEX)) %>% 
  rename(Area_name = AREA_NAME,
         Area_code = AREA_CODE,
         Sex = SEX) %>% 
  mutate(Sex = ifelse(Sex == "Females", "Female", ifelse(Sex == "Males", "Male", Sex))) %>% 
  gather(Year, Population, `2011`:`2018`, factor_key = TRUE) %>% 
  mutate(Data_type = "Estimates - ONS",
         Age_band_type = "broad years")

Areas_estimates_file <- ONS_MYE_quinary %>% 
  bind_rows(ONS_MYE_10_year) %>% 
  bind_rows(ONS_MYE_broad) %>% 
  left_join(Areas, by = c("Area_name" = "Area_Name")) %>% 
  select(Area_name, Area_Code, Area_Type, Sex, `Age group`,Age_band_type, Year, Population, Data_type) %>% 
  rename(Age_group = `Age group`,
         Area_Name = Area_name)


Areas_estimates_file <- Areas_estimates_file %>% 
  filter(!(Area_Type == "Clinical Commissioning Group (2018)" & Year == "2018"))

Areas_projections_file_ccg_18 <- Areas_projections_file %>% 
  filter(Area_Type == "Clinical Commissioning Group (2018)" & Year == "2018")

Areas_projections_file <- Areas_projections_file %>% 
  filter(Year != "2018")

Area_population_df <- Areas_estimates_file %>% 
  bind_rows(Areas_projections_file) %>% 
  bind_rows(Areas_projections_file_ccg_18) %>% 
  mutate(Source = "Office for National Statistics")

rm(df, i, ONS_mye_SYOA, ONS_mye_ccg_SYOA, ONS_projections_SYOA, ONS_ccg_projections_df, ONS_MYE_10_year, ONS_MYE_broad, ONS_MYE_quinary, ONS_projection_1941_10_year, ONS_projection_1941_broad, ONS_projection_1941_quinary)

write.csv(Area_population_df, file = "./Projecting-Health/Area_population_df.csv", row.names = FALSE)

}

Eng <- Area_population_df %>%
  filter(Area_Name == "England") %>%
  filter(Age_band_type == "5 years") %>%
  group_by(Year) %>%
  summarise(Population = sum(Population, na.rm = TRUE))


