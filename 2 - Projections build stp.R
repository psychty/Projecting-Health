options(java.parameters = "-Xmx2048m")

library(easypackages)

# This uses the easypackages package to load several libraries at once. Note: it should only be used when you are confident that all packages are installed as it will be more difficult to spot load errors compared to loading each one individually.
libraries(c("readxl", "readr", "plyr", "dplyr", "ggplot2", "png", "tidyverse", "reshape2", "scales", "viridis", "rgdal", "officer", "flextable", "tmaptools", "lemon", "fingertipsR", "PHEindicatormethods", "xlsx"))

# If you have downloaded/cloned the github repo for this project you will need to make sure the filepath is recorded in the object github_repo_dir
github_repo_dir <- "~/Documents/Repositories/Projecting-Health"

# You must specify a set of chosen areas.

if(!(exists("Areas_to_include"))){
  print("There are no areas defined. Please create or load 'Areas_to_include' which is a character string of chosen areas.")
}

# Areas_to_include <- c("Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "NHS Coastal West Sussex CCG","West Sussex", "NHS Crawley CCG", "NHS Horsham and Mid Sussex CCG")

#Areas_to_include <- c("Basingstoke and Deane","East Hampshire","Eastleigh","Fareham","Gosport","Hart","Havant","New Forest","Rushmoor","Test Valley","Hampshire", "Winchester", "NHS North East Hampshire and Farnham CCG", "NHS North Hampshire CCG", "NHS South Eastern Hampshire CCG", "NHS West Hampshire CCG", "NHS Fareham and Gosport CCG", "Portsmouth")


Areas_to_include <- c("Eastbourne", "Hastings", "Lewes","Rother", "Wealden","Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "Brighton and Hove", "NHS Brighton and Hove CCG", "NHS Coastal West Sussex CCG", "NHS Crawley CCG","NHS Eastbourne, Hailsham and Seaford CCG", "NHS Hastings and Rother CCG","NHS High Weald Lewes Havens CCG", "NHS Horsham and Mid Sussex CCG", "West Sussex", "East Sussex", "England")

if(!(file.exists("./Projecting-Health/Area_population_df.csv"))){
  print("Area_population_df is not available, it will be built using the 'Areas_to_include' object")
  source(paste0(github_repo_dir,"/Get data - mye and projections.R"))
  
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

Area_population_df <- Area_population_df %>% 
  group_by(Area_Name, Age_band_type, Year, Sex) %>% 
  mutate(All_age_population = sum(Population, na.rm = TRUE))

Counties_UA_selected <- Lookup %>% 
  filter(Area_Name %in% Areas_to_include & Area_Type %in% c("County", "Unitary Authority","Country")) 

Districts_UA_selected <- Lookup %>% 
  filter(Area_Name %in% Areas_to_include & Area_Type %in% c("District", "Unitary Authority","Country")) 

# Create a modified theme for line graphs (i.e. put the legend at the bottom)
ph_theme = function(){
  theme( 
    legend.position = "bottom", legend.title = element_text(colour = "#000000", size = 10), 
    legend.background = element_rect(fill = "#ffffff"), 
    legend.key = element_rect(fill = "#ffffff", colour = "#E2E2E3"), 
    legend.text = element_text(colour = "#000000", size = 8), 
    plot.background = element_rect(fill = "white", colour = "#E2E2E3"), 
    panel.background = element_rect(fill = "white"), 
    axis.text = element_text(colour = "#000000", size = 8), 
    plot.title = element_text(colour = "#327d9c", face = "bold", size = 12, vjust = 1), 
    axis.title = element_text(colour = "#327d9c", face = "bold", size = 8),     
    panel.grid.major.x = element_line(colour = "#E2E2E3", linetype = "longdash"), 
    panel.grid.minor.x = element_blank(), 
    panel.grid.major.y = element_blank(), 
    panel.grid.minor.y = element_blank(), 
    strip.text = element_text(colour = "white"), 
    strip.background = element_rect(fill = "#327d9c"), 
    axis.ticks = element_line(colour = "#327d9c") 
  ) 
}

# Axis labels can be formated in a number of ways. We need to have the absolute number (-10000 people would look weird) and it would also be good to include the comma separator (10,000 is easier to read than 10000).
# Using ggplot you can format both of these (by using labels = abs and labels = comma, respectively) but you cannot do them at the same time.

# So we need to create a function that does both
abs_comma <- function (x, ...) {
  format(abs(x), ..., big.mark = ",", scientific = FALSE, trim = TRUE)}

# You can also hijack the percent function and return axes labels with "%" as a suffix (or you could do any suffix)
abs_percent <- function (x, ...) {
  if (length(x) == 0) 
    return(character())
  paste0(comma(abs(x) * 100), "%")}

# Age groups ####
Area_population_broad_sex <- Area_population_df %>% 
  filter(Age_band_type == "broad years") %>% 
  mutate(Age_group = factor(Age_group, levels = c("0-17 years", "18-44 years", "45-64 years", "65+ years", "Total"))) %>% 
  arrange(Area_Name, Sex, Age_group, Year) %>% 
  group_by(Area_Name, Area_Code, Sex, Year) %>% 
  mutate(Total_population = sum(Population, na.rm = TRUE)) %>% 
  ungroup()

Area_population_broad_combined <- Area_population_broad_sex %>% 
  group_by(Area_Name, Area_Code, Area_Type, Age_band_type, Age_group, Data_type, Year) %>% 
  summarise(Sex = "Total", Population = sum(Population, na.rm = TRUE)) %>% 
  group_by(Area_Name, Area_Code, Area_Type, Age_band_type, Data_type, Year) %>% 
  mutate(Total_population = sum(Population, na.rm = TRUE)) %>% 
  ungroup() %>% 
  bind_rows(Area_population_broad_sex) %>% 
  mutate(Proportion = Population / Total_population)

# Sixty_five_plus <- Area_population_broad_combined %>% 
#   select(-c(Total_population, Proportion)) %>% 
#   filter(Age_group %in% c("65+ years")) %>% 
#   spread(Sex, Population)

Older_age_broad_projections <- Area_population_df %>% 
  filter(Age_band_type == "5 years") %>% 
  filter(Age_group %in% c("65-69 years", "70-74 years", "75-79 years", "80-84 years", "85-89 years", "90+ years")) %>% 
  mutate(Age_group = factor(ifelse(Age_group %in% c("65-69 years", "70-74 years"), "65-74 years", ifelse(Age_group %in% c("75-79 years", "80-84 years"), "75-84 years", "85+ years")), levels = c("65-74 years", "75-84 years", "85+ years"))) %>%
  ungroup() %>% 
  select(-Age_band_type) %>% 
  group_by(Area_Name, Area_Code, Area_Type, Age_group, Sex, Data_type, Year) %>% 
  summarise(Population = sum(Population, na.rm = TRUE))
 
# area_x = Older_age_broad_projections %>% 
#   filter(Area_Name == "Adur") %>% 
#   mutate(Year = as.character(Year))
# 
# ggplot(area_x, aes(x = Year, y = Population, fill = Age_group)) +
#   geom_bar(stat = "identity") +
#   ph_theme() +
#   labs(x= "Year",
#        y = "Population") +
#   scale_y_continuous( labels = comma) +
#   scale_fill_manual(values = c("#FD0000","#404040","#A6A6A6"), 
#                     breaks = c("65-74 years", "75-84 years", "85+ years"), 
#                     name = "Age band") +
#   geom_vline(xintercept = 6.5, 
#              lty = "dotted", 
#              colour = "#000000", 
#              lwd = .6) +
#   annotate(geom = "rect", 
#            xmin = 6.5, 
#            xmax = 31.5, 
#            ymin = 0, 
#            ymax = Inf, 
#            fill = "#e7e7e7", 
#            alpha = 0.25) +
#   facet_rep_wrap(. ~ Sex, nrow = 2, repeat.tick.labels = TRUE)

# paste0("The 2016 based population projections estimated that in 2018, there are ", format(subset(over_65_projected, Year == "2018", select = "Population_rounded"), big.mark = ","), " residents aged 65 years and over. Over the next 10 years it is anticipated that the population aged 65 years and older is going to increase by ", round(((subset(over_65_projected, Year == "2028", select = "Population_rounded")-subset(over_65_projected, Year == "2018", select = "Population_rounded"))/subset(over_65_projected, Year == "2018", select = "Population_rounded"))*100,0), "% to ", format(subset(over_65_projected, Year == "2028", select = "Population_rounded"),big.mark = ","), " residents and by 2038, there is expected to be an extra ", format(subset(over_65_projected, Year == "2038", select = "Population_rounded")-subset(over_65_projected, Year == "2018", select = "Population_rounded"),big.mark = ",")," residents aged 65 years and over compared to today.")

paste0("Long-term subnational population projections are an indication of the future trends in population by age and sex over the next 25 years. They are trend-based projections, which means assumptions for future levels of births, deaths and migration are based on observed levels mainly over the previous five years.")

working_age_projections <- Area_population_df %>% 
  filter(Age_group %in% c("18-44 years", "45-64 years")) %>%
  ungroup() %>% 
  select(-Age_band_type) %>% 
  group_by(Area_Name, Area_Code, Area_Type, Sex, Data_type, Year) %>% 
  summarise(Population = sum(Population, na.rm = TRUE), Age_group = "Working age")

older_age_projections <- Area_population_df %>% 
  filter(Age_group %in% c("65+ years")) %>%
  ungroup() %>% 
  select(-Age_band_type) %>% 
  group_by(Area_Name, Area_Code, Area_Type, Sex, Data_type, Year) %>% 
  summarise(Population = sum(Population, na.rm = TRUE), Age_group = "Old age")

OADR <- working_age_projections %>% 
                        bind_rows(older_age_projections) %>% 
                        spread(Age_group, Population) %>% 
                        mutate(Ratio_per_1000 = `Old age` / `Working age` * 1000) %>% 
  ungroup()

OADR <- as.data.frame(OADR)

rm(Area_population_broad_sex, Area_population_broad_combined)

paste0("This is the ratio of older dependents (people older than 64) to the working-age population (those ages 16-64). Data are shown as the proportion of dependents per 100 working-age population. ONS present it as a rate per 1,000")

# paste0("In West Sussex, in 2018, there is an estimated ", round(subset(OADR, Year == "2018", select = "Ratio_per_1000"),0), " residents aged 65+ for every 1,000 working age (16-64 years) population. This is set to increase over the next decade to ", round(subset(OADR, Year == "2028", select = "Ratio_per_1000"),0), " residents aged 65+ for every 1,000 16-64 year olds in 2028, and by 2038, there is anticipated to be ", round(subset(OADR, Year == "2038", select = "Ratio_per_1000"),0), " residents aged 65+ for every 1,000 16-64 years population; this is a ratio of one older person for every two working age residents.")

paste0("This statistic should be considered with caution. It does not account for those who are aged 65 and over who are still in work, nor does it account for those who are working age but out of work. Furthermore, the state pension age will likely increase over time meaning that people will work longer.")

paste0("What it does say is that if the population changes as projected, by 2041, there will be one pension aged person for every two working age people.")

# title = "West Sussex resident population aged 65+",
# subtitle = "Estimates to 2016; Projected to 2041",

# add in the older_age_broad_projection data to excel file to build the figure which changes dynamically

# Build life expectancy (all areas) dataframe for inserting into excel file (doesnt have to be run each time).

HALE_gbd <- read_csv("./Life and Health Expectancies/GBD_HALE_LAs_2007_2017.csv", col_types = cols(measure = col_character(),  location = col_character(),  sex = col_character(),  age = col_character(),  metric = col_character(),  year = col_double(),  val = col_double(),  upper = col_double(),  lower = col_double())) 

LE_gbd <- read_csv("./Life and Health Expectancies/GBD_LE_LAs_2007_2017.csv", col_types = cols(measure = col_character(),  location = col_character(),  sex = col_character(),  age = col_character(),  metric = col_character(),  year = col_double(),  val = col_double(),  upper = col_double(),  lower = col_double()))

gbd_LE_tables <- HALE_gbd %>% 
  bind_rows(LE_gbd) %>% 
  rename(value = val) %>% 
  rename(area = location) %>% 
  filter(age == "<1 year") %>% 
  select(-metric)

selected_areas_gbd_LE_tables <- HALE_gbd %>% 
  bind_rows(LE_gbd) %>% 
  rename(value = val) %>% 
  rename(area = location) %>% 
  filter(age != "All Ages") %>% 
  filter(sex == "Both") %>% 
  filter(area %in% Areas_to_include)%>% 
  select(-metric)

write.csv(gbd_LE_tables, "./Life and Health Expectancies/gbd_LE_tables.csv", row.names = FALSE)

rm(HALE_gbd, LE_gbd)

# The Global Burden of Disease Study 2017 (GBD 2017), coordinated by the Institute for Health Metrics and Evaluation (IHME), estimated the burden of diseases, injuries, and risk factors for 195 countries and territories, and at the subnational level for a subset of countries.

# Estimates for healthy life expectancy (HALE) by age and sex are available from the GBD Results Tool for 1990-2017.

# Healthy life expectancy, or HALE, is a measure of average population health summarizing both mortality and non‐fatal outcomes. HALE is used for comparisons of health across countries or for measuring change over time. These comparisons can shed light on key questions about how morbidity worsens or improves as mortality declines.

# HALE was calculated by extending the conventional life table that is used to translate a schedule of agespecific death rates into estimates of life expectancy at different ages. Information on the average level of health experienced over each age interval was incorporated into the life table.

# ONS ####

Pop_in_deprived_areas <- fingertips_data(IndicatorID = 1730, AreaTypeID = c(101,102)) %>% 
  filter(Timeperiod == 2014) %>% 
  select(AreaCode,Value, LowerCI95.0limit, UpperCI95.0limit, Count) %>% 
  rename(Proportion_in_deprived = Value,
         Prop_in_deprived_lci = LowerCI95.0limit,
         Prop_in_deprived_uci = UpperCI95.0limit,
         N_in_deprived = Count)

Dep <- fingertips_data(91872, AreaTypeID = c(101,102)) %>% 
  select(AreaCode, AreaName,Value) %>% 
  rename(Deprivation_score = Value) %>% 
  left_join(Pop_in_deprived_areas, by = "AreaCode") %>% 
  unique()

rm(Pop_in_deprived_areas)

# Life expectancy at birth is a measure of the average number of years a person would expect to live based on contemporary mortality rates. For a particular area and time period, it is an estimate of the average number of years a newborn baby would survive if he or she experienced the age-specific mortality rates for that area and time period throughout his or her life.

LE_ONS <- fingertips_data(90366, AreaTypeID = c(101,102)) %>% 
  bind_rows(fingertips_data(91102, AreaTypeID = c(101,102))) %>% 
  select(IndicatorName, AreaCode, AreaName,AreaType,Sex,Age,Timeperiod,Value,LowerCI95.0limit,UpperCI95.0limit)

LE_at_birth_UTLA_rank <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  group_by(AreaType, Sex, Age, Timeperiod) %>% 
  arrange(Value) %>% 
  mutate(Rank =  rank(Value, ties.method = "first")) %>%
  ungroup() %>% 
  filter(IndicatorName == "Life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " rank")) %>% 
  select(AreaName, AreaCode, Sex, Rank) %>% 
  spread(key = Sex, value = Rank)

LE_at_birth_UTLA_val <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName))) %>% 
  select(AreaName, AreaCode, Sex, Value) %>% 
  spread(key = Sex, value = Value)

LE_at_birth_UTLA_lci <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " lci")) %>% 
  select(AreaName, AreaCode, Sex, LowerCI95.0limit) %>% 
  spread(key = Sex, value = LowerCI95.0limit)

LE_at_birth_UTLA_uci <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " uci")) %>% 
  select(AreaName, AreaCode, Sex, UpperCI95.0limit) %>% 
  spread(key = Sex, value = UpperCI95.0limit)

LE_at_birth_UTLA <- LE_at_birth_UTLA_val %>% 
  left_join(LE_at_birth_UTLA_lci, by = c("AreaCode", "AreaName")) %>% 
  left_join(LE_at_birth_UTLA_uci, by = c("AreaCode", "AreaName")) %>%
  left_join(LE_at_birth_UTLA_rank, by = c("AreaCode", "AreaName")) %>% 
  select(AreaName,AreaCode,`Female life expectancy at birth`,`Female life expectancy at birth lci`,`Female life expectancy at birth uci`,`Female life expectancy at birth rank`, `Male life expectancy at birth`,`Male life expectancy at birth lci`,`Male life expectancy at birth uci`,`Male life expectancy at birth rank`)

rm(LE_at_birth_UTLA_val, LE_at_birth_UTLA_lci, LE_at_birth_UTLA_uci, LE_at_birth_UTLA_rank)

LE_at_65_UTLA_rank <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  group_by(AreaType, Sex, Age, Timeperiod) %>% 
  arrange(Value) %>% 
  mutate(Rank =  rank(Value, ties.method = "first")) %>%
  ungroup() %>% 
  filter(IndicatorName == "Life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " rank")) %>% 
  select(AreaName, AreaCode, Sex, Rank) %>% 
  spread(key = Sex, value = Rank)

LE_at_65_UTLA_val <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName))) %>% 
  select(AreaName, AreaCode, Sex, Value) %>% 
  spread(key = Sex, value = Value)

LE_at_65_UTLA_lci <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " lci")) %>% 
  select(AreaName, AreaCode, Sex, LowerCI95.0limit) %>% 
  spread(key = Sex, value = LowerCI95.0limit)

LE_at_65_UTLA_uci <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " uci")) %>% 
  select(AreaName, AreaCode, Sex, UpperCI95.0limit) %>% 
  spread(key = Sex, value = UpperCI95.0limit)

LE_at_65_UTLA <- LE_at_65_UTLA_val %>% 
  left_join(LE_at_65_UTLA_lci, by = c("AreaCode", "AreaName")) %>% 
  left_join(LE_at_65_UTLA_uci, by = c("AreaCode", "AreaName")) %>%
  left_join(LE_at_65_UTLA_rank, by = c("AreaCode", "AreaName")) %>% 
  select(AreaName,AreaCode,`Female life expectancy at 65`,`Female life expectancy at 65 lci`,`Female life expectancy at 65 uci`,`Female life expectancy at 65 rank`, `Male life expectancy at 65`,`Male life expectancy at 65 lci`,`Male life expectancy at 65 uci`,`Male life expectancy at 65 rank`)

rm(LE_at_65_UTLA_val, LE_at_65_UTLA_lci, LE_at_65_UTLA_uci, LE_at_65_UTLA_rank)

LE_UTLA <- LE_at_birth_UTLA %>% 
  left_join(LE_at_65_UTLA, by = c("AreaCode", "AreaName")) %>% 
  left_join(Dep, by = c("AreaCode", "AreaName"))

rm(LE_at_birth_UTLA, LE_at_65_UTLA)

# Districts ##

LE_at_birth_LTLA_rank <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "District & UA") %>% 
  group_by(AreaType, Sex, Age, Timeperiod) %>% 
  arrange(Value) %>% 
  mutate(Rank =  rank(Value, ties.method = "first")) %>%
  ungroup() %>% 
  filter(IndicatorName == "Life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " rank")) %>% 
  select(AreaName, AreaCode, Sex, Rank) %>% 
  spread(key = Sex, value = Rank)

LE_at_birth_LTLA_val <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "District & UA") %>% 
  filter(IndicatorName == "Life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName))) %>% 
  select(AreaName, AreaCode, Sex, Value) %>% 
  spread(key = Sex, value = Value)

LE_at_birth_LTLA_lci <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "District & UA") %>% 
  filter(IndicatorName == "Life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " lci")) %>% 
  select(AreaName, AreaCode, Sex, LowerCI95.0limit) %>% 
  spread(key = Sex, value = LowerCI95.0limit)

LE_at_birth_LTLA_uci <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "District & UA") %>% 
  filter(IndicatorName == "Life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " uci")) %>% 
  select(AreaName, AreaCode, Sex, UpperCI95.0limit) %>% 
  spread(key = Sex, value = UpperCI95.0limit)

LE_at_birth_LTLA <- LE_at_birth_LTLA_val %>% 
  left_join(LE_at_birth_LTLA_lci, by = c("AreaCode", "AreaName")) %>% 
  left_join(LE_at_birth_LTLA_uci, by = c("AreaCode", "AreaName")) %>%
  left_join(LE_at_birth_LTLA_rank, by = c("AreaCode", "AreaName")) %>% 
  select(AreaName,AreaCode,`Female life expectancy at birth`,`Female life expectancy at birth lci`,`Female life expectancy at birth uci`,`Female life expectancy at birth rank`, `Male life expectancy at birth`,`Male life expectancy at birth lci`,`Male life expectancy at birth uci`,`Male life expectancy at birth rank`)

rm(LE_at_birth_LTLA_val, LE_at_birth_LTLA_lci, LE_at_birth_LTLA_uci, LE_at_birth_LTLA_rank)

LE_at_65_LTLA_rank <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "District & UA") %>% 
  group_by(AreaType, Sex, Age, Timeperiod) %>% 
  arrange(Value) %>% 
  mutate(Rank =  rank(Value, ties.method = "first")) %>%
  ungroup() %>% 
  filter(IndicatorName == "Life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " rank")) %>% 
  select(AreaName, AreaCode, Sex, Rank) %>% 
  spread(key = Sex, value = Rank)

LE_at_65_LTLA_val <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "District & UA") %>% 
  filter(IndicatorName == "Life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName))) %>% 
  select(AreaName, AreaCode, Sex, Value) %>% 
  spread(key = Sex, value = Value)

LE_at_65_LTLA_lci <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "District & UA") %>% 
  filter(IndicatorName == "Life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " lci")) %>% 
  select(AreaName, AreaCode, Sex, LowerCI95.0limit) %>% 
  spread(key = Sex, value = LowerCI95.0limit)

LE_at_65_LTLA_uci <- LE_ONS %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "District & UA") %>% 
  filter(IndicatorName == "Life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " uci")) %>% 
  select(AreaName, AreaCode, Sex, UpperCI95.0limit) %>% 
  spread(key = Sex, value = UpperCI95.0limit)

LE_at_65_LTLA <- LE_at_65_LTLA_val %>% 
  left_join(LE_at_65_LTLA_lci, by = c("AreaCode", "AreaName")) %>% 
  left_join(LE_at_65_LTLA_uci, by = c("AreaCode", "AreaName")) %>%
  left_join(LE_at_65_LTLA_rank, by = c("AreaCode", "AreaName")) %>% 
  select(AreaName,AreaCode,`Female life expectancy at 65`,`Female life expectancy at 65 lci`,`Female life expectancy at 65 uci`,`Female life expectancy at 65 rank`, `Male life expectancy at 65`,`Male life expectancy at 65 lci`,`Male life expectancy at 65 uci`,`Male life expectancy at 65 rank`)

rm(LE_at_65_LTLA_val, LE_at_65_LTLA_lci, LE_at_65_LTLA_uci, LE_at_65_LTLA_rank)

LE_LTLA <- LE_at_birth_LTLA %>% 
  left_join(LE_at_65_LTLA, by = c("AreaCode", "AreaName")) %>% 
  left_join(Dep, by = c("AreaCode", "AreaName"))

rm(LE_at_birth_LTLA, LE_at_65_LTLA)

# Dep_wards <- fingertips_data(91872, AreaTypeID = 8) %>% 
#   filter(AreaType == "Ward")

Selected_LE_ONS_ts <- LE_ONS %>% 
  group_by(AreaType, Sex, Age, Timeperiod) %>% 
  arrange(Value) %>% 
  mutate(Rank =  rank(Value, ties.method = "first")) %>% 
  filter(AreaName %in% Areas_to_include) 

# HLE at birth is available only for the upper tier local authorities.
HLE_at_birth <- fingertips_data(90362, AreaTypeID = c(101,102)) %>% 
  select(IndicatorName, AreaCode, AreaName,AreaType,Sex,Age,Timeperiod,Value,LowerCI95.0limit,UpperCI95.0limit)

HLE_at_birth_UTLA_rank <- HLE_at_birth %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  group_by(AreaType, Sex, Age, Timeperiod) %>% 
  arrange(Value) %>% 
  mutate(Rank =  rank(Value, ties.method = "first")) %>%
  ungroup() %>% 
  filter(IndicatorName == "Healthy life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " rank")) %>% 
  select(AreaName, AreaCode, Sex, Rank) %>% 
  spread(key = Sex, value = Rank)

HLE_at_birth_UTLA_val <- HLE_at_birth %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Healthy life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName))) %>% 
  select(AreaName, AreaCode, Sex, Value) %>% 
  spread(key = Sex, value = Value)

HLE_at_birth_UTLA_lci <- HLE_at_birth %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Healthy life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " lci")) %>% 
  select(AreaName, AreaCode, Sex, LowerCI95.0limit) %>% 
  spread(key = Sex, value = LowerCI95.0limit)

HLE_at_birth_UTLA_uci <- HLE_at_birth %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Healthy life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " uci")) %>% 
  select(AreaName, AreaCode, Sex, UpperCI95.0limit) %>% 
  spread(key = Sex, value = UpperCI95.0limit)

HLE_at_birth_UTLA <- HLE_at_birth_UTLA_val %>% 
  left_join(HLE_at_birth_UTLA_lci, by = c("AreaCode", "AreaName")) %>% 
  left_join(HLE_at_birth_UTLA_uci, by = c("AreaCode", "AreaName")) %>%
  left_join(HLE_at_birth_UTLA_rank, by = c("AreaCode", "AreaName")) %>% 
  select(AreaName,AreaCode,`Female healthy life expectancy at birth`,`Female healthy life expectancy at birth lci`,`Female healthy life expectancy at birth uci`,`Female healthy life expectancy at birth rank`, `Male healthy life expectancy at birth`,`Male healthy life expectancy at birth lci`,`Male healthy life expectancy at birth uci`,`Male healthy life expectancy at birth rank`)

rm(HLE_at_birth_UTLA_val, HLE_at_birth_UTLA_lci, HLE_at_birth_UTLA_uci, HLE_at_birth_UTLA_rank)

HLE_at_birth_UTLA <- HLE_at_birth_UTLA %>% 
  left_join(Dep, by = c("AreaCode", "AreaName"))

selected_HLE_ONS_UTLA_ts <- HLE_at_birth %>% 
  group_by(AreaType, Sex, Age, Timeperiod) %>% 
  arrange(Value) %>% 
  mutate(Rank =  rank(Value, ties.method = "first")) %>% 
  filter(AreaName %in% Areas_to_include)

# Healthy Life Expectancy at birth is a measure of the average number of years a person would expect to live in good health based on contemporary mortality rates and prevalence of self-reported good health. The prevalence of good health is derived from responses to a survey question on general health. For a particular area and time period, it is an estimate of the average number of years a newborn baby would live in good general health if he or she experienced the age-specific mortality rates and prevalence of good health for that area and time period throughout his or her life. Figures are calculated from deaths from all causes, mid-year population estimates, and self-reported general health status, based on data aggregated over a three year period. Figures reflect the prevalence of good health and mortality among those living in an area in each time period, rather than what will be experienced throughout life among those born in the area. The figures are not therefore the number of years a baby born in the area could actually expect to live in good general health, both because the health prevalence and mortality rates of the area are likely to change in the future and because many of those born in the area will live elsewhere for at least some part of their lives.

# Slope inequalities 

# Life expectancy at birth is calculated for each deprivation decile of lower super output areas within each area and then the slope index of inequality (SII) is calculated based on these figures. The SII is a measure of the social gradient in life expectancy, i.e. how much life expectancy varies with deprivation. It takes account of health inequalities across the whole range of deprivation within each area and summarises this in a single number. This represents the range in years of life expectancy across the social gradient from most to least deprived, based on a statistical analysis of the relationship between life expectancy and deprivation across all deprivation deciles.

# Data source	Figures calculated by Public Health England using mortality data and mid-year population estimates from the Office for National Statistics and Index of Multiple Deprivation 2010 and 2015 (IMD 2010 / IMD 2015) scores from the Ministry of Housing, Communities and Local Government.

# The SII for England and for regions should not be considered as comparators for the local authority figures. The SII for England takes account of the full range of deprivation and mortality across the whole country. This does not therefore provide a suitable benchmark with which to compare local authority results, which take into account the range of deprivation and mortality within much smaller geographies.  

Slope_inequalities <- fingertips_data(92901, AreaTypeID = c(101,102)) %>% 
  bind_rows(fingertips_data(93190, AreaTypeID = c(101, 102))) %>% 
  filter(Timeperiod %in% c("2010 - 12", "2011 - 13", "2012 - 14", "2013 - 15", "2014 - 16", "2015 - 17")) %>%   
  select(IndicatorName, AreaCode, AreaName,AreaType,Sex,Age,Timeperiod,Value,LowerCI95.0limit,UpperCI95.0limit) 

Slope_inequalities_at_birth_UTLA_rank <- Slope_inequalities %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  group_by(AreaType, Sex, Age, Timeperiod) %>% 
  arrange(Value) %>% 
  mutate(Rank =  rank(Value, ties.method = "first")) %>%
  ungroup() %>% 
  filter(IndicatorName == "Inequality in life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " rank")) %>% 
  select(AreaName, AreaCode, Sex, Rank) %>% 
  spread(key = Sex, value = Rank)

Slope_inequalities_at_birth_UTLA_val <- Slope_inequalities %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Inequality in life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName))) %>% 
  select(AreaName, AreaCode, Sex, Value) %>% 
  spread(key = Sex, value = Value)

Slope_inequalities_at_birth_UTLA_lci <- Slope_inequalities %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Inequality in life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " lci")) %>% 
  select(AreaName, AreaCode, Sex, LowerCI95.0limit) %>% 
  spread(key = Sex, value = LowerCI95.0limit)

Slope_inequalities_at_birth_UTLA_uci <- Slope_inequalities %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Inequality in life expectancy at birth") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " uci")) %>% 
  select(AreaName, AreaCode, Sex, UpperCI95.0limit) %>% 
  spread(key = Sex, value = UpperCI95.0limit)

Slope_inequalities_at_birth_UTLA <- Slope_inequalities_at_birth_UTLA_val %>% 
  left_join(Slope_inequalities_at_birth_UTLA_lci, by = c("AreaCode", "AreaName")) %>% 
  left_join(Slope_inequalities_at_birth_UTLA_uci, by = c("AreaCode", "AreaName")) %>%
  left_join(Slope_inequalities_at_birth_UTLA_rank, by = c("AreaCode", "AreaName")) %>% 
  select(AreaName,AreaCode,`Female inequality in life expectancy at birth`,`Female inequality in life expectancy at birth lci`,`Female inequality in life expectancy at birth uci`,`Female inequality in life expectancy at birth rank`, `Male inequality in life expectancy at birth`,`Male inequality in life expectancy at birth lci`,`Male inequality in life expectancy at birth uci`,`Male inequality in life expectancy at birth rank`)

rm(Slope_inequalities_at_birth_UTLA_val, Slope_inequalities_at_birth_UTLA_lci, Slope_inequalities_at_birth_UTLA_uci, Slope_inequalities_at_birth_UTLA_rank)

Slope_inequalities_at_65_UTLA_rank <- Slope_inequalities %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  group_by(AreaType, Sex, Age, Timeperiod) %>% 
  arrange(Value) %>% 
  mutate(Rank =  rank(Value, ties.method = "first")) %>%
  ungroup() %>% 
  filter(IndicatorName == "Inequality in life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " rank")) %>% 
  select(AreaName, AreaCode, Sex, Rank) %>% 
  spread(key = Sex, value = Rank)

Slope_inequalities_at_65_UTLA_val <- Slope_inequalities %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Inequality in life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName))) %>% 
  select(AreaName, AreaCode, Sex, Value) %>% 
  spread(key = Sex, value = Value)

Slope_inequalities_at_65_UTLA_lci <- Slope_inequalities %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Inequality in life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " lci")) %>% 
  select(AreaName, AreaCode, Sex, LowerCI95.0limit) %>% 
  spread(key = Sex, value = LowerCI95.0limit)

Slope_inequalities_at_65_UTLA_uci <- Slope_inequalities %>% 
  filter(Timeperiod == "2015 - 17") %>% 
  filter(AreaType == "County & UA") %>% 
  filter(IndicatorName == "Inequality in life expectancy at 65") %>% 
  mutate(Sex = paste0(Sex, " ", tolower(IndicatorName), " uci")) %>% 
  select(AreaName, AreaCode, Sex, UpperCI95.0limit) %>% 
  spread(key = Sex, value = UpperCI95.0limit)

Slope_inequalities_at_65_UTLA <- Slope_inequalities_at_65_UTLA_val %>% 
  left_join(Slope_inequalities_at_65_UTLA_lci, by = c("AreaCode", "AreaName")) %>% 
  left_join(Slope_inequalities_at_65_UTLA_uci, by = c("AreaCode", "AreaName")) %>%
  left_join(Slope_inequalities_at_65_UTLA_rank, by = c("AreaCode", "AreaName")) %>% 
  select(AreaName,AreaCode,`Female inequality in life expectancy at 65`,`Female inequality in life expectancy at 65 lci`,`Female inequality in life expectancy at 65 uci`,`Female inequality in life expectancy at 65 rank`, `Male inequality in life expectancy at 65`,`Male inequality in life expectancy at 65 lci`,`Male inequality in life expectancy at 65 uci`,`Male inequality in life expectancy at 65 rank`)

rm(Slope_inequalities_at_65_UTLA_val, Slope_inequalities_at_65_UTLA_lci, Slope_inequalities_at_65_UTLA_uci, Slope_inequalities_at_65_UTLA_rank)

Slope_inequalities_UTLA <- Slope_inequalities_at_birth_UTLA %>% 
  left_join(Slope_inequalities_at_65_UTLA, by = c("AreaCode", "AreaName")) %>% 
  left_join(Dep, by = c("AreaCode", "AreaName"))

rm(Slope_inequalities_at_birth_UTLA, Slope_inequalities_at_65_UTLA)

Selected_Slope_inequalities_ts <- Slope_inequalities %>% 
  filter(AreaName %in% Areas_to_include)

# ONS highlight those areas where the life expectancy is lower than the state pension age of 65

# Also do data frame for deprivation

# build excel file ####

wb <- loadWorkbook(paste0(github_repo_dir,"/Projections Build.xlsx"))

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
cells <- createCell(rows, colIndex = 1:6)

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
addDataFrame(unique(Area_population_df$Source), sheet, 
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

setCellValue(cells[[1,5]], "Counties_UA_selected")
setCellStyle(cells[[1,5]], cs_left)
addDataFrame(unique(Counties_UA_selected$Area_Name), sheet, 
             startRow = 2, 
             startColumn = 5, 
             col.names = FALSE,
             row.names = FALSE)
setColumnWidth(sheet, colIndex = 5, colWidth = max(nchar(Counties_UA_selected$Area_Name)))

setCellValue(cells[[1,6]], "Districts_UA_selected")
setCellStyle(cells[[1,6]], cs_left)
addDataFrame(unique(Districts_UA_selected$Area_Name), sheet, 
             startRow = 2, 
             startColumn = 6, 
             col.names = FALSE,
             row.names = FALSE)
setColumnWidth(sheet, colIndex = 6, colWidth = max(nchar(Districts_UA_selected$Area_Name)))

createRange("Area_list", cells[[2,1]], cells[[length(unique(Area_population_df$Area_Name))+1,1]])
createRange("Source_list", cells[[2,2]], cells[[length(unique(Area_population_df$Source))+1,2]])
createRange("Districts_UA_selected", cells[[2,5]], cells[[length(unique(Districts_UA_selected$Area_Name))+1,5]])
createRange("Counties_UA_selected", cells[[2,6]], cells[[length(unique(Counties_UA_selected$Area_Name))+1,6]])

wb$setActiveSheet(0L)
wb$setSheetHidden(17L, 1L) # This assumes List is sheet number 17.

removeSheet(wb, sheetName = "Raw Data")
sheet <- createSheet(wb, "Raw Data")
rows <- createRow(sheet, rowIndex = 1:nrow(Area_population_df))
cells <- createCell(rows, colIndex = 1:11)

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
setCellValue(cells[[1,10]], "Source")
setCellStyle(cells[[1,10]], cs_left)
setCellValue(cells[[1,11]], "All_age_population")
setCellStyle(cells[[1,11]], cs_left)


addDataFrame(as.data.frame(Area_population_df), sheet, 
             startRow = 2, 
             startColumn = 1, 
             col.names = FALSE,
             row.names = FALSE)

# width of collumn
autoSizeColumn(sheet, colIndex = 1:11)

wb$setSheetHidden(17L, 1L) # This assumes Raw Data is sheet number 17 (which it should be now, as index starts at 0)

rm(Area_population_df)

removeSheet(wb, sheetName = "OADR Data")
sheet <- createSheet(wb, "OADR Data")
rows <- createRow(sheet, rowIndex = 1:nrow(OADR))
cells <- createCell(rows, colIndex = 1:9)

setCellValue(cells[[1,1]], "Area_name")
setCellStyle(cells[[1,1]], cs_left)
setCellValue(cells[[1,2]], "Area_code")
setCellStyle(cells[[1,2]], cs_left)
setCellValue(cells[[1,3]], "Area_type")
setCellStyle(cells[[1,3]], cs_left)
setCellValue(cells[[1,4]], "Sex")
setCellStyle(cells[[1,4]], cs_left)
setCellValue(cells[[1,5]], "Data_type")
setCellStyle(cells[[1,5]], cs_left)
setCellValue(cells[[1,6]], "Year")
setCellStyle(cells[[1,6]], cs_left)
setCellValue(cells[[1,7]], "Old_age")
setCellStyle(cells[[1,7]], cs_left)
setCellValue(cells[[1,8]], "Working_age")
setCellStyle(cells[[1,8]], cs_left)
setCellValue(cells[[1,9]], "Ratio_per_1000")
setCellStyle(cells[[1,9]], cs_left)

addDataFrame(as.data.frame(OADR), sheet, 
             startRow = 2, 
             startColumn = 1, 
             col.names = FALSE,
             row.names = FALSE)

rm(OADR)

wb$setSheetHidden(17L, 1L) # This assumes OADR Data is sheet number 10 (which it should be now, as index starts at 0)

# width of collumn
autoSizeColumn(sheet, colIndex = 1:9)

removeSheet(wb, sheetName = "Old age data")
sheet <- createSheet(wb, "Old age data")
rows <- createRow(sheet, rowIndex = 1:nrow(Older_age_broad_projections))
cells <- createCell(rows, colIndex = 1:8)

setCellValue(cells[[1,1]], "Area_name")
setCellStyle(cells[[1,1]], cs_left)
setCellValue(cells[[1,2]], "Area_code")
setCellStyle(cells[[1,2]], cs_left)
setCellValue(cells[[1,3]], "Area_type")
setCellStyle(cells[[1,3]], cs_left)
setCellValue(cells[[1,4]], "Age_group")
setCellStyle(cells[[1,4]], cs_left)
setCellValue(cells[[1,5]], "Sex")
setCellStyle(cells[[1,5]], cs_left)
setCellValue(cells[[1,6]], "Data_type")
setCellStyle(cells[[1,6]], cs_left)
setCellValue(cells[[1,7]], "Year")
setCellStyle(cells[[1,7]], cs_left)
setCellValue(cells[[1,8]], "Population")
setCellStyle(cells[[1,8]], cs_left)

addDataFrame(as.data.frame(Older_age_broad_projections), sheet, 
             startRow = 2, 
             startColumn = 1, 
             col.names = FALSE,
             row.names = FALSE)

rm(Older_age_broad_projections)

wb$setSheetHidden(17L, 1L) # This assumes OADR Data is sheet number 10 (which it should be now, as index starts at 0)

# width of collumn
autoSizeColumn(sheet, colIndex = 1:8)

#start here add the wider ons sheets ####
removeSheet(wb, sheetName = "ONS LE data")
sheet <- createSheet(wb, "ONS LE data")
rows <- createRow(sheet, rowIndex = 1:max(nrow(LE_LTLA), nrow(LE_UTLA),nrow(HLE_at_birth_UTLA),nrow(Selected_Slope_inequalities_ts)))
cells <- createCell(rows, colIndex = 1:95)

setCellValue(cells[[1,1]], "Area_Name")
setCellStyle(cells[[1,1]], cs_left)
setCellValue(cells[[1,2]], "Area_Code")
setCellStyle(cells[[1,2]], cs_left)
setCellValue(cells[[1,3]], "Female_LE_at_birth")
setCellStyle(cells[[1,3]], cs_left)
setCellValue(cells[[1,4]], "Female_LE_at_birth_LCI")
setCellStyle(cells[[1,4]], cs_left)
setCellValue(cells[[1,5]], "Female_LE_at_birth_UCI")
setCellStyle(cells[[1,5]], cs_left)
setCellValue(cells[[1,6]], "Female_LE_at_birth_rank")
setCellStyle(cells[[1,6]], cs_left)
setCellValue(cells[[1,7]], "Male_LE_at_birth")
setCellStyle(cells[[1,7]], cs_left)
setCellValue(cells[[1,8]], "Male_LE_at_birth_LCI")
setCellStyle(cells[[1,8]], cs_left)
setCellValue(cells[[1,9]], "Male_LE_at_birth_UCI")
setCellStyle(cells[[1,9]], cs_left)
setCellValue(cells[[1,10]], "Male_LE_at_birth_rank")
setCellStyle(cells[[1,10]], cs_left)
setCellValue(cells[[1,11]], "Female_LE_at_65")
setCellStyle(cells[[1,11]], cs_left)
setCellValue(cells[[1,12]], "Female_LE_at_65_LCI")
setCellStyle(cells[[1,12]], cs_left)
setCellValue(cells[[1,13]], "Female_LE_at_65_UCI")
setCellStyle(cells[[1,13]], cs_left)
setCellValue(cells[[1,14]], "Female_LE_at_65_rank")
setCellStyle(cells[[1,14]], cs_left)
setCellValue(cells[[1,15]], "Male_LE_at_65")
setCellStyle(cells[[1,15]], cs_left)
setCellValue(cells[[1,16]], "Male_LE_at_65_LCI")
setCellStyle(cells[[1,16]], cs_left)
setCellValue(cells[[1,17]], "Male_LE_at_65_UCI")
setCellStyle(cells[[1,17]], cs_left)
setCellValue(cells[[1,18]], "Male_LE_at_65_rank")
setCellStyle(cells[[1,18]], cs_left)
setCellValue(cells[[1,19]], "Deprivation_score")
setCellStyle(cells[[1,19]], cs_left)
setCellValue(cells[[1,20]], "Proportion_in_deprived_areas")
setCellStyle(cells[[1,20]], cs_left)
setCellValue(cells[[1,21]], "Proportion_in_deprived_areas_LCI")
setCellStyle(cells[[1,21]], cs_left)
setCellValue(cells[[1,22]], "Proportion_in_deprived_areas_UCI")
setCellStyle(cells[[1,22]], cs_left)
setCellValue(cells[[1,23]], "N_in_deprived_areas")
setCellStyle(cells[[1,23]], cs_left)

addDataFrame(as.data.frame(LE_LTLA), sheet, 
             startRow = 2, 
             startColumn = 1, 
             col.names = FALSE,
             row.names = FALSE)

rm(LE_LTLA)

setCellValue(cells[[1,25]], "Area_Name")
setCellStyle(cells[[1,25]], cs_left)
setCellValue(cells[[1,26]], "Area_Code")
setCellStyle(cells[[1,26]], cs_left)
setCellValue(cells[[1,27]], "Female_LE_at_birth")
setCellStyle(cells[[1,27]], cs_left)
setCellValue(cells[[1,28]], "Female_LE_at_birth_LCI")
setCellStyle(cells[[1,28]], cs_left)
setCellValue(cells[[1,29]], "Female_LE_at_birth_UCI")
setCellStyle(cells[[1,29]], cs_left)
setCellValue(cells[[1,30]], "Female_LE_at_birth_rank")
setCellStyle(cells[[1,30]], cs_left)
setCellValue(cells[[1,31]], "Male_LE_at_birth")
setCellStyle(cells[[1,31]], cs_left)
setCellValue(cells[[1,32]], "Male_LE_at_birth_LCI")
setCellStyle(cells[[1,32]], cs_left)
setCellValue(cells[[1,33]], "Male_LE_at_birth_UCI")
setCellStyle(cells[[1,33]], cs_left)
setCellValue(cells[[1,34]], "Male_LE_at_birth_rank")
setCellStyle(cells[[1,34]], cs_left)
setCellValue(cells[[1,35]], "Female_LE_at_65")
setCellStyle(cells[[1,35]], cs_left)
setCellValue(cells[[1,36]], "Female_LE_at_65_LCI")
setCellStyle(cells[[1,36]], cs_left)
setCellValue(cells[[1,37]], "Female_LE_at_65_UCI")
setCellStyle(cells[[1,37]], cs_left)
setCellValue(cells[[1,38]], "Female_LE_at_65_rank")
setCellStyle(cells[[1,38]], cs_left)
setCellValue(cells[[1,39]], "Male_LE_at_65")
setCellStyle(cells[[1,39]], cs_left)
setCellValue(cells[[1,40]], "Male_LE_at_65_LCI")
setCellStyle(cells[[1,40]], cs_left)
setCellValue(cells[[1,41]], "Male_LE_at_65_UCI")
setCellStyle(cells[[1,41]], cs_left)
setCellValue(cells[[1,42]], "Male_LE_at_65_rank")
setCellStyle(cells[[1,42]], cs_left)
setCellValue(cells[[1,43]], "Deprivation_score")
setCellStyle(cells[[1,43]], cs_left)
setCellValue(cells[[1,44]], "Proportion_in_deprived_areas")
setCellStyle(cells[[1,44]], cs_left)
setCellValue(cells[[1,45]], "Proportion_in_deprived_areas_LCI")
setCellStyle(cells[[1,45]], cs_left)
setCellValue(cells[[1,46]], "Proportion_in_deprived_areas_UCI")
setCellStyle(cells[[1,46]], cs_left)
setCellValue(cells[[1,47]], "N_in_deprived_areas")
setCellStyle(cells[[1,47]], cs_left)

addDataFrame(as.data.frame(LE_UTLA), sheet, 
             startRow = 5, 
             startColumn = 25, 
             col.names = FALSE,
             row.names = FALSE)

rm(LE_UTLA)
setCellValue(cells[[1,49]], "Area_Name")
setCellStyle(cells[[1,49]], cs_left)
setCellValue(cells[[1,50]], "Area_Code")
setCellStyle(cells[[1,50]], cs_left)
setCellValue(cells[[1,51]], "Female_HLE_at_birth")
setCellStyle(cells[[1,51]], cs_left)
setCellValue(cells[[1,52]], "Female_HLE_at_birth_LCI")
setCellStyle(cells[[1,52]], cs_left)
setCellValue(cells[[1,53]], "Female_HLE_at_birth_UCI")
setCellStyle(cells[[1,53]], cs_left)
setCellValue(cells[[1,54]], "Female_HLE_at_birth_rank")
setCellStyle(cells[[1,54]], cs_left)
setCellValue(cells[[1,55]], "Male_HLE_at_birth")
setCellStyle(cells[[1,55]], cs_left)
setCellValue(cells[[1,56]], "Male_HLE_at_birth_LCI")
setCellStyle(cells[[1,56]], cs_left)
setCellValue(cells[[1,57]], "Male_HLE_at_birth_UCI")
setCellStyle(cells[[1,57]], cs_left)
setCellValue(cells[[1,58]], "Male_HLE_at_birth_rank")
setCellStyle(cells[[1,58]], cs_left)
setCellValue(cells[[1,59]], "Female_HLE_at_65")
setCellStyle(cells[[1,59]], cs_left)
setCellValue(cells[[1,60]], "Female_HLE_at_65_LCI")
setCellStyle(cells[[1,60]], cs_left)
setCellValue(cells[[1,61]], "Female_HLE_at_65_UCI")
setCellStyle(cells[[1,61]], cs_left)
setCellValue(cells[[1,62]], "Female_HLE_at_65_rank")
setCellStyle(cells[[1,62]], cs_left)
setCellValue(cells[[1,63]], "Male_HLE_at_65")
setCellStyle(cells[[1,63]], cs_left)
setCellValue(cells[[1,64]], "Male_HLE_at_65_LCI")
setCellStyle(cells[[1,64]], cs_left)
setCellValue(cells[[1,65]], "Male_HLE_at_65_UCI")
setCellStyle(cells[[1,65]], cs_left)
setCellValue(cells[[1,66]], "Male_HLE_at_65_rank")
setCellStyle(cells[[1,66]], cs_left)
setCellValue(cells[[1,67]], "Deprivation_score")
setCellStyle(cells[[1,67]], cs_left)
setCellValue(cells[[1,68]], "Proportion_in_deprived_areas")
setCellStyle(cells[[1,68]], cs_left)
setCellValue(cells[[1,69]], "Proportion_in_deprived_areas_LCI")
setCellStyle(cells[[1,69]], cs_left)
setCellValue(cells[[1,70]], "Proportion_in_deprived_areas_UCI")
setCellStyle(cells[[1,70]], cs_left)
setCellValue(cells[[1,71]], "N_in_deprived_areas")
setCellStyle(cells[[1,71]], cs_left)

addDataFrame(as.data.frame(HLE_at_birth_UTLA), sheet, 
             startRow = 2, 
             startColumn = 49, 
             col.names = FALSE,
             row.names = FALSE)

rm(HLE_at_birth_UTLA)

setCellValue(cells[[1,73]], "Area_Name")
setCellStyle(cells[[1,73]], cs_left)
setCellValue(cells[[1,74]], "Area_Code")
setCellStyle(cells[[1,74]], cs_left)
setCellValue(cells[[1,75]], "Female_HLE_at_birth")
setCellStyle(cells[[1,75]], cs_left)
setCellValue(cells[[1,76]], "Female_HLE_at_birth_LCI")
setCellStyle(cells[[1,76]], cs_left)
setCellValue(cells[[1,77]], "Female_HLE_at_birth_UCI")
setCellStyle(cells[[1,77]], cs_left)
setCellValue(cells[[1,78]], "Female_HLE_at_birth_rank")
setCellStyle(cells[[1,78]], cs_left)
setCellValue(cells[[1,79]], "Male_HLE_at_birth")
setCellStyle(cells[[1,79]], cs_left)
setCellValue(cells[[1,80]], "Male_HLE_at_birth_LCI")
setCellStyle(cells[[1,80]], cs_left)
setCellValue(cells[[1,81]], "Male_HLE_at_birth_UCI")
setCellStyle(cells[[1,81]], cs_left)
setCellValue(cells[[1,82]], "Male_HLE_at_birth_rank")
setCellStyle(cells[[1,82]], cs_left)
setCellValue(cells[[1,83]], "Female_HLE_at_65")
setCellStyle(cells[[1,83]], cs_left)
setCellValue(cells[[1,84]], "Female_HLE_at_65_LCI")
setCellStyle(cells[[1,84]], cs_left)
setCellValue(cells[[1,85]], "Female_HLE_at_65_UCI")
setCellStyle(cells[[1,85]], cs_left)
setCellValue(cells[[1,86]], "Female_HLE_at_65_rank")
setCellStyle(cells[[1,86]], cs_left)
setCellValue(cells[[1,87]], "Male_HLE_at_65")
setCellStyle(cells[[1,87]], cs_left)
setCellValue(cells[[1,88]], "Male_HLE_at_65_LCI")
setCellStyle(cells[[1,88]], cs_left)
setCellValue(cells[[1,89]], "Male_HLE_at_65_UCI")
setCellStyle(cells[[1,89]], cs_left)
setCellValue(cells[[1,90]], "Male_HLE_at_65_rank")
setCellStyle(cells[[1,90]], cs_left)
setCellValue(cells[[1,91]], "Deprivation_score")
setCellStyle(cells[[1,91]], cs_left)
setCellValue(cells[[1,92]], "Proportion_in_deprived_areas")
setCellStyle(cells[[1,92]], cs_left)
setCellValue(cells[[1,93]], "Proportion_in_deprived_areas_LCI")
setCellStyle(cells[[1,93]], cs_left)
setCellValue(cells[[1,94]], "Proportion_in_deprived_areas_UCI")
setCellStyle(cells[[1,94]], cs_left)
setCellValue(cells[[1,95]], "N_in_deprived_areas")
setCellStyle(cells[[1,95]], cs_left)

addDataFrame(as.data.frame(Slope_inequalities_UTLA), sheet, 
             startRow = 2, 
             startColumn = 73, 
             col.names = FALSE,
             row.names = FALSE)

rm(Slope_inequalities_UTLA)

wb$setSheetHidden(17L, 1L) # This assumes ONS LE Data is sheet number 10 (which it should be now, as index starts at 0)

# width of collumn
autoSizeColumn(sheet, colIndex = 1:95)

removeSheet(wb, sheetName = "Selected LE data")
sheet <- createSheet(wb, "Selected LE data")
rows <- createRow(sheet, rowIndex = 1:max(nrow(Selected_LE_ONS_ts), nrow(selected_HLE_ONS_UTLA_ts), nrow(Selected_Slope_inequalities_ts)))
cells <- createCell(rows, colIndex = 1:34)

setCellValue(cells[[1,1]], "Indicator")
setCellStyle(cells[[1,1]], cs_left)
setCellValue(cells[[1,2]], "Area_Codee")
setCellStyle(cells[[1,2]], cs_left)
setCellValue(cells[[1,3]], "Area_Name")
setCellStyle(cells[[1,3]], cs_left)
setCellValue(cells[[1,4]], "AreaType")
setCellStyle(cells[[1,4]], cs_left)
setCellValue(cells[[1,5]], "Sex")
setCellStyle(cells[[1,5]], cs_left)
setCellValue(cells[[1,6]], "Age")
setCellStyle(cells[[1,6]], cs_left)
setCellValue(cells[[1,7]], "Time_Period")
setCellStyle(cells[[1,7]], cs_left)
setCellValue(cells[[1,8]], "Life Expectancy at birth")
setCellStyle(cells[[1,8]], cs_left)
setCellValue(cells[[1,8]], "Life Expectancy at birth LCI")
setCellStyle(cells[[1,8]], cs_left)
setCellValue(cells[[1,10]], "Life Expectancy at birth UCI")
setCellStyle(cells[[1,10]], cs_left)
setCellValue(cells[[1,11]], "Life Expectancy at birth rank")
setCellStyle(cells[[1,11]], cs_left)

addDataFrame(as.data.frame(Selected_LE_ONS_ts), sheet,
             startRow = 2,
             startColumn = 1,
             col.names = FALSE,
             row.names = FALSE)

setCellValue(cells[[1,13]], "Indicator")
setCellStyle(cells[[1,13]], cs_left)
setCellValue(cells[[1,14]], "Area_Codee")
setCellStyle(cells[[1,14]], cs_left)
setCellValue(cells[[1,15]], "Area_Name")
setCellStyle(cells[[1,15]], cs_left)
setCellValue(cells[[1,16]], "AreaType")
setCellStyle(cells[[1,16]], cs_left)
setCellValue(cells[[1,17]], "Sex")
setCellStyle(cells[[1,17]], cs_left)
setCellValue(cells[[1,18]], "Age")
setCellStyle(cells[[1,18]], cs_left)
setCellValue(cells[[1,19]], "Time_Period")
setCellStyle(cells[[1,19]], cs_left)
setCellValue(cells[[1,20]], "Healthy Life Expectancy at birth")
setCellStyle(cells[[1,20]], cs_left)
setCellValue(cells[[1,21]], "Healthy Life Expectancy at birth LCI")
setCellStyle(cells[[1,21]], cs_left)
setCellValue(cells[[1,22]], "Healthy Life Expectancy at birth UCI")
setCellStyle(cells[[1,22]], cs_left)
setCellValue(cells[[1,23]], "Healthy Life Expectancy at birth rank")
setCellStyle(cells[[1,23]], cs_left)

addDataFrame(as.data.frame(selected_HLE_ONS_UTLA_ts), sheet,
             startRow = 2,
             startColumn = 13,
             col.names = FALSE,
             row.names = FALSE)


setCellValue(cells[[1,25]], "Indicator")
setCellStyle(cells[[1,25]], cs_left)
setCellValue(cells[[1,26]], "Area_Codee")
setCellStyle(cells[[1,26]], cs_left)
setCellValue(cells[[1,27]], "Area_Name")
setCellStyle(cells[[1,27]], cs_left)
setCellValue(cells[[1,28]], "AreaType")
setCellStyle(cells[[1,28]], cs_left)
setCellValue(cells[[1,29]], "Sex")
setCellStyle(cells[[1,29]], cs_left)
setCellValue(cells[[1,30]], "Age")
setCellStyle(cells[[1,30]], cs_left)
setCellValue(cells[[1,31]], "Time_Period")
setCellStyle(cells[[1,31]], cs_left)
setCellValue(cells[[1,32]], "Slope inequality in Life Expectancy at birth")
setCellStyle(cells[[1,32]], cs_left)
setCellValue(cells[[1,33]], "Slope inequality in Life Expectancy at birth LCI")
setCellStyle(cells[[1,33]], cs_left)
setCellValue(cells[[1,34]], "Slope inequality in Life Expectancy at birth UCI")
setCellStyle(cells[[1,34]], cs_left)

addDataFrame(as.data.frame(Selected_Slope_inequalities_ts), sheet,
             startRow = 2,
             startColumn = 25,
             col.names = FALSE,
             row.names = FALSE)

wb$setSheetHidden(17L, 1L) # This assumes Raw Data is sheet number 17 (which it should be now, as index starts at 0)

autoSizeColumn(sheet, colIndex = 1:34)


# removeSheet(wb, sheetName = "GBD LE data")
# sheet <- createSheet(wb, "GBD LE data")
# rows <- createRow(sheet, rowIndex = 1:nrow(gbd_LE_tables))
# cells <- createCell(rows, colIndex = 1:8)
# 
# setCellValue(cells[[1,1]], "Indicator")
# setCellStyle(cells[[1,1]], cs_left)
# setCellValue(cells[[1,2]], "Area_Name")
# setCellStyle(cells[[1,2]], cs_left)
# setCellValue(cells[[1,3]], "Sex")
# setCellStyle(cells[[1,3]], cs_left)
# setCellValue(cells[[1,4]], "Age")
# setCellStyle(cells[[1,4]], cs_left)
# setCellValue(cells[[1,5]], "Year")
# setCellStyle(cells[[1,5]], cs_left)
# setCellValue(cells[[1,6]], "Life_years_left")
# setCellStyle(cells[[1,6]], cs_left)
# setCellValue(cells[[1,7]], "Life_years_left_LCI")
# setCellStyle(cells[[1,7]], cs_left)
# setCellValue(cells[[1,8]], "Life_years_left_UCI")
# setCellStyle(cells[[1,8]], cs_left)
# 
# addDataFrame(as.data.frame(gbd_LE_tables), sheet,
#              startRow = 2,
#              startColumn = 1,
#              col.names = FALSE,
#              row.names = FALSE)
# 
# wb$setSheetHidden(17L, 1L) # This assumes Raw Data is sheet number 17 (which it should be now, as index starts at 0)

# autoSizeColumn(sheet, colIndex = 1:8)

removeSheet(wb, sheetName = "Selected gbd data")
sheet <- createSheet(wb, "Selected gbd data")
rows <- createRow(sheet, rowIndex = 1:nrow(selected_areas_gbd_LE_tables))
cells <- createCell(rows, colIndex = 1:8)

setCellValue(cells[[1,1]], "Indicator")
setCellStyle(cells[[1,1]], cs_left)
setCellValue(cells[[1,2]], "Area_Name")
setCellStyle(cells[[1,2]], cs_left)
setCellValue(cells[[1,3]], "Sex")
setCellStyle(cells[[1,3]], cs_left)
setCellValue(cells[[1,4]], "Age")
setCellStyle(cells[[1,4]], cs_left)
setCellValue(cells[[1,5]], "Year")
setCellStyle(cells[[1,5]], cs_left)
setCellValue(cells[[1,6]], "Life_years_left")
setCellStyle(cells[[1,6]], cs_left)
setCellValue(cells[[1,7]], "Life_years_left_LCI")
setCellStyle(cells[[1,7]], cs_left)
setCellValue(cells[[1,8]], "Life_years_left_UCI")
setCellStyle(cells[[1,8]], cs_left)

addDataFrame(as.data.frame(selected_areas_gbd_LE_tables), sheet,
             startRow = 2,
             startColumn = 1,
             col.names = FALSE,
             row.names = FALSE)

# width of collumn
autoSizeColumn(sheet, colIndex = 1:8)

wb$setSheetHidden(17L, 1L) # This assumes Raw Data is sheet number 17 (which it should be now, as index starts at 0)

saveWorkbook(wb, file = "./Projecting-Health/STP Profile 2019 Update v1.xlsx")

# There are still problems.

# We can try openxlsx package as this aparently does not rely on java.

# For now though, we just need to crack on!

write.csv(selected_areas_gbd_LE_tables, "./Projecting-Health/selected_areas_gbd_LE_tables.csv", row.names = FALSE)

