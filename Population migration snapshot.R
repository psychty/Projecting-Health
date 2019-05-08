
library(easypackages)
# This uses the easypackages package to load several libraries at once. Note: it should only be used when you are confident that all packages are installed as it will be more difficult to spot load errors compared to loading each one individually.
libraries(c("readxl", "readr", "plyr", "dplyr", "ggplot2", "png", "tidyverse", "reshape2", "scales", "viridis", "rgdal", "officer", "flextable", "tmaptools", "lemon", "fingertipsR", "PHEindicatormethods"))

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

Areas_to_include <- c("Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "NHS Coastal West Sussex CCG","West Sussex", "NHS Crawley CCG", "NHS Horsham and Mid Sussex CCG")

if(!(file.exists("~/Projecting-Health/Area_population_df.csv"))){
  print("Area_population_df is not available, it will be built using the 'Areas_to_include' object")
  source("~/Projecting-Health/Get data - mye and projections.R")
  
  if(file.exists("~/Projecting-Health/Area_population_df.csv")){
    print("Area_population_df is now available.")
  }
  
}

if(exists("Areas_to_include") & file.exists("~/Projecting-Health/Area_population_df.csv")){
  print("Both objects are available")
Area_population_df <- read_csv("~/Projecting-Health/Area_population_df.csv", col_types = cols(Area_Name = col_character(),Area_Code = col_character(),Area_Type = col_character(),  Sex = col_character(),Age_group = col_character(),Age_band_type = col_character(),Year = col_double(),Population = col_double(),Data_type = col_character()))  

Areas <- read_csv("~/Projecting-Health/Area_lookup_table.csv", col_types = cols(LTLA17CD = col_character(),LTLA17NM = col_character(),UTLA17CD = col_character(),UTLA17NM = col_character(),FID = col_double()))
Lookup <- read_csv("~/Projecting-Health/Area_types_table.csv", col_types = cols(Area_Code = col_character(),Area_Name = col_character(),Area_Type = col_character()))

if(length(setdiff(Areas_to_include, Area_population_df$Area_Name))>0){
  print("There are some areas chosen that are not in the Area_population_df. The 'Get data - mye and projections' script will now run and will overwrite the Area_population_df.")
  source("~/Projecting-Health/Get data - mye and projections.R")
}

if(length(setdiff(Areas_to_include, Area_population_df$Area_Name))>0){
  print("There are still some areas chosen that are not in the Area_population_df. Check the Areas_to_include object.")
}

if(length(setdiff(Areas_to_include, Area_population_df$Area_Name))==0){
  print("The Area_population_df matches the Areas_to_include list.")
}
}

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
    panel.grid.major.y = element_line(colour = "#E2E2E3", linetype = "longdash"), 
    panel.grid.minor.x = element_blank(), 
    panel.grid.major.x = element_blank(), 
    panel.grid.minor.y = element_blank(), 
    strip.text = element_text(colour = "white"), 
    strip.background = element_rect(fill = "#327d9c"), 
    axis.ticks = element_line(colour = "#327d9c") 
  ) 
}


# Components of population change ####

if(!(file.exists("~/Projecting-Health/Components-of-change"))){
  dir.create("~/Projecting-Health/Components-of-change/")}

download.file("https://www.ons.gov.uk/file?uri=/peoplepopulationandcommunity/populationandmigration/populationestimates/datasets/populationestimatesforukenglandandwalesscotlandandnorthernireland/mid2001tomid2017detailedtimeseries/ukdetailedtimeseries2001to2017.zip", "~/Projecting-Health/Components-of-change/Components_of_change_01_17.zip", mode = "wb")
unzip("~/Projecting-Health/Components-of-change/Components_of_change_01_17.zip", exdir = "~/Projecting-Health/Components-of-change")
file.remove("~/Projecting-Health/Components-of-change/Components_of_change_01_17.zip")


latest_snapshot <- read_excel("~/Projecting-Health/Migration_snapshot.xlsx", sheet = "MYE3",col_names = TRUE, skip = 4)
latest_snapshot$`Internal Migration Inflow` <- as.numeric(latest_snapshot$`Internal Migration Inflow`)
latest_snapshot$`Internal Migration Outflow` <- as.numeric(latest_snapshot$`Internal Migration Outflow`)

WSX_snapshot <- subset(latest_snapshot, Name == "West Sussex")

WSX_snapshot$`Estimated Population mid-2016` - WSX_snapshot$`Estimated Population mid-2015`

(WSX_snapshot$`Estimated Population mid-2016` - WSX_snapshot$`Estimated Population mid-2015`) / WSX_snapshot$`Estimated Population mid-2015` * 100

# MYE5: Population estimates: Population density for the local authorities in the UK, mid-2001 to mid-2016 ####
pop_density_ts <- read_excel("./Mid2016.xls", sheet = "MYE5",col_names = TRUE, skip = 4)

pop_density_ts <- pop_density_ts[c("Name","Area (sq km)","2016 people per sq. km","2015 people per sq. km" , "2014 people per sq. km"   , "2013 people per sq. km"  , "2012 people per sq. km", "2011 people per sq. km"   , "2010 people per sq. km"      , "2009 people per sq. km"     ,"2008 people per sq. km"       , "2007 people per sq. km"       ,"2006 people per sq. km"   ,"2005 people per sq. km"       , "2004 people per sq. km"       ,"2003 people per sq. km"       , "2002 people per sq. km"      , "2001 people per sq. km"       )]

# WSX_pop_density <- subset(pop_density_ts, Name == "West Sussex")

pop_density_long <- gather(data = pop_density_ts, key = "Year", value = "Population per sq km", 3:18) # This puts the data into long format (so males and females population counts are in the same field)

pop_density_long$Year <- gsub(" people per sq. km", "", pop_density_long$Year)

areas_for_chart <- subset(pop_density_long, Name %in% West_Sussex_areas)

areas_for_chart <- merge(areas_for_chart, wsx_dist, by = "Name")

pdf("Population_density_wsx_districts_2016.pdf",  width = 11.7, height = 8.3) # a4 landscape
ggplot(data = areas_for_chart, aes(x = Year, y = `Population per sq km`, group = Name, colour = Name)) +
  geom_line() +
  geom_point() +
  # ph_theme() + 
  ggtitle("Population density (Residents per square km);\nWest Sussex districts; Mid 2016 estimates") +
  labs(caption = "Source: Office for National Statistics licensed under the Open Government Licence. © Crown copyright 2017") +
  scale_y_continuous(breaks = seq(0, 3500, 250), limits = c(0, 3500), labels = comma) + 
  xlab("Year") + ylab("Population per square km") +
  ph_theme()
dev.off()

subset(areas_for_chart, Year == "2016")

# MYE6: Median age of population for local authorities in the UK, mid-2001 to mid-2016 ####
median_age_ts <- read_excel("./Mid2016.xls", sheet = "MYE6",col_names = TRUE, skip = 4)

WSX_median_age <- subset(median_age_ts, Name %in% West_Sussex_areas)
WSX_median_age_long <- gather(data = WSX_median_age, key = "Year", value = "Median Age", 3:18) # This puts the data into long format (so males and females population counts are in the same field)

WSX_median_age_long$Year <- gsub("Mid-", "", WSX_median_age_long$Year)

pdf("Median_age_wsx_districts_2016.pdf",  width = 11.7, height = 8.3) # a4 landscape
ggplot(data = WSX_median_age_long, aes(x = Year, y = `Median Age`, group = Name, colour = Name)) +
  geom_line() +
  geom_point() +
  # ph_theme() + 
  ggtitle("Median age; West Sussex Districts and Boroughs; Mid 2016 estimates") +
  labs(caption = "Note: Y axis does not start at zero.\nSource: Office for National Statistics licensed under the Open Government Licence. © Crown copyright 2017") +
  scale_y_continuous(breaks = seq(30, 50, 5), limits = c(30, 50), labels = comma) + 
  xlab("Year") + ylab("Median age") +
  ph_theme()
dev.off()

WSX_median_age[c("Name", "Mid-2016")]

# District net migration by age ####

#download.file("https://www.ons.gov.uk/file?uri=/peoplepopulationandcommunity/populationandmigration/migrationwithintheuk/datasets/internalmigrationmovesbylocalauthoritiesandregionsinenglandandwalesby5yearagegroupandsex/yearendingjune2016/laandregionsby5yearagesgroupings.xls", "LA_five_year_age_migration.xls", mode = "wb")

migration <- read_excel("LA_five_year_age_migration.xls", sheet = "IM2016-T5", skip = 6)

colnames(migration) <- c("LA_code", "LA_name", "Age", "Inflow_persons", "Outflow_persons", "Net_moves_persons", "Inflow_male", "Outflow_male", "Net_moves_male", "Inflow_female", "Outflow_female", "Net_moves_female")

migration <- subset(migration, LA_name %in% West_Sussex_areas)
migration$Age <- factor(migration$Age, levels = c("0-4",  "5-9", "10-14", "15-19", "20-24", "25-29", "30-34", "35-39", "40-44", "45-49", "50-54", "55-59", "60-64", "65-69", "70-74", "75-79", "80-84", "85-89", "90+"))

library(dplyr)

migration_WSx <- migration %>%
  group_by(Age) %>%
  summarise(LA_code = NA, LA_name = "West Sussex", Inflow_persons = sum(Inflow_persons), Outflow_persons = sum(Outflow_persons), Net_moves_persons = sum(Net_moves_persons), Inflow_male = sum(Inflow_male), Outflow_male = sum(Outflow_male), Net_moves_male = sum(Net_moves_male), Inflow_female = sum(Inflow_female), Outflow_female = sum(Outflow_female), Net_moves_female = sum(Net_moves_female))

pdf("internal_migration_wsx_districts_2016.pdf",  width = 11.7, height = 8.3) # a4 landscape
ggplot(data = migration, aes(x = Age, y = Net_moves_persons)) +
  geom_bar(stat = "identity", width = 0.9) +
  ggtitle("Net internal migration flow; West Sussex districts by age group;\nYear ending June 2016") +
  scale_y_continuous(breaks = seq(-650,650,100), limits = c(-650,650)) +
  facet_wrap(~LA_name, nrow = 2) +
  xlab("Age group")+
  ylab("Number of residents (rounded to nearest 10)") +
  ph_theme() +
  labs(caption = "Note: A value above zero (0) indicates that more people entered the district than left.\nNumbers have been rounded to nearest 10") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.5), axis.text = element_text(colour = "#000000", size = 10)  )
dev.off()

pdf("internal_migration_wsx_2016.pdf",  width = 11.7, height = 8.3) # a4 landscape
ggplot(data = migration_WSx, aes(x = Age, y = Net_moves_persons)) +
  geom_bar(stat = "identity") +
  ggtitle("Net internal migration flow; West Sussex districts by age group;\nYear ending June 2016") +
  scale_y_continuous(breaks = seq(-5000,5000,500), limits = c(-2000,2000), labels = comma) +
  xlab("Age group")+
  ylab("Number of residents (rounded to nearest 10)") +
  ph_theme() +
  labs(caption = "Note: A value above zero (0) indicates than more people entered West Sussex than left.\nNumbers have been rounded to nearest 10, and as such may not add up") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.5))
dev.off()

# Internal migration flow ####
migration_flow_1 <- read.csv("./Detailed_Estimates_2016_Dataset_1.csv", header = TRUE)
migration_flow_2 <- read.csv("./Detailed_Estimates_2016_Dataset_2.csv", header = TRUE)
migration_flow <- rbind(migration_flow_1, migration_flow_2)
#rm(migration_flow_1);rm(migration_flow_2)

la_meta <- read.csv("https://opendata.arcgis.com/datasets/464be6191a434a91a5fa2f52c7433333_0.csv", header = TRUE)
la_meta$FID <- NULL
la_meta$LAD16CDO <- NULL
colnames(la_meta) <- c("LAD_code", "LAD_name")

wsxla_meta <- data.frame(LAD_code = c("E07000223","E07000224","E07000225","E07000226","E07000227","E07000228","E07000229"), LAD_name = c("Adur","Arun","Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing"))

migration_flow <- subset(migration_flow, OutLA %in% wsxla_meta$LAD_code | InLA %in% wsxla_meta$LAD_code)

# 40,386 rows of data
migration_flow <- merge(migration_flow, la_meta, by.x = "OutLA", by.y = "LAD_code", all.x = TRUE)

colnames(migration_flow) <- c("OutLA","InLA","Age","Sex", "Moves", "OutLA_name")

migration_flow <- merge(migration_flow, la_meta, by.x = "InLA", by.y = "LAD_code", keep.x = TRUE)

colnames(migration_flow) <- c("InLA","OutLA","Age","Sex", "Moves", "OutLA_name", "InLA_name")

migration_flow <- migration_flow[c("OutLA","InLA","Age","Sex", "Moves", "OutLA_name", "InLA_name")]

# Perhaps count the number of outs from Adur and see if it corresponds to other umbers

adur_mig_flow <- subset(migration_flow, OutLA_name == "Adur")

arun_mig_flow <- subset(migration_flow, OutLA_name == "Arun")
sum(arun_mig_flow$Moves)


#g####
# .	Age: Defined as age as at 30 June 2016, so in many cases will be one year older than age at actual move.

# .	Aggregating figures: LA-level inflows and outflows may be derived directly from the dataset. However, inflows and outflows for groups of LAs should not simply be added together as the totals will include moves between those LAs. Users will therefore need to take care to strip out those moves. However, the release includes ready-made tables of total moves at regional level. County level totals are included in the analysis tool accompanying ONS's annual population estimates.

# Need to account for moves to/from districts in WSx

# Could do map with lines

# components of change over time 2002 to 2016

components_change <- read.csv("./MYEB2_detailed_components_of_change_series_EW_(0216).csv", header = TRUE)

components_change <- subset(components_change, lad2014_name %in% c("Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "West Sussex"))
components_change$sex <- gsub("1","Male",components_change$sex)
components_change$sex <- gsub("2","Female",components_change$sex)

people_ts <- read.csv("./MYEB1_detailed_population_estimates_series_UK_(0116).csv", header = TRUE)
people_ts <- subset(people_ts, lad2014_name %in% c("Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "West Sussex"))
people_ts$sex <- gsub("1","Male",people_ts$sex)
people_ts$sex <- gsub("2","Female",people_ts$sex)

all_ages_summary_cchange <- read.csv("./MYEB3_summary_components_of_change_series_UK_(0216).csv", header = TRUE)
all_ages_summary_cchange <- subset(all_ages_summary_cchange, lad2014_name %in% c("Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "West Sussex"))

all_age_ts <- all_ages_summary_cchange[c("lad2014_name","population_2001", "population_2002", "population_2003", "population_2004", "population_2005", "population_2006", "population_2007", "population_2008", "population_2009", "population_2010", "population_2011", "population_2012", "population_2013", "population_2014", "population_2015", "population_2016")]

all_age_ts$population_2001 <- gsub(",", "", all_age_ts$population_2001);all_age_ts$population_2001 <- as.numeric(all_age_ts$population_2001)
all_age_ts$population_2002 <- gsub(",", "", all_age_ts$population_2002);all_age_ts$population_2002 <- as.numeric(all_age_ts$population_2002)
all_age_ts$population_2003 <- gsub(",", "", all_age_ts$population_2003);all_age_ts$population_2003 <- as.numeric(all_age_ts$population_2003)
all_age_ts$population_2004 <- gsub(",", "", all_age_ts$population_2004);all_age_ts$population_2004 <- as.numeric(all_age_ts$population_2004)
all_age_ts$population_2005 <- gsub(",", "", all_age_ts$population_2005);all_age_ts$population_2005 <- as.numeric(all_age_ts$population_2005)
all_age_ts$population_2006 <- gsub(",", "", all_age_ts$population_2006);all_age_ts$population_2006 <- as.numeric(all_age_ts$population_2006)
all_age_ts$population_2007 <- gsub(",", "", all_age_ts$population_2007);all_age_ts$population_2007 <- as.numeric(all_age_ts$population_2007)
all_age_ts$population_2008 <- gsub(",", "", all_age_ts$population_2008);all_age_ts$population_2008 <- as.numeric(all_age_ts$population_2008)
all_age_ts$population_2009 <- gsub(",", "", all_age_ts$population_2009);all_age_ts$population_2009 <- as.numeric(all_age_ts$population_2009)
all_age_ts$population_2010 <- gsub(",", "", all_age_ts$population_2010);all_age_ts$population_2010 <- as.numeric(all_age_ts$population_2010)
all_age_ts$population_2011 <- gsub(",", "", all_age_ts$population_2011);all_age_ts$population_2011 <- as.numeric(all_age_ts$population_2011)
all_age_ts$population_2012 <- gsub(",", "", all_age_ts$population_2012);all_age_ts$population_2012 <- as.numeric(all_age_ts$population_2012)
all_age_ts$population_2013 <- gsub(",", "", all_age_ts$population_2013);all_age_ts$population_2013 <- as.numeric(all_age_ts$population_2013)


all_age_ts$`2001` <- 0
all_age_ts$diff_2002 <- (all_age_ts$population_2002 - all_age_ts$population_2001)/all_age_ts$population_2001 
all_age_ts$diff_2003 <- (all_age_ts$population_2003 - all_age_ts$population_2002)/all_age_ts$population_2002 
all_age_ts$diff_2004 <- (all_age_ts$population_2004 - all_age_ts$population_2003)/all_age_ts$population_2003 
all_age_ts$diff_2005 <- (all_age_ts$population_2005 - all_age_ts$population_2004)/all_age_ts$population_2004
all_age_ts$diff_2006 <- (all_age_ts$population_2006 - all_age_ts$population_2005)/all_age_ts$population_2005 
all_age_ts$diff_2007 <- (all_age_ts$population_2007 - all_age_ts$population_2006)/all_age_ts$population_2006 
all_age_ts$diff_2008 <- (all_age_ts$population_2008 - all_age_ts$population_2007)/all_age_ts$population_2007 
all_age_ts$diff_2009 <- (all_age_ts$population_2009 - all_age_ts$population_2008)/all_age_ts$population_2008 
all_age_ts$diff_2010 <- (all_age_ts$population_2010 - all_age_ts$population_2009)/all_age_ts$population_2009 
all_age_ts$diff_2011 <- (all_age_ts$population_2011 - all_age_ts$population_2010)/all_age_ts$population_2010
all_age_ts$diff_2012 <- (all_age_ts$population_2012 - all_age_ts$population_2011)/all_age_ts$population_2011 
all_age_ts$diff_2013 <- (all_age_ts$population_2013 - all_age_ts$population_2012)/all_age_ts$population_2012
all_age_ts$diff_2014 <- (all_age_ts$population_2014 - all_age_ts$population_2013)/all_age_ts$population_2013 
all_age_ts$diff_2015 <- (all_age_ts$population_2015 - all_age_ts$population_2014)/all_age_ts$population_2014 
all_age_ts$diff_2016 <- (all_age_ts$population_2016 - all_age_ts$population_2015)/all_age_ts$population_2015

all_age_ts_diff <- all_age_ts[c("lad2014_name", "2001", "diff_2002", "diff_2003", "diff_2004", "diff_2005", "diff_2006", "diff_2007", "diff_2008", "diff_2009", "diff_2010", "diff_2011", "diff_2012", "diff_2013", "diff_2014", "diff_2015", "diff_2016")]        

all_age_ts_diff_long <- gather(data = all_age_ts_diff, key = "Year", value = "Percentage change", diff_2002:diff_2016) # This puts the data into long format (so males and females population counts are in the same field)

all_age_ts_diff_long$Year <- gsub("diff_", "", all_age_ts_diff_long$Year)

pdf("District_percentage_change_wsx_2016.pdf",  width = 11.7, height = 8.3) # a4 landscape
ggplot(data = all_age_ts_diff_long, aes(x = Year, y = `Percentage change`, group = lad2014_name, fill = lad2014_name)) +
  geom_bar(stat = "identity") +
  ggtitle("Year-on-year percentage change in population;\nWest Sussex Districts and Boroughs; Mid 2016 estimates") +
  labs(caption = "Source: Office for National Statistics licensed under the Open Government Licence. © Crown copyright 2017") +
  scale_y_continuous(breaks = seq(-.025, .025, 0.005), limits = c(-.025, .025), labels = percent) +
  scale_color_discrete(name = "District") +
  xlab("Year") + ylab("Percentage change") +
  facet_wrap(~ lad2014_name, nrow = 2) +
  ph_theme() +
  theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.5), axis.text = element_text(size = 10))
dev.off()

wsx_change <- all_ages_summary_cchange[c("lad2014_name","population_2015","population_2016")]

wsx_change$diff <- wsx_change$population_2016 - wsx_change$population_2015

wsx_change$p_diff <- (wsx_change$diff / wsx_change$population_2015) * 100

wsx_change
# % Change in population, for custom age bands, 0-15 16-64, 65+

# pyramid ####

# Load packages
library(tidyr); library(reshape2); library(plyr); library(scales); library(ggplot2)

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

LA_Codes <- c("E07000223","E07000224","E07000225","E07000226","E07000227","E07000228","E07000229")

df_reg_long <- read.csv("./mye16pyramiddata.csv", header = TRUE)

df_reg_long_wsx <- subset(df_reg_long, Area == "West Sussex")

df_reg_long_wsx = df_reg_long_wsx[,2:4] # we only need age group and population counts for each sex
df_reg_long_wsx = gather(data = df_reg_long_wsx, key = "Sex", value = "Population_reg", 2:3) # This puts the data into long format (so males and females population counts are in the same field)

# Axis labels can be formated in a number of ways. We need to have the absolute number (-10000 people would look weird) and it would also be good to include the comma separator (10,000 is easier to read than 10000).
# Using ggplot you can format both of these (by using labels = abs and labels = comma, respectively) but you cannot do them at the same time.

df_reg_long_wsx$age_band = factor(df_reg_long_wsx$age_band, levels = c("0 to 4","5 to 9", "10 to 14", "15 to 19", "20 to 24", "25 to 29", "30 to 34", "35 to 39", "40 to 44", "45 to 49", "50 to 54", "55 to 59", "60 to 64", "65 to 69", "70 to 74", "75 to 79", "80 to 84", "85 to 89","90+"))

# So we need to create a function that does both
abs_comma <- function (x, ...) {
  format(abs(x), ..., big.mark = ",", scientific = FALSE, trim = TRUE)}

# This is the population pyramid using actual population numbers
ggplot(data = df_reg_long_wsx, aes(x = age_band, y = Population_reg, fill = Sex)) +
  geom_bar(data = subset(df_reg_long_wsx, Sex=="Female"),
           stat = "identity") +
  geom_bar(data = subset(df_reg_long_wsx, Sex=="Male"),
           stat = "identity",
           position = "identity",
           mapping = aes(y = -Population_reg)) +
  scale_fill_manual(values = c("#6080c2", "#af9662"), breaks = c("Male","Female")) +
  coord_flip()+ 
  ph_theme() + 
  ggtitle("West Sussex population (number of residents);\nMid 2016 estimates") +
  labs(caption = "Source: Office for National Statistics licensed under the Open Government Licence. © Crown copyright 2017") +
  scale_y_continuous(breaks = seq(-35000, 35000, 5000), limits = c(-35000, 35000), labels = abs_comma) + 
  xlab("Age group") + ylab("Population") + 
  theme(axis.ticks.y = element_blank()) + 
  theme(legend.position = c(0.9, 0.9))


f_wsx <- sum(subset(df_reg_long, Area == "West Sussex", select = "Female")) 
m_wsx <- sum(subset(df_reg_long, Area == "West Sussex", select = "Male")) 
f_eng <- sum(subset(df_reg_long, Area == "England", select = "Female")) 
m_eng <- sum(subset(df_reg_long, Area == "England", select = "Male")) 

df_percentage_wsx <- subset(df_reg_long, Area == "West Sussex")
df_percentage_eng <- subset(df_reg_long, Area == "England")

df_percentage_wsx$p_female <- (df_percentage_wsx$Female / f_wsx) *100
df_percentage_wsx$p_male <- (df_percentage_wsx$Male / m_wsx) *100

df_percentage_eng$p_female <- (df_percentage_eng$Female / f_eng) *100
df_percentage_eng$p_male <- (df_percentage_eng$Male / m_eng)*100

df_percentage_wsx = df_percentage_wsx[c("age_band","p_female","p_male")] # we only need age group and population counts for each sex

# df_percentage_wsx = gather(data = df_percentage_wsx, key = "Sex", value = "Population_reg", 2:3) # This puts the data into long format (so males and females population counts are in the same field)

df_percentage_eng = df_percentage_eng[c("age_band","p_female","p_male")] # we only need age group and population counts for each sex

df_percentage_wsx$Area <- "West Sussex"
df_percentage_eng$Area <- "England"

df_perc <- merge(df_percentage_wsx, df_percentage_eng, by = "age_band")

df_perc <- read.csv("./lostit.csv",header = TRUE)

df_perc$age_band = factor(df_perc$age_band, levels = c("0 to 4","5 to 9", "10 to 14", "15 to 19", "20 to 24", "25 to 29", "30 to 34", "35 to 39", "40 to 44", "45 to 49", "50 to 54", "55 to 59", "60 to 64", "65 to 69", "70 to 74", "75 to 79", "80 to 84", "85 to 89","90+"))

df_perc <- arrange(df_perc, age_band)

ggplot(data = df_perc, aes(x = age_band, y = West.Sussex, fill = Sex)) +
  geom_bar(data = subset(df_perc, Sex=="Female"),
           stat = "identity") +
  geom_line(data = subset(df_perc, Sex=="Female"), aes(x = as.numeric(age_band), y = England), colour="#3B5A9B", size = 1.25) +
  geom_bar(data = subset(df_perc, Sex=="Male"),
           stat = "identity",
           position = "identity",
           mapping = aes(y = -West.Sussex)) +
  geom_line(data = subset(df_perc, Sex=="Male"), aes(x = as.numeric(age_band), y = -England), colour="#9b7c3b", size = 1.25) + 
  scale_fill_manual(values = c("#6080c2", "#af9662"), breaks = c("Male","Female")) +
  coord_flip()  + 
  ph_theme() + 
  ggtitle("Resident population (percentage of population in each in age group by sex);\nWest Sussex compared to England; Mid 2016 estimates") +  
  scale_y_continuous(breaks = seq(-8, 8, 2), limits = c(-8, 8), labels = abs_comma) + xlab("Age group") + ylab("Population (percentage of population in age group)") + 
  theme(axis.ticks.y = element_blank()) + theme(legend.position = c(0.9, 0.9)) + 
  labs(caption= "Note: Lines represent the resident population of England\nSource: Office for National Statistics licensed under the Open Government Licence. © Crown copyright 2017")


