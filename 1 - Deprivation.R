

# download.file("https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/464467/File_13_ID_2015_Clinical_Commissioning_Group_Summaries.xlsx", "./Projecting-Health/CCG_IMD_summary.xlsx", mode = "wb")
# 
# download.file("https://assets.publishing.service.gov.uk/government/uploads/system/uploads/attachment_data/file/464464/File_10_ID2015_Local_Authority_District_Summaries.xlsx", "./Projecting-Health/LTLA_imd_summary.xlsx", mode = "wb")
# 
# download.file("https://assets.publishing.service.gov.uk/government/uploads/system/uploads/attachment_data/file/464465/File_11_ID_2015_Upper-tier_Local_Authority_Summaries.xlsx", "./Projecting-Health/UTLA_imd_summary.xlsx", mode = "wb")
# 
library(easypackages)

# This uses the easypackages package to load several libraries at once. Note: it should only be used when you are confident that all packages are installed as it will be more difficult to spot load errors compared to loading each one individually.
libraries(c("readxl", "readr", "plyr", "dplyr", "ggplot2", "png", "tidyverse", "reshape2", "scales", "viridis", "rgdal", "officer", "flextable", "tmaptools", "lemon", "fingertipsR", "PHEindicatormethods", "xlsx", "leaflet", "htmltools"))

# If you have downloaded/cloned the github repo for this project you will need to make sure the filepath is recorded in the object github_repo_dir
github_repo_dir <- "~/Documents/Repositories/Projecting-Health"

# You must specify a set of chosen areas.

if(!(exists("Areas_to_include"))){
  print("There are no areas defined. Please create or load 'Areas_to_include' which is a character string of chosen areas.")
}

# Areas_to_include <- c("Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "NHS Coastal West Sussex CCG","West Sussex", "NHS Crawley CCG", "NHS Horsham and Mid Sussex CCG")

#Areas_to_include <- c("Basingstoke and Deane","East Hampshire","Eastleigh","Fareham","Gosport","Hart","Havant","New Forest","Rushmoor","Test Valley","Hampshire", "Winchester", "NHS North East Hampshire and Farnham CCG", "NHS North Hampshire CCG", "NHS South Eastern Hampshire CCG", "NHS West Hampshire CCG", "NHS Fareham and Gosport CCG", "Portsmouth")


Areas_to_include <- c("Eastbourne", "Hastings", "Lewes","Rother", "Wealden","Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing", "Brighton and Hove", "NHS Brighton and Hove CCG", "NHS Coastal West Sussex CCG", "NHS Crawley CCG","NHS Eastbourne, Hailsham and Seaford CCG", "NHS Hastings and Rother CCG","NHS High Weald Lewes Havens CCG", "NHS Horsham and Mid Sussex CCG", "West Sussex", "East Sussex", "England")

imd <- read_csv("https://assets.publishing.service.gov.uk/government/uploads/system/uploads/attachment_data/file/467774/File_7_ID_2015_All_ranks__deciles_and_scores_for_the_Indices_of_Deprivation__and_population_denominators.csv")

# Summary

names(CCG_income_scale)

CCG_income_scale <- read_excel("./Projecting-Health/CCG_IMD_summary.xlsx", sheet = "Income")%>% 
  rename(AreaCode = `Clinical Commissioning Group Code (2015)`,
         AreaName = `Clinical Commissioning Group Name (2015)`) %>% 
  select(AreaCode, AreaName, `Income - Average score`, `Income - Scale`)

CCG_empl_scale <- read_excel("./Projecting-Health/CCG_IMD_summary.xlsx", sheet = "Employment")%>% 
  rename(AreaCode = `Clinical Commissioning Group Code (2015)`,
         AreaName = `Clinical Commissioning Group Name (2015)`) %>% 
  select(AreaCode, AreaName, `Employment - Average score`, `Employment - Scale`)

CCG_IMD_summary <- read_excel("./Projecting-Health/CCG_IMD_summary.xlsx", sheet = "IMD") %>% 
  rename(AreaCode = `Clinical Commissioning Group Code (2015)`,
         AreaName = `Clinical Commissioning Group Name (2015)`) %>% 
  left_join(CCG_income_scale, by = c("AreaCode", "AreaName")) %>% 
  left_join(CCG_empl_scale, by = c("AreaCode", "AreaName")) %>% 
  mutate(Type = "CCG")

rm(CCG_income_scale, CCG_empl_scale)

LTLA_income_scale <- read_excel("./Projecting-Health/LTLA_IMD_summary.xlsx", sheet = "Income")%>% 
  rename(AreaCode = `Local Authority District code (2013)`,
         AreaName = `Local Authority District name (2013)`) %>% 
  select(AreaCode, AreaName, `Income - Average score`, `Income - Scale`)

LTLA_empl_scale <- read_excel("./Projecting-Health/LTLA_IMD_summary.xlsx", sheet = "Employment")%>% 
  rename(AreaCode = `Local Authority District code (2013)`,
         AreaName = `Local Authority District name (2013)`) %>% 
  select(AreaCode, AreaName, `Employment - Average score`, `Employment - Scale`)

LTLA_IMD_summary <- read_excel("./Projecting-Health/LTLA_IMD_summary.xlsx", sheet = "IMD") %>% 
  rename(AreaCode = `Local Authority District code (2013)`,
         AreaName = `Local Authority District name (2013)`) %>% 
  left_join(LTLA_income_scale, by = c("AreaCode", "AreaName")) %>% 
  left_join(LTLA_empl_scale, by = c("AreaCode", "AreaName")) %>% 
  mutate(Type = "LTLA")

rm(LTLA_income_scale, LTLA_empl_scale)

UTLA_income_scale <- read_excel("./Projecting-Health/UTLA_IMD_summary.xlsx", sheet = "Income") %>% 
  rename(AreaCode = `Upper Tier Local Authority District code (2013)`,
         AreaName = `Upper Tier Local Authority District name (2013)`) %>% 
  select(AreaCode, AreaName, `Income - Average score`, `Income - Scale`)

UTLA_empl_scale <- read_excel("./Projecting-Health/UTLA_IMD_summary.xlsx", sheet = "Employment") %>% 
  rename(AreaCode = `Upper Tier Local Authority District code (2013)`,
         AreaName = `Upper Tier Local Authority District name (2013)`) %>% 
  select(AreaCode, AreaName, `Employment - Average score`, `Employment - Scale`)

UTLA_IMD_summary <- read_excel("./Projecting-Health/UTLA_IMD_summary.xlsx", sheet = "IMD") %>% 
  rename(AreaCode = `Upper Tier Local Authority District code (2013)`,
         AreaName = `Upper Tier Local Authority District name (2013)`) %>% 
  left_join(UTLA_income_scale, by = c("AreaCode", "AreaName")) %>% 
  left_join(UTLA_empl_scale, by = c("AreaCode", "AreaName")) %>% 
  mutate(Type = "UTLA")

rm(UTLA_empl_scale, UTLA_income_scale)

IMD_summary <- LTLA_IMD_summary %>% 
  bind_rows(UTLA_IMD_summary) %>% 
  bind_rows(CCG_IMD_summary) %>% 
  select(AreaCode,AreaName,Type, `IMD - Average score`,`IMD - Rank of average score`,`IMD - Proportion of LSOAs in most deprived 10% nationally`,`IMD - Extent`,`Income - Average score`,`Income - Scale`,`Employment - Average score`,`Employment - Scale`)

rm(CCG_IMD_summary, LTLA_IMD_summary, UTLA_IMD_summary)

Selected_summary_IMD <- IMD_summary %>% 
  filter(AreaName %in% Areas_to_include)

write.csv(Selected_summary_IMD, "./Projecting-Health/Selected_summary_IMD.csv", row.names = FALSE)


# Map

selected_lsoas <- imd %>% 
  filter(`Local Authority District name (2013)` %in% Areas_to_include) %>% 
  select(`LSOA code (2011)`,`LSOA name (2011)`,`Local Authority District code (2013)`,`Local Authority District name (2013)`,`Index of Multiple Deprivation (IMD) Score`,`Index of Multiple Deprivation (IMD) Rank (where 1 is most deprived)`,`Index of Multiple Deprivation (IMD) Decile (where 1 is most deprived 10% of LSOAs)`,`Income Score (rate)`,`Income Rank (where 1 is most deprived)`,`Income Decile (where 1 is most deprived 10% of LSOAs)`,`Employment Score (rate)`,`Employment Rank (where 1 is most deprived)`,`Employment Decile (where 1 is most deprived 10% of LSOAs)`) %>% 
  mutate(IMD_decile_within_all_coverage = ntile(`Index of Multiple Deprivation (IMD) Score`, 10)) %>% 
  mutate(IMD_decile_within_all_coverage = factor(11 - IMD_decile_within_all_coverage, levels = c(1,2,3,4,5,6,7,8,9,10))) %>% 
  group_by(`Local Authority District name (2013)`) %>% 
  mutate(IMD_decile_within_ltla = ntile(`Index of Multiple Deprivation (IMD) Score`, 10)) %>% 
  mutate(IMD_decile_within_ltla = factor(11 - IMD_decile_within_ltla, levels = c(1,2,3,4,5,6,7,8,9,10))) %>% 
  ungroup()

# Download the LSOA boundary file if it is not already in the local directory
url = 'http://geoportal.statistics.gov.uk/datasets/da831f80764346889837c72508f046fa_2.zip'
file = paste("./Projecting-Health",basename(url),sep='/')
if (!file.exists(file)) {
  download.file(url, file) 
  unzip(file,exdir= "./Projecting-Health") # unzip the folder into the directory
}

# layerName is the name of the unzipped shapefile without file type extensions 
layerName = "Lower_Layer_Super_Output_Areas_December_2011_Generalised_Clipped__Boundaries_in_England_and_Wales"  
# Read in the data
data_projected = readOGR(dsn= "./Projecting-Health", layer = "Lower_Layer_Super_Output_Areas_December_2011_Generalised_Clipped__Boundaries_in_England_and_Wales") 

data_projected <- subset(data_projected, data_projected@data$lsoa11cd %in% selected_lsoas$`LSOA code (2011)`)

data_projected@data$lsoa11cd = factor(data_projected@data$lsoa11cd)
data_projected@data$lsoa11nm = factor(data_projected@data$lsoa11nm)

### create an interactive map ####

# Transform the projection of the map to fit the Leaflet expectation
data_projected <- spTransform(data_projected, CRS("+init=epsg:4326"))

# Append the IMD data to the
IMD2015_map = append_data(data_projected, selected_lsoas, key.shp = "lsoa11cd", key.data = "LSOA code (2011)") 

IDPal = colorFactor(c("#0000FF", "#2080FF", "#40E0FF" , "#70FFD0", "#90FFB0", "#C0E1B0", "#E0FFA0", "#E0FF70", "#F0FF30", "#FFFF00"), c(1,2,3,4,5,6,7,8,9,10))


LTUA_boundary <- readOGR(dsn = "./GIS/Boundaries", layer = "Local_Authority_Districts_May_2018_UK_BUC")

selected_ltua <- subset(LTUA_boundary, LTUA_boundary@data$lad18nm %in% Areas_to_include)

selected_ltua <- spTransform(selected_ltua, CRS("+init=epsg:4326"))


leaf <- leaflet() %>%
  addTiles() %>%
  addPolygons(data=IMD2015_map, 
              stroke=FALSE, 
              smoothFactor = 0.1, 
              fillOpacity = .8, 
              color = ~IDPal(IMD2015_map@data$`Index of Multiple Deprivation (IMD) Decile (where 1 is most deprived 10% of LSOAs)`), 
              group = "Show national decile")%>%
  addScaleBar(position = "bottomleft") %>% 
  addPolygons(data=IMD2015_map, 
              stroke=FALSE, 
              smoothFactor = 0.1, 
              fillOpacity = .8, 
              color = ~IDPal(IMD2015_map@data$IMD_decile_within_all_coverage), 
              group = "Show covered area decile") %>%
  addPolygons(data=IMD2015_map, 
              stroke=FALSE, 
              smoothFactor = 0.1, 
              fillOpacity = .8, 
              color = ~IDPal(IMD2015_map@data$IMD_decile_within_ltla), 
              group = "Show within area decile") %>%
  addPolygons(data = IMD2015_map, 
              stroke = TRUE, 
              col = "#000000", 
              weight = 0.3, 
              fillOpacity = 0) %>%
  addPolygons(data = selected_ltua, 
              stroke = TRUE, 
              color = "#000000",
              fill = FALSE,
              weight = 2, 
              group = "Show national decile") %>%  
  addPolygons(data = selected_ltua, 
              stroke = TRUE, 
              color = "#000000",
              fill = FALSE,
              weight = 2, 
              group = "Show within area decile") %>%  
  addPolygons(data = selected_ltua, 
              stroke = TRUE, 
              color = "#000000",
              fill = FALSE,
              weight = 2, 
              group = "Show covered area decile") %>%  
  addLegend(position = "bottomright", 
            colors=c("#0000FF", "#2080FF", "#40E0FF" , "#70FFD0", "#90FFB0", "#C0E1B0", "#E0FFA0", "#E0FF70", "#F0FF30", "#FFFF00"),
            labels = c("Most deprived 10%","decile 2","decile 3","decile 4","decile 5", "decile 6", "decile 7", "decile 8", "decile 9", "Least deprived 10%"), 
            title = "Decile", 
            opacity = 1) %>%
  addLayersControl(
    baseGroups = c("Show covered area decile","Show within area decile","Show national decile"), 
    options = layersControlOptions(collapsed = FALSE),
    position = "topright") 

browsable(
  tagList(list(
    tags$head(
      tags$style(
        ".leaflet .legend i{
            border-radius: 75%;
            width: 12px;
            height: 12px;
            padding: 0px;
        }
         .leaflet .legend {
         line-height: 10px;
         font-size: 11px;
         background: #ccc;
         }
        "
      )
    ),
    leaf
  ))
)

# CCG relevant score

gp_dep_score <- fingertips_data(91872, AreaTypeID = 7) %>% 
  filter(AreaType == "GP") %>% 
  mutate(IMD_decile = ntile(Value, 10)) %>% 
  mutate(IMD_decile = factor(11 - IMD_decile, levels = c(1,2,3,4,5,6,7,8,9,10))) %>% 
  filter(ParentCode %in% Selected_summary_IMD$AreaCode) %>% 
  group_by(ParentName, IMD_decile) %>% 
  summarise(Practices = n()) %>% 
  complete(IMD_decile, fill = list(Practices = 0)) %>% 
  mutate(Proportion = Practices / sum(Practices)) %>% 
  ungroup()

gp_dep_score_table <- gp_dep_score %>% 
  select(ParentName, IMD_decile, Practices) %>% 
  spread(IMD_decile, Practices)

write.csv(gp_dep_score_table, "./Projecting-Health/gp_dep_score.csv", row.names = FALSE)
