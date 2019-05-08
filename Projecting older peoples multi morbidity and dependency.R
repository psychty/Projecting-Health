
# Projecting older peoples ill health, multi-morbidity and dependency

library(readr)
library(plyr)
library(dplyr)
library(ggplot2)
library(png)
library(tidyverse)
library(readxl)
library(reshape2)
library(scales)
library(viridis)
library(rgdal)
library(officer)
library(flextable)
library(tmaptools)
library(lemon)
library(leaflet)
library(fingertipsR)
library(PHEindicatormethods)

options(scipen = 9)


capwords = function(s, strict = FALSE) {
  cap = function(s) paste(toupper(substring(s, 1, 1)),
                          {s = substring(s, 2); if(strict) tolower(s) else s},sep = "", collapse = " " )
  sapply(strsplit(s, split = " "), cap, USE.NAMES = !is.null(names(s)))}

ph_theme = function(){
  theme( 
    axis.text.y = element_text(colour = "#000000", size = 9), 
    axis.text.x = element_text(colour = "#000000", angle = 0, hjust = 1, vjust = .5, size = 8), 
    axis.title =  element_text(colour = "#000000", size = 9, face = "bold"),     
    plot.title = element_text(colour = "#000000", face = "bold", size = 10),    
    plot.subtitle = element_text(colour = "#000000", size = 9),
    panel.grid.major.x = element_blank(), 
    panel.grid.minor.x = element_blank(),
    panel.background = element_rect(fill = "#FFFFFF"), 
    panel.grid.major.y = element_line(colour = "#E7E7E7", size = .3),
    panel.grid.minor.y = element_blank(), 
    strip.text = element_text(colour = "#000000", size = 10, face = "bold"), 
    strip.background = element_blank(), 
    axis.ticks = element_line(colour = "#E7E7E7"), 
    legend.position = "bottom", 
    legend.title = element_text(colour = "#000000", size = 10, face = "bold"), 
    legend.background = element_rect(fill = "#ffffff"), 
    legend.key = element_rect(fill = "#ffffff", colour = "#ffffff"), 
    legend.text = element_text(colour = "#000000", size = 9), 
    axis.line = element_line(colour = "#dbdbdb")
  ) 
}

# Intro

paste0("The capacity to understand and forecast future demand has increasingly been seen as an important component of commissioning as an activity across health and social care. The Projecting Older People Population Information System (POPPI) is database is a user-friendly, practical and straightforward way to analyse population data, identify key characteristics within that population, project numbers into the future, and compare future populations against performance data. POPPI was originally developed by the Institute of Public Care (IPC) for the Care Services Efficiency Delivery Programme (CSED). The system is now provided solely by the Institute of Public Care on licence from the Department of Health.")

paste0("The system applies assumptions about conditions at national prevalence levels to local population estimates and projections. Currently the POPPI system uses the 2014 based population projections by the Office for National Statistics (ONS). ONS has since published 2016 based population projections. The updated projections give a slightly more conservative estimate of the increases expected later in the projections (beyond the next decade) and particularly affect projections for older old age groups. Based on the 2014 projections used in the POPPI system, it is anticipated that by 2035, there will be 109,400 over 65's living alone, whereas the same assumptions applied to the 2016 projections anticipate that there will be 105,800 over 65's living alone, a reduction in over 3,500 residents.")

POPPI_site_ill_health <- data.frame(Age = c("65+","65+","65+","65+","75+","65+","65+","65+","65+","65+","65+"), Condition = c("Diabetes", "Dementia", "Depression", "Severe Depression", "Registrable eye conditions", "Severe hearing loss", "Bronchitis and emphysema", "Longstanding health condition caused by a stroke", "Bladder problem at least once a week", "Fall", "Admission to hospital as a result of a fall"), Number = c(24100,14600,16800,5300,5900,16200,3250,4500,32500,52700,4100))

paste0("We know that projecting ill health in this way has a number of issues. First, as with all projections, they are best guesses based on applying prevalence rates nationally to local populations, which assumes West Sussex population is similar to England overall in terms of its characteristics. Second, the assumptions applied are likely to be from out of date national surveys more than a decade old. Third, POPPI uses 2014 based ONS projections, whilst more recent (2016 based) are publically available. ONS projections are largely based on trends seen in the previous five years, and so can change quite significantly every release. Whichever year ONS projections, they do not include additional information on local housing policies and housing stock decisions, which local authorities often produce to augment the ONS release.")

paste0("The method used for POPPI based estimates is to apply what we know about a disease prevalence to the number of people expected to be living in West Sussex in the future. This assumes that prevalence rates will remain the same (e.g. the same percentage of people will develop a condition in 2018 and in 2038. This is unlikely to be the case with developments in preventative treatments, diagnostic improvements and education around risk factors.")

paste0("Current epidemiological work indicates that the risk factors observed in those who were younger adults in the past (now aged 65+) are different to those who are younger adults now (to be aged 65+ in the future). As such, it is difficult to conclude that generations of older people in the future will experience ill health in the same way as older people now.")

paste0("Instead, we would be more accurate in attempting to apply what we think the prevalence of a condition or conditions will be like in the future, and then apply that figure to the estimated population projection.")

paste0("A second fundamental issue is the measurement of co-morbidities, which is increasingly receiving traction among commentators as the necessary paradigm shift required to address health and social care in the future (Kingston et al, 2018).")

paste0("There are several steps to this work. First we try to assess the current prevalence of health conditions and multi morbidity among older people in West Sussex, then we look to project what the morbidity landscape of older people in West Sussex could look like in the future.")

# The current picture ####

# QOF registers

paste0("What do we know about the observed prevalence of ill health in West Sussex?")

if(!(file.exists("./QOF"))){
  dir.create("./QOF")
}

if(!(file.exists("./QOF/QOF_1718.zip"))){
download.file("https://files.digital.nhs.uk/72/CB2869/qof-1718-csv.zip", "./QOF/QOF_1718.zip", mode = "wb")
  unzip("./QOF/QOF_1718.zip", exdir = "./QOF")
}

QOF_Prevalence <- read_csv("./QOF/PREVALENCE.csv", col_types = cols(PRACTICE_CODE = col_character(),INDICATOR_GROUP_CODE = col_character(),REGISTER = col_integer(),PATIENT_LIST_TYPE = col_character(),PATIENT_LIST_SIZE = col_integer()))

ind_group <- data.frame(Condition_code = c("AF", "CHD", "CVDPP", "HF", "HYP", "PAD", "STIA", "AST", "COPD", "OB", "CAN", "CKD", "DM", "PC", "DEM", "DEP", "EP", "LD", "MH", "OST", "RA"), Condition = c("Atrial fibrillation", "Coronary heart disease", "Cardiovascular disease (30-74)", "Heart failure", "Hypertension", "Peripheral arterial disease", "Stroke and transient ischaemic attack", "Asthma", "Chronic obstructive pulmonary disease", "Obesity (18+)", "Cancer", "Chronic kidney disease (18+)", "Diabetes mellitus (17+)", "Palliative care", "Dementia", "Depression (18+)", "Epilepsy (18+)", "Learning disabilities", "Mental health", "Osteoporosis (50+)", "Rheumatoid arthritis (16+)"), disease_group = c(rep("Cardiovascular", 7),rep("Respiratory", 2),rep("Lifestyle", 1), rep("High dependency and long term conditions", 4), rep("Mental health and neurology", 5), rep ("Musculoskeletal",2))) 

QOF_Indicators <- read_csv("./QOF/INDICATOR_MAPPINGS.csv", col_types = cols(INDICATOR_CODE = col_character(),INDICATOR_DESCRIPTION = col_character(),INDICATOR_POINT_VALUE = col_integer(),INDICATOR_GROUP_CODE = col_character(),INDICATOR_GROUP_DESCRIPTION = col_character(),DOMAIN_CODE = col_character(),DOMAIN_DESCRIPTION = col_character(),PATIENT_LIST_TYPE = col_character())) %>% 
  select(INDICATOR_GROUP_CODE, INDICATOR_GROUP_DESCRIPTION) %>% 
  unique()

QOF_org <- read_csv("./QOF/ORGANISATION_REFERENCE.csv", col_types = cols(PRACTICE_CODE = col_character(),PRACTICE_NAME = col_character(),CCG_CODE = col_character(),CCG_GEOGRAPHY_CODE = col_skip(),CCG_NAME = col_character(),STP_CODE = col_skip(),STP_NAME = col_skip(),SUBREGION_CODE = col_character(),SUBREGION_GEOGRAPHY_CODE = col_skip(),SUBREGION_NAME = col_character(),REGION_CODE = col_character(),REGION_GEOGRAPHY_CODE = col_skip(),REGION_NAME = col_character(),COUNTRY = col_skip(),REVISED_MAXIMUM_POINTS = col_skip()))

QOF_Prevalence <- QOF_Prevalence %>% 
  left_join(QOF_org, by = "PRACTICE_CODE") %>% 
  left_join(ind_group, by = c("INDICATOR_GROUP_CODE" = "Condition_code")) %>% 
  select(PRACTICE_CODE, PRACTICE_NAME, CCG_CODE, CCG_NAME, SUBREGION_CODE, SUBREGION_NAME, REGION_CODE, REGION_NAME, INDICATOR_GROUP_CODE, Condition, disease_group, REGISTER, PATIENT_LIST_SIZE) %>% 
  rename(Number_on_register = REGISTER, 
         Appropriate_population = PATIENT_LIST_SIZE) %>% 
  mutate(Prevalence = Number_on_register / Appropriate_population)

QOF_Prevalence_Region <- QOF_Prevalence %>% 
  group_by(REGION_NAME, REGION_CODE, INDICATOR_GROUP_CODE, Condition, disease_group) %>% 
  summarise(Number_on_register = sum(Number_on_register, na.rm = TRUE),
            Appropriate_population = sum(Appropriate_population, na.rm = TRUE)) %>% 
  mutate(Prevalence = Number_on_register / Appropriate_population,
         Area_type = "Region") %>% 
  ungroup() %>% 
  rename(Area_name = REGION_NAME,
         Area_code = REGION_CODE)

QOF_Prevalence_Subregion <- QOF_Prevalence %>% 
  group_by(SUBREGION_NAME, SUBREGION_CODE, INDICATOR_GROUP_CODE, Condition, disease_group) %>% 
  summarise(Number_on_register = sum(Number_on_register, na.rm = TRUE),
            Appropriate_population = sum(Appropriate_population, na.rm = TRUE)) %>% 
  mutate(Prevalence = Number_on_register / Appropriate_population,
         Area_type = "Region") %>% 
  ungroup() %>% 
  rename(Area_name = SUBREGION_NAME,
         Area_code = SUBREGION_CODE)

QOF_Prevalence_CCG <- QOF_Prevalence %>% 
  group_by(CCG_CODE, CCG_NAME, INDICATOR_GROUP_CODE, Condition, disease_group) %>% 
  summarise(Number_on_register = sum(Number_on_register, na.rm = TRUE),
            Appropriate_population = sum(Appropriate_population, na.rm = TRUE)) %>% 
  mutate(Prevalence = Number_on_register / Appropriate_population,
         Area_type = "CCG") %>% 
  ungroup() %>% 
  rename(Area_name = CCG_NAME,
         Area_code = CCG_CODE)

QOF_Prevalence_Wsx <- QOF_Prevalence %>% 
  filter(CCG_CODE %in% c("09G", "09H", "09X")) %>% 
  group_by(INDICATOR_GROUP_CODE, Condition, disease_group) %>% 
  summarise(Number_on_register = sum(Number_on_register, na.rm = TRUE),
            Appropriate_population = sum(Appropriate_population, na.rm = TRUE)) %>% 
  mutate(Prevalence = Number_on_register / Appropriate_population,
         Area_name = "West Sussex",
         Area_code = "West Sussex",
         Area_type = "West Sussex") %>% 
  ungroup() 

QOF_Prevalence_England <- QOF_Prevalence %>% 
  group_by(INDICATOR_GROUP_CODE, Condition, disease_group) %>% 
  summarise(Number_on_register = sum(Number_on_register, na.rm = TRUE),
            Appropriate_population = sum(Appropriate_population, na.rm = TRUE)) %>% 
  mutate(Prevalence = Number_on_register / Appropriate_population,
         Area_name = "England",
         Area_code = "England",
         Area_type = "England") %>% 
  ungroup()

QOF_Prevalence_Practice_Wsx <- QOF_Prevalence %>% 
  filter(CCG_CODE %in% c("09G", "09H", "09X")) %>% 
  select(INDICATOR_GROUP_CODE, Condition, disease_group, Number_on_register, Appropriate_population, Prevalence, PRACTICE_CODE, PRACTICE_NAME) %>% 
  mutate(Area_type = "Practice") %>% 
  rename(Area_name = "PRACTICE_NAME",
         Area_code = "PRACTICE_CODE")

QOF_prev <- QOF_Prevalence_Wsx %>% 
  bind_rows(QOF_Prevalence_Practice_Wsx, QOF_Prevalence_CCG, QOF_Prevalence_Subregion, QOF_Prevalence_Region, QOF_Prevalence_England) %>% 
  mutate(Prevalence_LCI = wilson_lower(Number_on_register, Appropriate_population, confidence = 0.95),
         Prevalence_UCI = wilson_upper(Number_on_register, Appropriate_population, confidence = 0.95))

QOF_prev_theme = function(){
  theme( 
    axis.text.y = element_text(colour = "#000000", size = 8), 
    axis.text.x = element_text(colour = "#000000", angle = 0, hjust = 0.5, size = 8), 
    axis.title =  element_text(colour = "#000000", size = 10, face = "bold"),
    axis.ticks = element_line(colour = "#6d6d6d"), 
    plot.title = element_text(colour = "#000000", face = "bold", size = 10),    
    plot.subtitle = element_text(colour = "#000000", size = 10),
    panel.background = element_rect(fill = "#FFFFFF"), 
    panel.grid.minor.x = element_line(colour = "#E7E7E7", size = .3),
    panel.grid.major = element_line(colour = "#E7E7E7", size = .3), 
    panel.grid.minor.y = element_blank(), 
    panel.grid.major.y = element_blank(),
    strip.text = element_text(colour = "#000000", size = 10, face = "bold"), 
    strip.background = element_blank(), 
    legend.position = "bottom", 
    legend.title = element_text(colour = "#000000", size = 8, face = "bold"), 
    legend.background = element_rect(fill = "#ffffff"), 
    legend.key = element_rect(fill = "#ffffff", colour = "#ffffff"), 
    legend.text = element_text(colour = "#000000", size = 8), 
    axis.line = element_line(colour = "#dbdbdb")
  ) 
}

QOF_prev$disease_group <- factor(QOF_prev$disease_group, levels = rev(c("Cardiovascular", "Respiratory", "Lifestyle", "High dependency and long term conditions", "Mental health and neurology", "Musculoskeletal")))

QOF_prev_wsx <- QOF_prev %>% 
  filter(Area_name == "West Sussex") %>% 
  arrange(disease_group, Prevalence)

QOF_prev_wsx$Condition <- factor(QOF_prev_wsx$Condition, levels = QOF_prev_wsx$Condition)

ggplot(data = QOF_prev_wsx, aes(x =  Condition, y = Prevalence, fill = disease_group)) + 
  geom_bar(stat = "identity", position = position_dodge()) + 
  geom_errorbar(aes(ymin = Prevalence_LCI, ymax = Prevalence_UCI), 
                width = 0.5, 
                position = position_dodge(0.9)) +
  coord_flip() + 
  labs(title = "Recorded disease prevalence - QOF 2017/18",
       subtitle = "West Sussex GPs",
       x = "Condition",
       y = "Prevalence (percentage)") + 
  scale_y_continuous(breaks = seq(0, .20, .01), 
                     expand = c(0.01, 0), 
                     limits = c(0,.2), 
                     labels = percent) + 
  scale_fill_manual(values = rev(c("#12004b","#594c81","#6495ed","#2790ab","#38cef5","#6d6d6d")), 
                    breaks = c("Cardiovascular", "Respiratory", "Lifestyle", "High dependency and long term conditions", "Mental health and neurology", "Musculoskeletal"),
                    name = "Condition group") +
  QOF_prev_theme() 
  
QOF_prev_wsx_Eng <- QOF_prev %>% 
  filter(Area_name %in% c("West Sussex", "England"))

QOF_prev_wsx_Eng$Condition <- factor(QOF_prev_wsx_Eng$Condition, levels = QOF_prev_wsx$Condition)

ggplot(data = QOF_prev_wsx_Eng, aes(x =  Condition, y = Prevalence, fill = Area_name)) + 
  geom_bar(stat = "identity", position = position_dodge()) + 
  geom_errorbar(aes(ymin = Prevalence_LCI, ymax = Prevalence_UCI), 
                width = 0.5, 
                position = position_dodge(0.9)) +
  coord_flip() + 
  labs(title = "Recorded disease prevalence - QOF 2017/18",
       subtitle = "West Sussex GPs",
       x = "Condition",
       y = "Prevalence (percentage)") + 
  scale_y_continuous(breaks = seq(0, .20, .02), 
                     expand = c(0.01, 0), 
                     limits = c(0,.2), 
                     labels = percent) + 
  scale_fill_manual(values = c("#a7a7a7","#12004b"),
                    breaks = c("West Sussex", "England"),
                    name = "Area") +
  QOF_prev_theme()

Prevalence_LCI <- QOF_prev_wsx_Eng %>% 
  select(Area_name, Condition, Prevalence_LCI) %>% 
  spread(Area_name, Prevalence_LCI) %>% 
  rename(England_LCI = England,
         Wsx_LCI = `West Sussex`)

Prevalence_UCI <- QOF_prev_wsx_Eng %>% 
  select(Area_name, Condition, Prevalence_UCI) %>% 
  spread(Area_name, Prevalence_UCI)%>% 
  rename(England_UCI = England,
         Wsx_UCI = `West Sussex`)

Prevalence <- left_join(Prevalence_LCI, Prevalence_UCI, by = "Condition") %>% 
  mutate(Significant_diff_wsx = ifelse(Wsx_LCI > England_UCI, "Significantly higher", ifelse(Wsx_UCI < England_LCI, "Significantly lower", "Similar")))

paste0("The above figure shows that disease registers in West Sussex GPs have a higher prevalence, compared to England, for 13 conditions (including hypertension, coronary heart disease, atrial fibrillation, dementia, depression, and cancer). West Sussex has significantly lower prevalence rates of chronic obstructive pulmonary disease, obsesity, and peripheral arterial disease. It should be noted that these figures represent the percentage of people with recorded diagnoses, and as such they represent the GP practice diagnosing appropriately. It may underrepresent true prevalence among those who do not go to their GPs.")

# QOF dementia over 65s

# Is there any way to scowel the internets to find the latest available month

Recorded_dementia_65_plus <- read_csv(url("https://files.digital.nhs.uk/47/047F85/rec-dem-diag-Aug-2018-csv.csv"), col_types = cols(ACH_DATE = col_character(),PRACTICE_CODE = col_character(),NAME = col_character(),COMMISSIONER_ORGANISATION_CODE = col_character(),GEOGRAPHY_CODE = col_character(),Measure = col_character(),Value = col_integer())) %>% 
  filter(COMMISSIONER_ORGANISATION_CODE %in% c("09G", "09H", "09X"),
         ACH_DATE == "31AUG2018") %>% 
  spread(Measure, Value) %>% 
  select(ACH_DATE,PRACTICE_CODE,NAME,COMMISSIONER_ORGANISATION_CODE,GEOGRAPHY_CODE,DEMENTIA_REGISTER_65_PLUS,PAT_LIST_65_PLUS) %>% 
  mutate(Prevalence_65_plus = DEMENTIA_REGISTER_65_PLUS/PAT_LIST_65_PLUS,
         Prevalence_65_plus_lci = wilson_lower(DEMENTIA_REGISTER_65_PLUS, PAT_LIST_65_PLUS, .95),
         Prevalence_65_plus_uci = wilson_upper(DEMENTIA_REGISTER_65_PLUS, PAT_LIST_65_PLUS, .95)) %>% 
  arrange(Prevalence_65_plus) %>% 
  mutate(Label = paste0(PRACTICE_CODE, " - ", NAME))

Recorded_dementia_65_plus$Label <- factor(Recorded_dementia_65_plus$Label, levels = Recorded_dementia_65_plus$Label)

Recorded_dem_65_wsx <- Recorded_dementia_65_plus %>% 
  summarise(DEMENTIA_REGISTER_65_PLUS = sum(DEMENTIA_REGISTER_65_PLUS, na.rm = TRUE),
            PAT_LIST_65_PLUS = sum(PAT_LIST_65_PLUS, na.rm = TRUE)) %>% 
  mutate(Prevalence_65_plus = DEMENTIA_REGISTER_65_PLUS/PAT_LIST_65_PLUS,
         Prevalence_65_plus_lci = wilson_lower(DEMENTIA_REGISTER_65_PLUS, PAT_LIST_65_PLUS, .95),
         Prevalence_65_plus_uci = wilson_upper(DEMENTIA_REGISTER_65_PLUS, PAT_LIST_65_PLUS, .95))

Recorded_dementia_65_plus$Significance_compared_WSx <- factor(ifelse(Recorded_dementia_65_plus$Prevalence_65_plus_lci > Recorded_dem_65_wsx$Prevalence_65_plus_uci, "Significantly higher", ifelse(Recorded_dementia_65_plus$Prevalence_65_plus_uci < Recorded_dem_65_wsx$Prevalence_65_plus_lci, "Significantly lower", "Similar")), levels = c("Significantly lower", "Similar", "Significantly higher"))

ggplot(Recorded_dementia_65_plus, aes(x = Label, y = Prevalence_65_plus, fill = Significance_compared_WSx))+
  geom_bar(stat = "identity") +
  geom_bar(stat = "identity", position = position_dodge()) + 
  geom_errorbar(aes(ymin = Prevalence_65_plus_lci, ymax = Prevalence_65_plus_uci), 
                width = 0.5, 
                position = position_dodge(0.9)) +
  scale_y_continuous(limits = c(0,.16), breaks = seq(0,.16,.01), labels = percent, expand = c(0,0.001)) +
  labs(title = "Recorded dementia prevalence - QOF; patients aged 65+",
       subtitle = paste0("West Sussex GPs; as at ", "August 2018"), 
       x = "Practice",
       y = "Prevalence (percentage)",
       caption = "Note: dashed lines either side of the West Sussex value (in red) are the 95% confidence limits.\nThese are used to calculate whether the prevalence in a practice is significantly different compared to the whole county.") + 
  coord_flip() +
  scale_fill_manual(values = c("#5555E6", "#FFC000", "#BED2FF"), name = "Compared to West Sussex overall") +
  geom_hline(yintercept = Recorded_dem_65_wsx$Prevalence_65_plus_lci, col = "#ff0015", lty = "dashed", lwd = 0.8) +
  geom_hline(aes(yintercept = Recorded_dem_65_wsx$Prevalence_65_plus, linetype = "West Sussex\npatients combined"), col = "#ff0015", lwd = 1.2) +
  geom_hline(yintercept = Recorded_dem_65_wsx$Prevalence_65_plus_uci, col = "#ff0015", lty = "dashed", lwd = 0.8) +
 QOF_prev_theme() +
  theme(axis.text.y = element_text(size = 6, hjust = 0)) +
  scale_linetype_manual(name = "", values = 1)# +
  #facet_rep_grid( ~ COMMISSIONER_ORGANISATION_CODE, repeat.tick.labels = F)

paste0("The above figure shows wide variation among GP populations in West Sussex in terms of the proportion of patients aged 65 and over who are on disease registers for dementia. The latest prevalence is highest in ", capwords(as.character(subset(Recorded_dementia_65_plus, Prevalence_65_plus == max(Recorded_dementia_65_plus$Prevalence_65_plus, na.rm = TRUE), select = "NAME")), strict = TRUE), " (", round(max(Recorded_dementia_65_plus$Prevalence_65_plus, na.rm = TRUE)*100,1), "%, ", subset(Recorded_dementia_65_plus, Prevalence_65_plus == max(Recorded_dementia_65_plus$Prevalence_65_plus, na.rm = TRUE), select = "DEMENTIA_REGISTER_65_PLUS"), " patients) and this is ", round(max(Recorded_dementia_65_plus$Prevalence_65_plus, na.rm = TRUE)/min(Recorded_dementia_65_plus$Prevalence_65_plus, na.rm = TRUE),0), " times greater than in the lowest prevalence GP population (",  capwords(as.character(subset(Recorded_dementia_65_plus, Prevalence_65_plus == min(Recorded_dementia_65_plus$Prevalence_65_plus, na.rm = TRUE), select = "NAME")), strict = TRUE), ", ", round(min(Recorded_dementia_65_plus$Prevalence_65_plus, na.rm = TRUE)*100,1), "% of patients aged 65+, ", subset(Recorded_dementia_65_plus, Prevalence_65_plus == min(Recorded_dementia_65_plus$Prevalence_65_plus, na.rm = TRUE), select = "DEMENTIA_REGISTER_65_PLUS"), " patients).")

# At CCG level demenetia diagnoses are broked down by age and sex

Recorded_dementia_ccg <- read_csv(url("https://files.digital.nhs.uk/5D/C18E54/dem-diag-ccg-quin-Aug-2018.csv"), col_types = cols( ach_date = col_character(),COMMISSIONER_ORGANISATION_CODE = col_character(),GEOGRAPHY_CODE = col_character(),CCG_NAME = col_character(),Measure = col_character(),Value = col_integer())) %>% 
  filter(COMMISSIONER_ORGANISATION_CODE %in% c("09G", "09H", "09X"),
         ach_date == "31AUG2018") %>% 
  group_by(Measure) %>% 
  summarise(Value = sum(Value, na.rm = TRUE)) %>% 
  mutate(Age = ifelse(grepl("65_69", Measure) == TRUE, "Age 65-69", ifelse(grepl("70_75", Measure) == TRUE | grepl("70_74", Measure) == TRUE, "Age 70-74", ifelse(grepl("75_79", Measure) == TRUE, "Age 75-79", ifelse(grepl("80_84", Measure) == TRUE, "Age 80-84", ifelse(grepl("85_89", Measure) == TRUE, "Age 85-89",  ifelse(grepl("90_PLUS", Measure) == TRUE, "Age 90+", NA)))))),
         Sex = ifelse(grepl("FEMALE_", Measure) == TRUE, "Female", ifelse(grepl("MALE_", Measure) == TRUE, "Male", ifelse(grepl("ALL_", Measure) == TRUE, "Persons", NA)))) %>% 
  select(Age, Sex, Value) %>% 
  spread(Sex, Value)

paste0("Not everyone with dementia has a formal diagnosis. The Dementia 65+ estimated diagnosis rate indicator, published by NHS Digital, compares the number of people thought to have dementia with the number of people diagnosed with dementia, aged 65 and over. The target is for at least two thirds of people with dementia to be diagnosed.")

paste0("Since 2012, the NHS has been seeking to ensure that patients suffering from dementia are given a formal diagnosis so they can receive appropriate care and support. The national target is for two thirds of people with dementia to be formally diagnosed.")

Dem_diagnosis_rate <- read_csv(url("https://files.digital.nhs.uk/A1/A06253/dem-diag-ind-nhs-Aug-2018.csv")) %>% 
  filter(ORG_CODE %in% c("09G", "09H", "09X")) %>% 
  group_by(NAME, ACH_DATE, MEASURE) %>% 
  filter(MEASURE %in% c("DEMENTIA_ESTIMATE_65_PLUS", "DEMENTIA_REGISTER_65_PLUS")) %>% 
  summarise(VALUE = sum(VALUE, na.rm = TRUE)) %>% 
  spread(MEASURE, VALUE) %>% 
  rename(Register = DEMENTIA_REGISTER_65_PLUS,
         Estimate = DEMENTIA_ESTIMATE_65_PLUS) %>% 
  mutate(Diagnosis_rate = Register/Estimate,
         Diagnosis_rate_lci = wilson_lower(Register, Estimate, 0.95),
         Diagnosis_rate_uci = wilson_upper(Register, Estimate, 0.95),
         Day = substr(ACH_DATE, 1,2),
         Month = capwords(substr(ACH_DATE, 3,5), strict = TRUE),
         Year = substr(ACH_DATE, 6,9)) %>% 
  mutate(Date = as.Date(paste(Year, Month, Day, sep = "-"), format = "%Y-%b-%d")) %>% 
  arrange(Date) %>% 
  mutate(Month_year = format(Date, "%B-%Y"))

Dem_diagnosis_rate$Month_year <- factor(Dem_diagnosis_rate$Month_year, levels = Dem_diagnosis_rate$Month_year)

Dem_diagnosis_rate$Target_met <- ifelse(Dem_diagnosis_rate$Diagnosis_rate >= .66, "Ambition achieved or exceeded", "Ambition not met")

d_ends <- Dem_diagnosis_rate %>% 
  group_by(NAME) %>% 
  top_n(1, Month_year) %>% 
  pull(Diagnosis_rate)

ggplot(Dem_diagnosis_rate, aes(x = Month_year, y = Diagnosis_rate, group = NAME, colour = NAME, fill = Target_met)) +
  labs(title = "Number of patients aged 65+ with a recorded diagnosis, as a proportion\nof the number of patients estimated to have dementia",
       subtitle = "West Sussex CCGs",
       caption = "The national ambition is for GP's to diagnose at least two thirds of the estimated population with dementia.\nThis is represented by the solid line across the figure at 66.6%.",
       x = "Month",
       y = "Diagnosis rate") +
  geom_line() +
  geom_errorbar(aes(ymin = Diagnosis_rate_lci, ymax = Diagnosis_rate_uci), 
                width = 0.25) +
  geom_hline(aes(yintercept = 0.66), col = "#474444", lwd = .6) +
  scale_y_continuous(limits = c(0.5,.8), breaks = seq(0,1,.1), labels = percent, sec.axis = sec_axis(~ ., breaks = d_ends, labels = percent)) +
  scale_colour_manual(values = c("#43a4cf","#f1d05d", "#7e267b"), name = "CCG") +
  scale_fill_manual(values = c("#f1a907","#1a4876"), name = "66.6% ambition") +
  QOF_prev_theme() +
  geom_point(size = 3, shape = 21, colour = "#000000", aes(fill = Target_met)) +
  theme(panel.grid.major.y = element_line(colour = "#dbdbdb"),
        panel.grid.major.x = element_blank(),
        axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.5),
        legend.box = "vertical",
        legend.direction = "horizontal",
        legend.spacing.y = unit(0, "cm")) 

paste0("Diagnosis rates in NHS Crawley CCG and NHS Horsham and Mid Sussex CCG are consistently above the 66.6% ambition target, whilst NHS Coastal West Sussex CCG diagnosis rates among those aged 65 and over are consistently lower than the 66.6% ambition.")

# Majority have ethnicity not defined
#Recorded_dementia_ccg_ethnicity <- read_csv("https://files.digital.nhs.uk/29/7806D0/rec-dem-ccg-Aug-2018-csv.csv")

# Projections for MM ####

# To follow the methodology we must apply assumptions to smallest possible geography and aggregate later

wsx_district_pop_MM <- read_csv("./Population/Population_time_series_0217.csv", col_types = cols(Area_name = col_character(),Area_code = col_character(), Age = col_integer(),  Sex = col_character(),Year = col_integer(),Population = col_integer(),Births = col_integer(),Deaths = col_integer(),Area_type = col_character())) %>% 
  mutate(Data_type = "Estimate",
         `Age group` = ifelse(Age <= 24, "0-24", ifelse(Age <= 44, "25-44", ifelse(Age <= 64, "45-64", ifelse(Age <= 84, "65-84", "85+"))))) %>% 
  group_by(Area_name, Area_code, Sex, Year, `Age group`, Data_type) %>% 
  filter(Area_name %in% c("Adur", "Arun", "Chichester", "Horsham", "Mid Sussex", "Worthing")) %>% 
  summarise(Population = sum(Population, na.rm = TRUE)) %>% 
  ungroup()

ONS_projection_1641_MM <- read_csv("./Population/2016 SNPP Population males.csv", col_types = cols(.default = col_double(),AREA_CODE = col_character(),AREA_NAME = col_character(),COMPONENT = col_character(),SEX = col_character(),AGE_GROUP = col_character())) %>% 
  bind_rows(read_csv("./Population/2016 SNPP Population females.csv", col_types = cols(.default = col_double(),AREA_CODE = col_character(),AREA_NAME = col_character(),COMPONENT = col_character(),SEX = col_character(),AGE_GROUP = col_character()))) %>% 
  filter(AREA_NAME %in% c("Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing") & AGE_GROUP != "All ages") %>% 
  mutate(Age = as.numeric(gsub(" and over", "", AGE_GROUP))) %>% 
  mutate(`Age group` = ifelse(Age <= 24, "0-24", ifelse(Age <= 44, "25-44", ifelse(Age <= 64, "45-64", ifelse(Age <= 84, "65-84", "85+"))))) %>% 
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
  mutate(Data_type = "Projected")

# This file was produced using the scenario file I:\DevCon\Sam Mason\Demographic Work\2015 projection back up\Popgroup\1. POPGROUP V4.0\PGSysfiles\2016 Projection_inp\\scenario_2016 - test.xls

WSCC_HH_projection_1641_MM <- data.frame(Area_name = "Adur", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Adur", skip = 4, n_max = 186), check.names = FALSE) %>% 
  bind_rows(data.frame(Area_name = "Arun", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Arun", skip = 4, n_max = 186), check.names = FALSE)) %>% 
  bind_rows(data.frame(Area_name = "Chichester", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Chichester", skip = 4, n_max = 186), check.names = FALSE)) %>% 
  bind_rows(data.frame(Area_name = "Crawley", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Crawley", skip = 4, n_max = 186), check.names = FALSE)) %>% 
  bind_rows(data.frame(Area_name = "Horsham", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Horsham", skip = 4, n_max = 186), check.names = FALSE)) %>% 
  bind_rows(data.frame(Area_name = "Mid Sussex", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Mid_Sussex", skip = 4, n_max = 186), check.names = FALSE)) %>% 
  bind_rows(data.frame(Area_name = "Worthing", read_excel("./Population/WSCC 2016 Projection - Single Year of Age.xls", sheet = "Worthing", skip = 4, n_max = 186), check.names = FALSE)) %>% 
  filter(Sex %in% c("male", "female")) %>% 
  mutate(Age = as.numeric(ifelse(Age == "90+", "90", Age)),
         Sex = capwords(Sex, strict = TRUE),
        `Age group` = ifelse(Age <= 24, "0-24", ifelse(Age <= 44, "25-44", ifelse(Age <= 64, "45-64", ifelse(Age <= 84, "65-84", "85+"))))) %>% 
  group_by(Area_name, Sex, `Age group`) %>% 
  summarise(`2012` = sum(`2012`, na.rm = TRUE),
            `2013` = sum(`2013`, na.rm = TRUE),
            `2014` = sum(`2014`, na.rm = TRUE),
            `2015` = sum(`2015`, na.rm = TRUE),
            `2016` = sum(`2016`, na.rm = TRUE),
            `2017` = sum(`2017`, na.rm = TRUE),
            `2018` = sum(`2018`, na.rm = TRUE),
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
            `2036` = sum(`2036`, na.rm = TRUE)) %>% 
  gather(Year, Population, `2012`:`2036`, factor_key = TRUE) %>% 
  filter(!(Year %in% c(2012,2013,2014,2015,2016,2017))) %>% 
  mutate(Data_type = "Projected with HH",
         Area_code = ifelse(Area_name == "Adur", "E07000223", ifelse(Area_name == "Arun", "E07000224", ifelse(Area_name == "Chichester", "E07000225", ifelse(Area_name == "Crawley", "E07000226", ifelse(Area_name == "Horsham", "E07000227", ifelse(Area_name == "Mid Sussex", "E07000228", ifelse(Area_name == "Worthing", "E07000229",NA)))))))) %>% 
  ungroup()

ONS_plus_estimates <- rbind(wsx_district_pop_MM, ONS_projection_1641_MM)
WSCC_HH_plus_estimates <- rbind(wsx_district_pop_MM, WSCC_HH_projection_1641_MM)

rm(wsx_district_pop_MM, ONS_projection_1641_MM, WSCC_HH_projection_1641_MM)

# Multi morbidity ####

paste0("Data analysis using registered patient populations in NHS Crawley CCG and NHS Horsham and Mid Sussex CCG identifed that 35.1% of those aged 65 years and over had 2 or more co-occuring long term conditions. Among 65-74 year olds, the prevalence was 25.8% which increased to 43% among 75-84 year olds and 49.6% among those aged 85 and over. Applying these percentages to the 2016 resident population estimates for West Sussex, suggests that just over 66,500 over 65's have two or more co-occuring long term conditions. It should be reiterated that these prevalence estimates are based on data from the two north CCGs, with no data available on patients in NHS Coastal West Sussex CCG (which comprises almost 60% of the population of West Sussex).")

paste0("Other prevalence estimates are available based on large scale observed prevalence data from academic studies which details age specific multimorbidity estimates by sex at local authority (lower tier) level. These multimorbidity estimates are an application of a scenario-based approach. This involves applying the observed prevalence in a number of studies identified through a literature review to the South East region and local populations within it, in order to build up a picture of multimorbidity under different scenarios. The estimates apply equal weights to all conditions and disorders which produces a count of the number of conditions present, even though the effect of multimorbidity on individuals will vary with the combination and severity of disorders.")

paste0("The study uses raw data from observed co-occurence of conditions in Scotland. However, whilst the prevalence of long-term conditions appeared to be more favourable in Scotland than in England the differences between both countries were small. PHE therefore consider Barnett and colleagues' observed multi-morbidity prevalence suitable for deriving the expected prevalence estimates for areas in England.")

# Source: South West Local Knowledge and Intelligence Service, Public Health England

paste0("Other prevalence estimates are available based on large scale observed prevalence data from academic studies which details age specific multimorbidity estimates by sex at local authority (lower tier) level. These multimorbidity estimates are an application of a scenario-based approach. This involves applying the observed prevalence in a number of studies identified through a literature review to the South East region and local populations within it, in order to build up a picture of multimorbidity under different scenarios. The estimates apply equal weights to all conditions and disorders which produces a count of the number of conditions present, even though the effect of multimorbidity on individuals will vary with the combination and severity of disorders.")

paste0("The study uses raw data from observed co-occurence of conditions in Scotland. However, whilst the prevalence of long-term conditions appeared to be more favourable in Scotland than in England the differences between both countries were small. PHE therefore consider Barnett and colleagues' observed multi-morbidity prevalence suitable for deriving the expected prevalence estimates for areas in England.")

# Source: South West Local Knowledge and Intelligence Service, Public Health England

# Although multi-morbidity (presence of multiple chronic (long-term) conditions) has been researched extensively, there is currently no consensus on its precise definition. The number, type (physical or mental health) and selection criteria for conditions included in multi-morbidity indices vary from one author to another. For example, while Salisbury et. al., (2011) defined multi-morbidity as the presence of more than one chronic condition, Barnett et. al., (2012) defined it as the presence of two or more. 

# In addition, both authors also differ in terms of their selection criteria. While the former focused on conditions listed in the Quality and Outcomes Framework (QOF) and a much wider list of chronic conditions identified using the Johns Hopkins University Adjusted Clinical Groups Case-Mix System, the latter considered diseases included in the QOF, conditions recommended in a systematic review and long-term disorders identified as important by NHS Scotland.		

paste("The differences in definitions and measurement tools give rise to non-comparable information on the prevalence of multi-morbidity across various studies.")

lookup <- read_csv(url("https://opendata.arcgis.com/datasets/41828627a5ae4f65961b0e741258d210_0.csv"), col_types = cols(LTLA17CD = col_character(),  LTLA17NM = col_character(),  UTLA17CD = col_character(),  UTLA17NM = col_character(),  FID = col_integer()))
# This is a lower tier LA to upper tier LA lookup
UA <- subset(lookup, LTLA17NM == UTLA17NM)

Conditions_MM <- read_excel("./Reference table - Sex and age-specific prevalence estimates of multi-morbidity by English region, local authority and .xlsx", sheet = "Definition", skip = 22, n_max = 40)

Original_ref <- "Barnett K, Mercer SW, Norbury M, Watt G, Wyke S, and Guthrie B. Epidemiology of multimorbidity and implications for health care, research, and medical education: a cross-sectional study. The Lancet. 2012;380(9836):37-43"

MM_2_LTLA <- read_excel("./Reference table - Sex and age-specific prevalence estimates of multi-morbidity by English region, local authority and .xlsx", sheet = "2+MM by sex, LA & age", skip = 4, n_max = 3260) %>% 
  rename(Sex = `Sex (where 1=Male & 2=Female)`,
         LTLA17CD = `Local authority code`) %>% 
  mutate(Sex = ifelse(Sex == 1, "Male", ifelse(Sex == 2, "Female", NA))) %>% 
  left_join(lookup[c("LTLA17CD", "UTLA17CD", "UTLA17NM")], by = "LTLA17CD")

MM_2_LTLA_wsx <- MM_2_LTLA %>% 
  filter(`Local authority name` %in% c("Adur", "Arun", "Chichester", "Crawley", "Horsham", "Mid Sussex", "Worthing")) %>% 
 mutate(`Age group` = factor(`Age group`, levels = c("85+", "65-84", "45-64", "25-44", "0-24")))

names(MM_2_LTLA_wsx)

MM_2_LTLA_range <- as.data.frame(MM_2_LTLA_wsx %>% 
         mutate(Estimate_lower_range = (Population/100)*`Lower 95% CI`, 
                Estimate_upper_range = (Population/100)*`Upper 95% CI`) %>% 
         group_by(`Local authority name`, LTLA17CD) %>%
         summarise(Estimate_applied = sum(`Number of people with 2+ MM`, na.rm = TRUE),
                Estimate_lower_range = sum(Estimate_lower_range, na.rm = TRUE),
                Estimate_upper_range = sum(Estimate_upper_range, na.rm = TRUE)) %>% 
         mutate(Range = paste0(format(round(Estimate_lower_range,0),big.mark = ","), "-\n", format(round(Estimate_upper_range,0),big.mark = ","))) %>% 
         ungroup())

MM_2_LTLA_range_65_plus <- as.data.frame(MM_2_LTLA_wsx %>% 
          filter(`Age group` %in% c("65-84", "85+")) %>% 
          mutate(Estimate_lower_range = (Population/100)*`Lower 95% CI`, 
                 Estimate_upper_range = (Population/100)*`Upper 95% CI`) %>% 
          group_by(`Local authority name`, LTLA17CD) %>%
          summarise(Estimate_applied = sum(`Number of people with 2+ MM`, na.rm = TRUE),
                    Estimate_lower_range = sum(Estimate_lower_range, na.rm = TRUE),
                    Estimate_upper_range = sum(Estimate_upper_range, na.rm = TRUE)) %>% 
          mutate(Range = paste0(format(round(Estimate_lower_range,0),big.mark = ","), "-\n", format(round(Estimate_upper_range,0),big.mark = ","))) %>% 
          ungroup())

ggplot(MM_2_LTLA_wsx) +
  geom_bar(aes(x = `Local authority name`, y = `Number of people with 2+ MM`, fill = `Age group`, colour = `Age group`), stat = "identity", width = .5) +
  scale_fill_manual(values = c("#331455","#d43a70","#f0a173","#f0c09a","#faffcd"), breaks = c("0-24", "25-44", "45-64", "65-84", "85+")) +
  scale_colour_manual(values = c("#331455","#d43a70","#f0a173","#f0c09a","#faffcd"), breaks = c("0-24", "25-44", "45-64", "65-84", "85+")) +
  labs(title = "Estimated number of people with two or more chronic conditions; ",
       subtitle = "West Sussex districts; 2011",
       x = "District") +
  scale_y_continuous(breaks = seq(0,50000,5000), limits = c(0,50000), labels = comma, expand = c(0,0.01)) +
  ph_theme() +
  theme(panel.background = element_rect(fill = "#d3d3d3"),
        plot.background = element_rect(fill = "#d3d3d3"),
        legend.background = element_rect(fill = "#d3d3d3"),
        axis.text.x = element_text(hjust =  0.5)) +
  geom_text(data = MM_2_LTLA_range, aes(x = `Local authority name`, y = Estimate_upper_range + 2400, label = unique(Range)), size = 2.75,  show.legend = FALSE)

ggplot(subset(MM_2_LTLA_wsx, `Age group` %in% c("65-84", "85+"))) +
  geom_bar(aes(x = `Local authority name`, y = `Number of people with 2+ MM`, fill = `Age group`, colour = `Age group`), stat = "identity", width = .5) +
  scale_fill_manual(values = c("#331455","#d43a70"), breaks = c("65-84", "85+")) +
  scale_colour_manual(values = c("#331455","#d43a70"), breaks = c("65-84", "85+")) +
  labs(title = "Estimated number of people with two or more chronic conditions; Those aged 65+",
       subtitle = "West Sussex districts; 2011",
       x = "District") +
  scale_y_continuous(breaks = seq(0,50000,5000), limits = c(0,50000), labels = comma, expand = c(0,0.01)) +
  ph_theme() +
  theme(panel.background = element_rect(fill = "#d3d3d3"),
        plot.background = element_rect(fill = "#d3d3d3"),
        legend.background = element_rect(fill = "#d3d3d3"),
        axis.text.x = element_text(hjust =  0.5)) +
  geom_text(data = MM_2_LTLA_range_65_plus, aes(x = `Local authority name`, y = Estimate_upper_range + 2400, label = unique(Range)), size = 2.75,  show.legend = FALSE)


# Upper tier

MM_2_UTLA <- MM_2_LTLA %>% 
  group_by(UTLA17CD, UTLA17NM, Sex, `Age group`) %>% 
  summarise(`Number of people with 2+ MM` = sum(`Number of people with 2+ MM`, na.rm = TRUE),
            Population = sum(Population, na.rm = TRUE)) %>% 
  mutate(`Prevalence of 2+ MM (%)` = (`Number of people with 2+ MM` / Population)*100,
         `Lower 95% CI` = wilson_lower(`Number of people with 2+ MM`,Population, .95)*100,
         `Upper 95% CI` = wilson_upper(`Number of people with 2+ MM`,Population, .95)*100)

MM_2_UTLA_Total <- MM_2_LTLA %>% 
  group_by(UTLA17CD, UTLA17NM, `Age group`) %>% 
  summarise(`Number of people with 2+ MM` = sum(`Number of people with 2+ MM`, na.rm = TRUE),
            Population = sum(Population, na.rm = TRUE)) %>% 
  mutate(Sex = "Persons",
         `Prevalence of 2+ MM (%)` = (`Number of people with 2+ MM` / Population)*100,
         `Lower 95% CI` = wilson_lower(`Number of people with 2+ MM`,Population, .95)*100,
         `Upper 95% CI` = wilson_upper(`Number of people with 2+ MM`,Population, .95)*100)

MM_2_UTLA <- rbind(MM_2_UTLA, MM_2_UTLA_Total)
rm(MM_2_UTLA_Total)

MM_2_UTLA_wsx <- MM_2_UTLA %>% 
  filter(UTLA17NM == "West Sussex") 

ggplot(MM_2_UTLA_wsx, aes(x = Sex, y = `Prevalence of 2+ MM (%)`, fill = Sex)) +
  geom_bar(stat = "identity") +
  facet_rep_wrap(~`Age group`, nrow = 1, repeat.tick.labels = TRUE)+
  geom_errorbar(data = MM_2_UTLA_wsx, aes(ymin = `Lower 95% CI`, ymax = `Upper 95% CI`), width = 0.2) +
  ph_theme() +
  scale_y_continuous(breaks = seq(0,100,10), limits = c(0,100), labels = unit_format("%")) +
  scale_fill_manual(values = c("#00C3FF", "#fd6400", "#172243")) +
  labs(title = "Estimated prevalence of multi-morbidity by age; West Sussex; 2011",
       subtitle = "Prevalence of two or more morbidities",
       x = "Age") +
  theme(axis.ticks.y = element_line(colour = "#000000"),
        axis.text.y = element_text(hjust = 0.5),
        axis.text.x = element_blank(),
        strip.text = element_text(colour = "#000000", face = "bold"),
        panel.spacing = unit(.2, "lines")) 

MM_2_UTLA_wsx_numb <- MM_2_UTLA_wsx
MM_2_UTLA_wsx_numb$`Age group` <- factor(MM_2_UTLA_wsx_numb$`Age group`, levels = c("85+", "65-84", "45-64", "25-44", "0-24"))

ggplot(data = subset(MM_2_UTLA_wsx_numb, Sex != "Persons"), aes(x = Sex, y = `Number of people with 2+ MM`, fill = `Age group`, colour = `Age group`)) +
  geom_bar(stat = "identity", width = .5) +
  scale_fill_manual(values = c("#331455","#d43a70","#f0a173","#f0c09a","#faffcd"), breaks = c("0-24", "25-44", "45-64", "65-84", "85+")) +
  scale_colour_manual(values = c("#331455","#d43a70","#f0a173","#f0c09a","#faffcd"), breaks = c("0-24", "25-44", "45-64", "65-84", "85+")) +
  labs(title = "Estimated number of people with two or more chronic conditions; ",
       subtitle = "West Sussex; 2011",
       x = "Age") +
  scale_y_continuous(breaks = seq(0,150000,25000), limits = c(0,150000), labels = comma) +
  ph_theme() +
  theme(panel.background = element_rect(fill = "#d3d3d3"),
        plot.background = element_rect(fill = "#d3d3d3"),
        legend.background = element_rect(fill = "#d3d3d3"))
  
MM_2_UTLA_wsx_table <- MM_2_UTLA %>% 
  filter(UTLA17NM == "West Sussex") %>% 
  # filter(`Age group` %in% c("65-84", "85+")) %>% 
  mutate(Number_label = round(`Number of people with 2+ MM`),
         Comorbidities = "2+") %>% 
  dcast(Sex + Comorbidities ~ `Age group`, value.var = "Number_label")

MM_2_UTLA_wsx_perc_table <- MM_2_UTLA %>% 
  filter(UTLA17NM == "West Sussex") %>% 
  # filter(`Age group` %in% c("65-84", "85+")) %>% 
  mutate(Comorbidities = "2+",
         Percent_label = paste0(round(`Prevalence of 2+ MM (%)`,1), "%,\n95% CI = ", round(`Lower 95% CI`,1),"-", round(`Upper 95% CI`,1),"%")) %>% 
  dcast(Sex + Comorbidities ~ `Age group`, value.var = "Percent_label")

# Projecting Two + MM 

wsx_pop_2_MM <- rbind(wsx_pop_MM, Project_2016_2041_MM) %>% 
  left_join(MM_2_UTLA_wsx[c("Sex","Age group","Prevalence of 2+ MM (%)","Lower 95% CI","Upper 95% CI")], by = c("Age group", "Sex")) %>% 
  mutate(Estimate_applied = (Population/100) * `Prevalence of 2+ MM (%)`,
         Estimate_lower_range = (Population/100) * `Lower 95% CI`,
         Estimate_upper_range = (Population/100) * `Upper 95% CI`) %>% 
  ungroup()

Total_range <- as.data.frame(wsx_pop_2_MM %>% 
  group_by(Year) %>% 
  summarise(Estimate_applied = sum(Estimate_applied, na.rm = TRUE),
            Estimate_lower_range = sum(Estimate_lower_range, na.rm = TRUE),
            Estimate_upper_range = sum(Estimate_upper_range, na.rm = TRUE)) %>% 
  mutate(Range = paste0(format(round(Estimate_lower_range,0),big.mark = ","), "-\n", format(round(Estimate_upper_range,0),big.mark = ","))) %>% 
  ungroup())

Range_65_plus <- as.data.frame(wsx_pop_MM %>% 
  filter(`Age group` %in% c("65-84", "85+")) %>% 
  group_by(Year) %>% 
  summarise(Estimate_applied = sum(Estimate_applied, na.rm = TRUE),
            Estimate_lower_range = sum(Estimate_lower_range, na.rm = TRUE),
            Estimate_upper_range = sum(Estimate_upper_range, na.rm = TRUE)) %>% 
  mutate(Range = paste0(format(round(Estimate_lower_range,0),big.mark = ","), "-\n", format(round(Estimate_upper_range,0),big.mark = ","))) %>% 
  ungroup())

wsx_pop_MM$`Age group` <- factor(wsx_pop_MM$`Age group`, levels = c("85+", "65-84", "45-64", "25-44", "0-24"))

ggplot(wsx_pop_MM) +
  geom_bar(aes(x = Year, y = Estimate_applied, fill = `Age group`, colour = `Age group`), stat = "identity", width = .5) +
  scale_fill_manual(values = c("#331455","#d43a70","#f0a173","#f0c09a","#faffcd"), breaks = c("0-24", "25-44", "45-64", "65-84", "85+")) +
  scale_colour_manual(values = c("#331455","#d43a70","#f0a173","#f0c09a","#faffcd"), breaks = c("0-24", "25-44", "45-64", "65-84", "85+")) +
  labs(title = "Estimated number of people with two or more chronic conditions;",
       subtitle = "West Sussex; 2011 estimates applied 2002-2041",
       caption = "The numbers above each bar show the lower and upper 95% confidence limits of the estimate.",
       x = "Year",
       y = "Population") +
  scale_y_continuous(breaks = seq(0,350000,25000), limits = c(0,350000), labels = comma, expand = c(0,.01)) +
  ph_theme() +
  theme(panel.background = element_rect(fill = "#d3d3d3"),
        plot.background = element_rect(fill = "#d3d3d3"),
        legend.background = element_rect(fill = "#d3d3d3"),
        axis.text.x = element_text(hjust = 0.5, angle = 90)) +
  geom_text(data = subset(Total_range, Year %in% c(2011,2017,2020,2025,2030,2035,2040)), aes(x = Year, y = Estimate_upper_range + 35000, label = unique(Year)), size = 3,  show.legend = FALSE, fontface = "bold") +
  geom_text(data = subset(Total_range, Year %in% c(2011,2017,2020,2025,2030,2035,2040)), aes(x = Year, y = Estimate_upper_range + 16000, label = unique(Range)), size = 2.5,  show.legend = FALSE)

ggplot(data = subset(wsx_pop_MM, `Age group` %in% c("65-84", "85+"))) +
  geom_bar(aes(x = Year, y = Estimate_applied, fill = `Age group`, colour = `Age group`), stat = "identity", width = .5) +
  scale_fill_manual(values = c("#331455","#d43a70"), breaks = c("65-84", "85+")) +
  scale_colour_manual(values = c("#331455","#d43a70"), breaks = c("65-84", "85+")) +
  labs(title = "Estimated number of people with two or more chronic conditions; Those aged 65+",
       subtitle = "West Sussex; 2011 estimates applied 2002-2041",
       caption = "The numbers above each bar show the lower and upper 95% confidence limits of the estimate.",
       x = "Year",
       y = "Population") +
  scale_y_continuous(breaks = seq(0,350000,25000), limits = c(0,350000), labels = comma, expand = c(0,.01)) +
  ph_theme() +
  theme(panel.background = element_rect(fill = "#d3d3d3"),
        plot.background = element_rect(fill = "#d3d3d3"),
        legend.background = element_rect(fill = "#d3d3d3"),
        axis.text.x = element_text(hjust = 0.5, angle = 90)) +
  geom_text(data = subset(Range_65_plus, Year %in% c(2011,2015,2018,2020,2025,2030,2035,2040)), aes(x = Year, y = Estimate_upper_range + 35000, label = unique(Year)), size = 3, show.legend = FALSE, fontface = "bold") +
  geom_text(data = subset(Range_65_plus, Year %in% c(2011,2015,2018,2020,2025,2030,2035,2040)), aes(x = Year, y = Estimate_upper_range + 16000, label = unique(Range)), size = 2.5, show.legend = FALSE)
  
paste0("Applying the estimated prevalence rates of two or more chronic morbidities to what we expect the population of West Sussex might look like in the future (based on a 'business as usual' scenario with no changes in health or quality of care), we might expect there to be approximately ", format(round(subset(Range_65_plus, Year == "2018", select = "Estimate_applied"),-2),big.mark = ","), " residents aged 65 years and over living with multi-morbidity in 2018. By 2025, this might increase to ", format(round(subset(Range_65_plus, Year == "2025", select = "Estimate_applied"),-2),big.mark = ","), " residents and by 2040, there may be more than 200,000 residents aged 65 and over living with two or more chronic conditions.")

# Three plus MM

MM_3_LTLA <- read_excel("./Reference table - Sex and age-specific prevalence estimates of multi-morbidity by English region, local authority and .xlsx", sheet = "3+MM by sex, LA & age", skip = 4, n_max = 3260) %>% 
  rename(Sex = `Sex (where 1=Male & 2=Female)`,
         LTLA17CD = `Local authority code`) %>% 
  mutate(Sex = ifelse(Sex == 1, "Male", ifelse(Sex == 2, "Female", NA))) %>% 
  left_join(lookup[c("LTLA17CD", "UTLA17CD", "UTLA17NM")], by = "LTLA17CD")



# Physical and mental health MM

Phys_Mental_MM_LTLA <- read_excel("./Reference table - Sex and age-specific prevalence estimates of multi-morbidity by English region, local authority and .xlsx", sheet = "Phy&Mental by sex, LA & age", skip = 4, n_max = 3260) %>% 
  rename(Sex = `Sex (where 1=Male & 2=Female)`,
         LTLA17CD = `Local authority code`) %>% 
  mutate(Sex = ifelse(Sex == 1, "Male", ifelse(Sex == 2, "Female", NA))) %>% 
  left_join(lookup[c("LTLA17CD", "UTLA17CD", "UTLA17NM")], by = "LTLA17CD")





Phys_Mental_MM_UTLA <- Phys_Mental_MM_LTLA %>% 
  group_by(UTLA17CD, UTLA17NM, Sex, `Age group`) %>% 
  summarise(`Number with Physical & Mental health comorbidity` = sum(`Number with Physical & Mental health comorbidity`, na.rm = TRUE),
            Population = sum(Population, na.rm = TRUE)) %>% 
  mutate(`Prevalence of Physical & Mental health comorbidity (%)` = round((`Number with Physical & Mental health comorbidity` / Population)*100,2),
         `Lower 95% CI` = round(wilson_lower(`Number with Physical & Mental health comorbidity`,Population, .95)*100,2),
         `Upper 95% CI` = round(wilson_upper(`Number with Physical & Mental health comorbidity`,Population, .95)*100,2))

Phys_Mental_MM_UTLA_Total <- Phys_Mental_MM_LTLA %>% 
  group_by(UTLA17CD, UTLA17NM, `Age group`) %>% 
  summarise(`Number with Physical & Mental health comorbidity` = sum(`Number with Physical & Mental health comorbidity`, na.rm = TRUE),
            Population = sum(Population, na.rm = TRUE)) %>% 
  mutate(Sex = "Persons",
         `Prevalence of Physical & Mental health comorbidity (%)` = round((`Number with Physical & Mental health comorbidity` / Population)*100,2),
         `Lower 95% CI` = round(wilson_lower(`Number with Physical & Mental health comorbidity`,Population, .95)*100,2),
         `Upper 95% CI` = round(wilson_upper(`Number with Physical & Mental health comorbidity`,Population, .95)*100,2))

Phys_Mental_MM_UTLA <- rbind(Phys_Mental_MM_UTLA, Phys_Mental_MM_UTLA_Total)
rm(Phys_Mental_MM_UTLA_Total)




Phys_Men_UTLA_wsx_table <- Phys_Mental_MM_UTLA %>% 
  filter(UTLA17NM == "West Sussex") %>% 
  # filter(`Age group` %in% c("65-84", "85+")) %>% 
  mutate(Number_label = round(`Number with Physical & Mental health comorbidity`),
         Percent_label = paste0(round(`Prevalence of Physical & Mental health comorbidity (%)`,1), "%,\n95% CI = ", round(`Lower 95% CI`,1),"-", round(`Upper 95% CI`,1),"%"),
         Comorbidities = "Physical & Mental health") %>% 
  dcast(Sex + Comorbidities ~ `Age group`, value.var = "Number_label")

MM_number_table_wsx <- rbind(MM_2_UTLA_wsx_table, MM_3_UTLA_wsx_table, Phys_Men_UTLA_wsx_table)
rm(MM_2_UTLA_wsx_table, MM_3_UTLA_wsx_table, Phys_Men_UTLA_wsx_table)

MM_2_UTLA_wsx_table <- MM_2_UTLA %>% 
  filter(UTLA17NM == "West Sussex") %>% 
  #  filter(`Age group` %in% c("65-84", "85+")) %>% 
  mutate(Number_label = round(`Number of people with 2+ MM`),
         Percent_label = paste0(round(`Prevalence of 2+ MM (%)`,1), "%,\n95% CI = ", round(`Lower 95% CI`,1),"-", round(`Upper 95% CI`,1),"%"),
         Comorbidities = "2+") %>% 
  dcast(Sex + Comorbidities ~ `Age group`, value.var = "Percent_label")

MM_3_UTLA_wsx_table <- MM_3_UTLA %>% 
  filter(UTLA17NM == "West Sussex") %>% 
  #filter(`Age group` %in% c("65-84", "85+")) %>% 
  mutate(Number_label = round(`Number of people with 3+ MM`),
         Percent_label = paste0(round(`Prevalence of 3+ MM (%)`,1), "%,\n95% CI = ", round(`Lower 95% CI`,1),"-", round(`Upper 95% CI`,1),"%"),
         Comorbidities = "3+") %>% 
  dcast(Sex + Comorbidities ~ `Age group`, value.var = "Percent_label")

Phys_Men_UTLA_wsx_table <- Phys_Mental_MM_UTLA %>% 
  filter(UTLA17NM == "West Sussex") %>% 
  # filter(`Age group` %in% c("65-84", "85+")) %>% 
  mutate(Number_label = round(`Number with Physical & Mental health comorbidity`),
         Percent_label = paste0(round(`Prevalence of Physical & Mental health comorbidity (%)`,1), "%,\n95% CI = ", round(`Lower 95% CI`,1),"-", round(`Upper 95% CI`,1),"%"),
         Comorbidities = "Physical & Mental health") %>% 
  dcast(Sex + Comorbidities ~ `Age group`, value.var = "Percent_label")

MM_table_wsx <- rbind(MM_2_UTLA_wsx_table, MM_3_UTLA_wsx_table, Phys_Men_UTLA_wsx_table)
rm(MM_2_UTLA_wsx_table, MM_3_UTLA_wsx_table, Phys_Men_UTLA_wsx_table)

write.csv(MM_number_table_wsx, "./MM_number.csv")
write.csv(MM_table_wsx, "./MM_tanle.csv")

paste("The estimates suggest that the prevalence of multimorbidities is significantly higher among females compared to males in every age group. In terms of absolute numbers of patients in the very old age group (85+ years), there are more than twice the number of females with 2+ and 3+ co-occuring conditions and more than three times the number of patients with a co-occuring physical and mental health condition.")


# MM_north_CCGs <- MM_2_LTLA %>% 
#   filter(`Local authority name` %in% c("Horsham", "Mid Sussex", "Crawley")) %>% 
#   group_by(`Age group`) %>% 
#   summarise(`Number of people with 2+ MM` = sum(`Number of people with 2+ MM`, na.rm = TRUE),
#             Population = sum(Population, na.rm = TRUE)) %>% 
#   mutate(Sex = "Persons",
#          `Prevalence of 2+ MM (%)` = round((`Number of people with 2+ MM` / Population)*100,2),
#          `Lower 95% CI` = round(wilson_lower(`Number of people with 2+ MM`,Population, .95)*100,2),
#          `Upper 95% CI` = round(wilson_upper(`Number of people with 2+ MM`,Population, .95)*100,2))

paste0("The scenario based estimates are applied to 2011 mid year estimates. We could apply the percentage estimates to more recent mid year estimates or future projections.")





# The study (https://academic.oup.com/ageing/article/47/3/374/4815738) aims o project the future burden of multi-morbidity among older adults using stochastic microsimulation of younger cohorts in England from a 1% random sample of the Understanding Society, English Longitudinal Study of Ageing and the Cognitive Function and Ageing Study II.

# This analysis takes the prevalence assumptions from this study and applies it to the population projections for West Sussex.

# The limitations of this are that it assumes the population of west sussex will be the same as the national population used as a sample in the research paper. Also whilst the populations take into account recent local housing stock decisions, it does not take into future policy decisions or decisions that have not taken affect (including Brexit).



# Population and projections ####

population_df <- read_csv("./Population/Population_time_series_0217.csv", col_types = cols(Area_name = col_character(),Area_code = col_character(), Age = col_integer(),  Sex = col_character(),Year = col_integer(),Population = col_integer(),Births = col_integer(),Deaths = col_integer(),Area_type = col_character()))

population_df$Age_band <- ifelse(population_df$Age <= 14, "0-14", ifelse(population_df$Age <= 29, "15-29", ifelse(population_df$Age <= 44, "30-44", ifelse(population_df$Age <= 64, "45-64", ifelse(population_df$Age <= 79, "65-79", "80+")))))

population_df$Age_band <- factor(population_df$Age_band, levels = c("0-14", "15-29", "30-44", "45-64", "65-79", "80+", "Total"))

wsx_pop_band <- population_df %>% 
  group_by(Area_name,Area_code,Sex, Year,Age_band) %>% 
  filter(Area_name == "West Sussex" & Year %in% c("2017")) %>% 
  summarise(Population = sum(Population, na.rm = TRUE)) %>% 
  dcast(Age_band ~ Sex, value.var = "Population") %>% 
  mutate(Total = Female + Male) %>% 
  mutate(`Percentage in age group (total)` = round((Total / sum(Total))*100,1))


# Emergency admissions West Sussex patients aged 65+ ####

# Local authorities

wsx_district_pop_Emerg <- read_csv("./Population/Population_time_series_0217.csv", col_types = cols(Area_name = col_character(),Area_code = col_character(), Age = col_integer(),  Sex = col_character(),Year = col_integer(),Population = col_integer(),Births = col_integer(),Deaths = col_integer(),Area_type = col_character())) %>% 
 mutate(Data_type = "Estimate",
        `Age group` = ifelse(Age <= 4, "Age 0-4",  ifelse(Age <= 9, "Age 5-9",  ifelse(Age <= 14, "Age 10-14",  ifelse(Age <= 19, "Age 15-19", ifelse(Age <= 24, "Age 20-24", ifelse(Age <= 29, "Age 25-29", ifelse(Age <= 34, "Age 30-34", ifelse(Age <= 39, "Age 35-39", ifelse(Age <= 44, "Age 40-44", ifelse(Age <= 49, "Age 45-49", ifelse(Age <= 54, "Age 50-54", ifelse(Age <= 59, "Age 55-59", ifelse(Age <= 64, "Age 60-64", ifelse(Age <= 69, "Age 65-69", ifelse(Age <= 74, "Age 70-74", ifelse(Age <= 79, "Age 75-79", ifelse(Age <= 84, "Age 80-84", ifelse(Age <= 89, "Age 85-89", "Age 90+"))))))))))))))))))) %>% 
  group_by(Area_name, Area_code, Year, `Age group`, Data_type) %>% 
  filter(Area_name %in% c("Adur", "Arun", "Chichester","Crawley", "Horsham", "Mid Sussex", "Worthing")) %>% 
  summarise(Population = sum(Population, na.rm = TRUE)) %>% 
  filter(Year %in% c("2015", "2016", "2017")) %>% 
  mutate(FYEAR = ifelse(Year == "2015", "1516", ifelse(Year == "2016", "1617", ifelse(Year == "2017", "1718", NA)))) %>% 
  ungroup()

# CCG of responsibility is GP first then residence second (if no GP)
# Emerg_admissions_age_CCG_responsibility <- read_csv("//typhon/groups2.bu/Public Health Directorate/PH Research Unit/HDIS/HDIS EXTRACTS (VMWARE)/Emergency Admissions 65+/WORK_Aggregated.csv", col_types = cols(FYEAR = col_integer(),CCG_RESPONSIBILITY = col_character(),QuinaryAges = col_character(),COUNT_of_FAE_EMERGENCY = col_integer()))
# 
# Emerg_admissions_age_CCG_responsibility$QuinaryAges <- gsub(" to ", "-", Emerg_admissions_age_CCG_responsibility$QuinaryAges)
# 
# Emerg_admissions_age_CCG_responsibility$QuinaryAges <- factor(Emerg_admissions_age_CCG_responsibility$QuinaryAges, levels = c("Age 0-4","Age 5-9","Age 10-14","Age 15-19","Age 20-24","Age 25-29","Age 30-34","Age 35-39","Age 40-44","Age 45-49","Age 50-54","Age 55-59","Age 60-64","Age 65-69","Age 70-74","Age 75-79","Age 80-84","Age 85-89","Age 90+"))

# Emerg_admissions_age_LAD_res <- read_csv("//typhon/groups2.bu/Public Health Directorate/PH Research Unit/HDIS/HDIS EXTRACTS (VMWARE)/Emergency Admissions 65+/WORK_AggregatedRESLADST.csv", col_types = cols(FYEAR = col_character(),RESLADST_ONS = col_character(),QuinaryAges = col_character(),COUNT_of_FAE_EMERGENCY = col_integer())) %>% 
#   rename(Area_code = RESLADST_ONS,
#          `Age group` = QuinaryAges) %>% 
#   mutate(`Age group` = gsub(" to ", "-", `Age group`))# %>% 

paste0("The European Standard Population (ESP) is an artificial population structure which is used in the weighting of mortality or incidence data to produce age standardised rates . Eurostat, the statistical institute of the European Union, has decided to bring this population structure up to date. The ESP is used to calculate age standardised rates (ASRs), which allow comparison of rates, including those across populations that may have different age and sex structures.")

Emerg_admissions_age_LAD_res <- read_csv("./WORK_AggregatedRESLADST.csv", col_types = cols(FYEAR = col_character(),RESLADST_ONS = col_character(),QuinaryAges = col_character(),COUNT_of_FAE_EMERGENCY = col_integer())) %>% 
  rename(Area_code = RESLADST_ONS,
         `Age group` = QuinaryAges) %>% 
  mutate(`Age group` = gsub(" to ", "-", `Age group`)) %>% 
  left_join(wsx_district_pop_Emerg, by = c("FYEAR","Age group", "Area_code")) %>% 
  left_join(data.frame(`Age group` = c("Age 0-4","Age 5-9","Age 10-14","Age 15-19","Age 20-24","Age 25-29","Age 30-34","Age 35-39","Age 40-44","Age 45-49","Age 50-54","Age 55-59","Age 60-64","Age 65-69","Age 70-74","Age 75-79","Age 80-84","Age 85-89","Age 90+"), ESP_2013 = c(5000,5500,5500,5500,6000,6000,6500,7000,7000,7000,7000,6500,6000,5500,5000,4000,2500,1500,1000), check.names = FALSE), by = "Age group") %>% 
  mutate(`Age group` = factor(`Age group`, levels = c("Age 0-4","Age 5-9","Age 10-14","Age 15-19","Age 20-24","Age 25-29","Age 30-34","Age 35-39","Age 40-44","Age 45-49","Age 50-54","Age 55-59","Age 60-64","Age 65-69","Age 70-74","Age 75-79","Age 80-84","Age 85-89","Age 90+"))) %>% 
  select(Area_code, Area_name, FYEAR, Year, `Age group`, COUNT_of_FAE_EMERGENCY, Population, ESP_2013) %>% 
  rename(Emergency_admissions = COUNT_of_FAE_EMERGENCY,
         FYear = FYEAR) %>% 
  arrange(Area_name, FYear, `Age group`)

Emerg_admissions_age_West_Sussex_res <- Emerg_admissions_age_LAD_res %>% 
  group_by(FYear, Year, `Age group`) %>% 
  summarise(Emergency_admissions = sum(Emergency_admissions, na.rm = TRUE),
            ESP_2013 = unique(ESP_2013),
            Population = sum(Population)) %>% 
  mutate(Area_code = "E10000032",
         Area_name = "West Sussex") %>% 
  select(Area_code, Area_name, FYear, Year, `Age group`, Emergency_admissions, Population, ESP_2013)

Emerg_admissions_all_geographies <- Emerg_admissions_age_West_Sussex_res %>% 
  bind_rows(Emerg_admissions_age_LAD_res) %>% 
  mutate(Age_specific_rate_per_1000 = (Emergency_admissions / Population)*1000,
         Age_specific_rate_per_100000 = (Emergency_admissions / Population) * 100000) %>% 
  mutate(Age_standardised_rate = Age_specific_rate_per_100000 * ESP_2013) 

DSR_Emerg_admiss_area_year <- Emerg_admissions_all_geographies %>% 
  group_by(Area_code, Area_name, FYear, Year) %>% 
  summarise(Emergency_admissions = sum(Emergency_admissions, na.rm = TRUE),
            Population = sum(Population, na.rm = TRUE),
            ESP_2013 = sum(ESP_2013, na.rm = TRUE), 
            Age_standardised_rate = sum(Age_standardised_rate, na.rm = TRUE)) %>% 
  mutate(Rate_per_100000 = Age_standardised_rate / ESP_2013) %>% 
  mutate(Rate_per_100000_LCI = Rate_per_100000 - (1.96*(Rate_per_100000/sqrt(Emergency_admissions))),
         Rate_per_100000_UCI = Rate_per_100000 + (1.96*(Rate_per_100000/sqrt(Emergency_admissions))))

DSR_65_plus_Emerg_admiss_area_year <- Emerg_admissions_all_geographies %>%
  filter(`Age group` %in% c("Age 65-69", "Age 70-74", "Age 75-79", "Age 80-84", "Age 85-89", "Age 90+")) %>% 
  group_by(Area_code, Area_name, FYear, Year) %>% 
  summarise(Emergency_admissions = sum(Emergency_admissions, na.rm = TRUE),
            Population = sum(Population, na.rm = TRUE),
            ESP_2013 = sum(ESP_2013, na.rm = TRUE), 
            Age_standardised_rate = sum(Age_standardised_rate, na.rm = TRUE)) %>% 
  mutate(Rate_per_100000 = Age_standardised_rate / ESP_2013) %>% 
  mutate(Rate_per_100000_LCI = Rate_per_100000 - (1.96*(Rate_per_100000/sqrt(Emergency_admissions))),
         Rate_per_100000_UCI = Rate_per_100000 + (1.96*(Rate_per_100000/sqrt(Emergency_admissions))))

WSx_DSR_65_plus <- as.data.frame(DSR_65_plus_Emerg_admiss_area_year %>% 
  ungroup() %>% 
  select(Area_name, FYear, Emergency_admissions, Population, Rate_per_100000, Rate_per_100000_LCI, Rate_per_100000_UCI) %>% 
  filter(Area_name == "West Sussex") %>% 
  rename(Rate_per_100000_wsx = Rate_per_100000,
         Rate_per_100000_LCI_wsx = Rate_per_100000_LCI,
         Rate_per_100000_UCI_wsx = Rate_per_100000_UCI,
         Area_name_1 = Area_name) %>% 
  mutate(Year_on_year_trend = ifelse(Rate_per_100000_LCI_wsx > lag(Rate_per_100000_UCI_wsx), "increased (become worse)", ifelse(Rate_per_100000_UCI_wsx < lag(Rate_per_100000_LCI_wsx), "decreased (become better)", "not changed significantly")))) # lag does not work with group_by


paste0("In 2017/18* there were ", format(subset(WSx_DSR_65_plus, FYear == "1718", select = "Emergency_admissions"),big.mark = ","), " emergency admissions among residents aged 65 years and over in West Sussex. This represents an age standardised rate of ", format(round(subset(WSx_DSR_65_plus, FYear == "1718", select = "Rate_per_100000_wsx"),0),big.mark = ","), " admissions per 100,000 population aged 65+ (95% CI: ",format(round(subset(WSx_DSR_65_plus, FYear == "1718", select = "Rate_per_100000_LCI_wsx"),0),big.mark = ","),"-", format(round(subset(WSx_DSR_65_plus, FYear == "1718", select = "Rate_per_100000_UCI_wsx"),0),big.mark = ","), " admissions). Compared to 2017/18, the number of admissions in this age group has ", ifelse(subset(WSx_DSR_65_plus, FYear == "1718", select = "Emergency_admissions") > subset(WSx_DSR_65_plus, FYear == "1617", select = "Emergency_admissions"), "increased", ifelse(subset(WSx_DSR_65_plus, FYear == "1718", select = "Emergency_admissions") < subset(WSx_DSR_65_plus, FYear == "1617", select = "Emergency_admissions"), "decreased", "stayed the same")), " by ", abs(subset(WSx_DSR_65_plus, FYear == "1718", select = "Emergency_admissions") -subset(WSx_DSR_65_plus, FYear == "1617", select = "Emergency_admissions")), " admissions; the standardised rate has ", subset(WSx_DSR_65_plus, FYear == "1718", select = "Year_on_year_trend"),".")

DSR_wsx_districts <- DSR_65_plus_Emerg_admiss_area_year %>% 
  filter(Area_name != "West Sussex") %>% 
  left_join(WSx_DSR_65_plus[c("FYear","Rate_per_100000_wsx", "Rate_per_100000_LCI_wsx","Rate_per_100000_UCI_wsx")], by = "FYear") %>% 
  mutate(Compared_to_wsx = factor(ifelse(Rate_per_100000_LCI > Rate_per_100000_UCI_wsx, "Significantly higher", ifelse(Rate_per_100000_UCI < Rate_per_100000_LCI_wsx, "Significantly lower", "Similar")), levels = c("Significantly lower", "Similar", "Significantly higher")))

paste0("* 2017/18 data should be considered as provisional at the time of this report.")

paste0("The figure below shows the age standardised admission rate per 100,000 for West Sussex districts compared to the total in West Sussex for years 2015/16 to 2017/18.")


ggplot() +
  geom_line(data = WSx_DSR_65_plus, aes(x = FYear, y = Rate_per_100000_wsx), colour = "#7e8084", group = "West Sussex") +
  geom_errorbar(data = WSx_DSR_65_plus, aes(x = FYear, ymin = Rate_per_100000_LCI_wsx, ymax = Rate_per_100000_UCI_wsx), colour = "#7e8084", width = 0.1) +
  geom_point(data = WSx_DSR_65_plus, aes(x = FYear, y = Rate_per_100000_wsx), size = 2, shape = 21, colour = "#7e8084", fill = "#7e8084") +
  labs(title = "Trends in all cause emergency admissions aged 65+ per 100,000 ESP 2013;",
       subtitle = "West Sussex districts; 2015/16-2017/16",
       x = "Year",
       y = "Directly age-standardised rate\nper 100,000",
       caption = "Note: This is the number of admissions per 100,000, and not the number of people admitted.\nSome people will have multiple admissions throughout the year.") +
  geom_line(data = DSR_wsx_districts, aes(x = FYear, y = Rate_per_100000, group = Area_name, colour = Area_name)) +
  geom_errorbar(data = DSR_wsx_districts, aes(x = FYear, ymin = Rate_per_100000_LCI, ymax = Rate_per_100000_UCI), width = 0.1) +
  geom_point(data = DSR_wsx_districts, aes(x = FYear, y = Rate_per_100000, group = Area_name, fill = Compared_to_wsx), size = 3, shape = 21, colour = "#000000") +
  scale_y_continuous(limits = c(15000,30000), breaks = seq(0,30000,2500), labels = comma) +
  scale_fill_manual(values = c("#5555E6","#f1a907","#1a4876"), name = "District compared\ntoWest Sussex") +
  ph_theme() +
  theme(panel.grid.major.y = element_line(colour = "#dbdbdb"),
        panel.grid.major.x = element_blank(),
        axis.text.x = element_text(hjust = .5, vjust = 0.5),
        legend.box = "vertical",
        legend.direction = "horizontal",
        legend.spacing.y = unit(0, "cm")) +
  facet_rep_wrap( ~ Area_name, nrow = 3, repeat.tick.labels = TRUE)



Emerg_admissions_age_LAD_res %>% 
  group_by(FYear) %>% 
  summarise(Emergency_admissions = sum(Emergency_admissions))


# LCI = DSR_per_100000-1.96*(DSR_per_100000/SQRT(Admissions))
# UCI = DSR_per_100000+1.96*(DSR_per_100000/SQRT(Admissions))

# Try to do this for 0-4s at wsx level.

#table(Emerg_admissions_age$QuinaryAges, Emerg_admissions_age$FYEAR, Emerg_admissions_age$CCG_RESPONSIBILITY)

# ESP_2013_04_90plus <- data.frame(Age_group = c("Age 0-4","Age 5-9","Age 10-14","Age 15-19","Age 20-24","Age 25-29","Age 30-34","Age 35-39","Age 40-44","Age 45-49","Age 50-54","Age 55-59","Age 60-64","Age 65-69","Age 70-74","Age 75-79","Age 80-84","Age 85-89","Age 90+"), Population = c(5000,5500,5500,5500,6000,6000,6500,7000,7000,7000,7000,6500,6000,5500,5000,4000,2500,1500,1000))
# 
# ESP_2013_04_90plus$Age_group <- factor(ESP_2013_04_90plus$Age_group, levels = c("Age 0-4","Age 5-9","Age 10-14","Age 15-19","Age 20-24","Age 25-29","Age 30-34","Age 35-39","Age 40-44","Age 45-49","Age 50-54","Age 55-59","Age 60-64","Age 65-69","Age 70-74","Age 75-79","Age 80-84","Age 85-89","Age 90+"))

# Denominator  Local Authority estimates of resident population, Office for National Statistics (ONS)
# 
# Unrounded mid year population estimates produced by ONS and supplied to Public Health England. Analysis uses the quinary age bands 80-84, 85-89 and 90+.

# Hospital admissions amenable to healthcare ####
# Number of people ending life in hospital and usual place of residence ####
