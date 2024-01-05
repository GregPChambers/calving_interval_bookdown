###########################################################################################################################################
###########################################################################################################################################

# LA2046 CALVING INTERVAL ANALYSIS

# ANALYSIS OUTSIDE OF BOOKDOWN

###########################################################################################################################################
###########################################################################################################################################

# SETUP

library(plyr)
library(reshape)
library(ggplot2)
library(reshape2)
library(lubridate)
library(car)
library(dplyr)
library(extrafont)
library(arsenal)
library(gt)
library(gtsummary)
#library(data.table)
#library(gpairs)
#library(corrplot)
#library(survival)
#library(lmtest)
#library(scales)
#library(survminer)
library(gtools)
library(stringr)
library(forcats)
library(openxlsx)
library(flextable)
library(readxl)
library(english)
library(arsenal)
library(tidyr)
#library(ggpattern)
#library(magrittr)
#library(HLMdiag)
#library(ROCR)
#library(OptimalCutpoints)
#library(splines)
library(tidytext)
library(janitor)
library(purrr)
library(anytime)


# Set thousands separator for inline text
knitr::knit_hooks$set(inline = function(x) {   if(!is.numeric(x)){     x   }else{    prettyNum(round(x,2), big.mark=",")    } })

# ggplot theme
theme <- theme_bw(base_size = 11, base_family = "arial") +
  theme(legend.position = "top") +
  theme(
    #axis.title.x = element_text(colour="black", size=11),
        #axis.text.y  = element_text(size=11),
        #axis.text.x  = element_text(size=11, colour = "black"),
        #axis.title.y = element_text(colour="black", size=11),
        panel.grid.minor=element_blank(),
        panel.grid.major=element_blank(),
        panel.grid.major.y = element_line()) +
  #theme(legend.title = element_text(size=11), legend.text = element_text(size=11)) + 
  theme(legend.key.size = unit(0.5,"line")) #+ 
  #theme(strip.text.x = element_text(size=11, colour = "black"))

# Set flextable default settings
set_flextable_defaults(font.size = 10,
                       font.family = "cambria",
                       color = "black",
                       align = "left",
                       padding = 2)

mixedrank <- function(x) order(gtools::mixedorder(x))

############################################################################################################################################

# Data importation and cleaning

season_19_20 <- read_excel(path = "REPORT2019.xlsx")
season_20_21 <- read_excel(path = "REPORT2020.xlsx")
season_21_22 <- read_excel(path = "REPORT2021.xlsx")
season_22_23 <- read_excel(path = "REPORT2022.xlsx")

alldata <- rbind(season_19_20, season_20_21, season_21_22, season_22_23)

dat <- alldata %>%
  clean_names() %>%
  mutate(sid = paste(animal_id, client_id, sep="_")) %>%
  relocate(sid)
  
length(unique(dat$sid))
length(unique(dat$client_id))

table(dat$sex, useNA = "always")
table(dat$is_active)
sum(is.na(dat$calving_date))
table(dat[is.na(dat$calving_date), ]$sex)

# Sort out multiple calving dates issue
dat <- dat %>%
  mutate(multiple_dates = str_detect(calving_date, ";")) %>%
  mutate(calving_date = gsub(" ", "", calving_date, fixed = TRUE)) %>%
  mutate(different_dates = ifelse(multiple_dates==TRUE, map_int(str_split(calving_date, ";"), ~length(unique(.))), 1))

table(dat$multiple_dates)
table(dat$different_dates)
sum(dat$different_dates>1, na.rm = T)


# Filter out males and multiple unique calving dates, then tidy up calving date column
dat <- dat %>%
  filter(sex=="F") %>%
  filter(different_dates==1) %>%
  mutate(calving_date2 = if_else(str_detect(calving_date, ";"),
                                 as.Date(str_extract(str_replace(calving_date, "\\s+", ""), "^[^;]+"), format = "%Y-%m-%d"),
                                 as.Date(as.numeric(calving_date), origin = "1899-12-30")))




  select(1:7) %>%
  set_colnames(c("date", "region", "farm", "sn", "tag", "age", "se_blood")) %>%
  mutate(sn = as.factor(sn)) %>%
  mutate(date = ymd(date)) %>%
  mutate(age = as.numeric(age)) %>%
  mutate(sid = paste(sn, tag, sep="_")) %>%
  arrange(sn, as.numeric(tag)) %>%
  mutate(region = ifelse(region=="Oto/TA", "Otorohanga/Te Awamutu", ifelse(region=="Totally Vets", "Manawatu", ifelse(region=="Te Puke", "Bay of Plenty", region)))) %>%
  mutate(region = fct_relevel(region, "Bay of Plenty", "Otorohanga/Te Awamutu", "Putaruru", "Taranaki", "Manawatu", "Canterbury", "Otago")) %>%
  mutate(visit=1) %>%
  group_by(sn) %>%
  mutate(mean_se_blood = mean(se_blood)) %>%
  ungroup()

blood <- read_excel(path = "Data Entry - Bloods 29.11.23.xlsx", sheet = 1) %>%
  select(1:7, 9) %>%
  set_colnames(c("date", "region", "farm", "sn", "tag", "age", "se_blood", "notes")) %>%
  mutate(visit = ifelse(!is.na(notes) & notes=="2nd test", 2, 1)) %>%
  mutate(sn = as.factor(sn)) %>%
  mutate(date = ymd(date)) %>%
  mutate(age = as.numeric(age)) %>%
  mutate(sid = paste(sn, tag, sep="_")) %>%
  arrange(sn, as.numeric(tag)) %>%
  mutate(region = ifelse(region=="Oto/TA", "Otorohanga/Te Awamutu", ifelse(region=="Totally Vets", "Manawatu", ifelse(region=="Te Puke", "Bay of Plenty", region)))) %>%
  mutate(region = fct_relevel(region, "Bay of Plenty", "Otorohanga/Te Awamutu", "Putaruru", "Taranaki", "Manawatu", "Canterbury", "Otago")) %>%
  arrange(sn) %>%
  group_by(sn) %>%
  mutate(mean_se_blood = mean(se_blood)) %>%
  ungroup()

# Filter out first visit from the 5 farms
bld <- blood %>% 
  group_by(sn) %>%
  slice_max(visit) %>%
  ungroup()  %>%
  select(-notes)


# Bulk milk data (original dataset and then updated after replacing visit 1 with visit 2 data for the 5 farms)
old_bm_data <- read_excel(path = "Analytica - LA2032 - Milk Sample reults - Excel-R00.xlsx", sheet = 4) %>%
  select(5,6,9) %>%
  set_colnames(c("sn", "date", "se_bm")) %>%
  mutate(sn = as.factor(sn)) %>%
  mutate(date = ymd(date)) %>%
  filter(!(sn=="36843" & se_bm==0.033)) %>%
  filter(!(sn=="74570" & se_bm==0.011)) %>%
  mutate(visit=1)

bulk_milk <- read_excel(path = "Data Entry - Milk 29.11.23.xlsx", sheet = 1) %>%
  select(-3) %>%
  set_colnames(c("sn", "date", "se_bm")) %>%
  mutate(sn = as.factor(sn)) %>%
  mutate(date = ymd(date)) %>%
  mutate(se_bm = ifelse(se_bm=="<0.0025", "0", se_bm)) %>%
  mutate(se_bm = as.numeric(se_bm)) %>%
  arrange(sn) %>%
  filter(!(sn=="36843" & date==dmy("26/09/2023") & se_bm==0.033)) %>%
  filter(!(sn=="74570" & date==dmy("02/10/2023") & se_bm==0.011)) %>%
  filter(sn %in% bld$sn) %>% # Restrict to the 60 farms with blood data
  group_by(sn) %>%
  mutate(visit = seq_along(sn)) %>%
  ungroup()

# Filter out first visit from the 5 farms
bm <- bulk_milk %>%
  filter(!(visit==1 & sn %in% repeats), !(visit==2 & sn %nin% repeats))

dat <- merge(bld, bm, by="sn") %>%
  rename(blood_dt = "date.x", bm_dt = "date.y", visit_bld = "visit.x", visit_bm = "visit.y") %>%
  mutate(diff_dt = bm_dt - blood_dt) %>%
  mutate(se_bm_scaled = se_bm*1000) %>%
  mutate(cp_250 = ifelse(se_blood >250, 1, 0)) %>%
  mutate(cp_500 = ifelse(se_blood >500, 1, 0)) %>%
  mutate(cp_1000 = ifelse(se_blood >1000, 1, 0)) %>%
  mutate(cp_2000 = ifelse(se_blood >2000, 1, 0)) %>%
  mutate(cp_3000 = ifelse(se_blood >3000, 1, 0)) %>%
  mutate(cp_3500 = ifelse(se_blood >3500, 1, 0))

dat_old <- merge(old_blood_data, old_bm_data, by="sn") %>%
  rename(blood_dt = "date.x", bm_dt = "date.y", visit_bld = "visit.x", visit_bm = "visit.y") %>%
  mutate(diff_dt = bm_dt - blood_dt) %>%
  mutate(se_bm_scaled = se_bm*1000) %>%
  mutate(cp_250 = ifelse(se_blood >250, 1, 0)) %>%
  mutate(cp_500 = ifelse(se_blood >500, 1, 0)) %>%
  mutate(cp_1000 = ifelse(se_blood >1000, 1, 0)) %>%
  mutate(cp_2000 = ifelse(se_blood >2000, 1, 0)) %>%
  mutate(cp_3000 = ifelse(se_blood >3000, 1, 0)) %>%
  mutate(cp_3500 = ifelse(se_blood >3500, 1, 0))

dat_both <- rbind(dat %>% mutate(stage="Version 2"), dat_old %>% mutate(stage="Version 1"))

# Number of cows sampled on each farm for the 2 analysis approaches:
n_cows_farm1 <- dat_old %>%
  group_by(sn) %>%
  summarise(n()) 
#table(n_cows_farm1$`n()`)

n_cows_farm2 <- dat %>%
  group_by(sn) %>%
  summarise(n()) 
#table(n_cows_farm2$`n()`)

# First visit stats - original dataset
n_farms1   <- length(unique(old_blood_data$sn))
n_cows1    <- nrow(old_blood_data)
dates_bld1 <- paste0(format(as.Date(min(dat_old$blood_dt), format = "%Y-%m-%d"), "%d/%m/%Y"), " and ", format(as.Date(max(dat_old$blood_dt), format = "%Y-%m-%d"), "%d/%m/%Y"))
dates_bm1  <- paste0(format(as.Date(min(dat_old$bm_dt), format = "%Y-%m-%d"), "%d/%m/%Y"), " and ", format(as.Date(max(dat_old$bm_dt), format = "%Y-%m-%d"), "%d/%m/%Y"))

# Second visit stats - updated dataset
n_farms2   <- length(unique(bld$sn))
n_cows2    <- nrow(bld)
dates_bld2 <- paste0(format(as.Date(min(dat$blood_dt), format = "%Y-%m-%d"), "%d/%m/%Y"), " and ", format(as.Date(max(dat$blood_dt), format = "%Y-%m-%d"), "%d/%m/%Y"))
dates_bm2  <- paste0(format(as.Date(min(dat$bm_dt), format = "%Y-%m-%d"), "%d/%m/%Y"), " and ", format(as.Date(max(dat$bm_dt), format = "%Y-%m-%d"), "%d/%m/%Y"))

############################################################################################################################################
# Questionnaire setup
questions <- read_excel(path = "Data Entry - Questionnaire .xlsx", range = cell_rows(3), col_names = F) %>%
  transpose() %>%
  set_colnames("question") %>%
  mutate(q_code = paste0("q_", gsub( " .*$", "", question)), q_text = sub("^\\S+\\s+", '', question)) %>%
  mutate(labels = paste0(q_code, "~", "'", q_text, "'"))

question_categories    <- read_excel(path = "Data Entry - Questionnaire .xlsx", range = cell_rows(1), col_names = F) %>%
  transpose() %>%
  set_colnames("category") %>%
  filter(!is.na(category))

missing_subcategories <- data.frame(subcategory = c("2.1 Do you provide any selenium supplementation to your adult dairy cattle",
                           "2.3 How much selenium do you provide?",
                           "2.5 What is the estimated cost of supplementing the herd in the way that you do?",
                           "2.6 Have you changed how you have been supplementing your dairy cattle in the past 3 years?",
                           "2.7 If you use multiple methods of supplementation for your dairy cattle - why do you do that?",
                           "2.9 Have you ever seen any signs of selenium deficienty in your dairy cattle?",
                           "2.10 Have you ever seen signs of selenium deficiency in your neighbours stock?",
                           "2.11 What do you think are the signs of selenium deficiency in cattle?",
                           "2.12 What do you think the impacts are of lower that optimal selenium in your herd?",
                           "2.16 When was the last time you tested your dairy cattle?"))


question_subcategories <- read_excel(path = "Data Entry - Questionnaire .xlsx", range = cell_rows(2), col_names = F) %>%
  transpose() %>%
  set_colnames("subcategory") %>%
  filter(!is.na(subcategory)) %>%
  rbind(missing_subcategories) %>%
  mutate(subcat_code = gsub( " .*$", "", subcategory), subcat_text = sub("^\\S+\\s+", '', subcategory)) %>%
  mutate(subcat_code2 = as.numeric(sub(".*\\.", "", subcat_code))) %>%
  arrange(subcat_code2) %>%
  select(-subcat_code2) %>%
  mutate(subcategory = ifelse(subcat_code == "2.17", "Why do you choose not to test your dairy cattle for selenium?", subcategory))
# Edit Q17 - the questionnaire actually used didn't contain "Or what are your reasons for a reduced frequency of testing?"

que <- read_excel(path = "Data Entry - Questionnaire .xlsx", skip = 2) %>%
  set_colnames(questions$q_code) %>%
  mutate(q_1.2 = ifelse(q_1.2=="Oto/TA", "Otorohanga/Te Awamutu", q_1.2)) %>%
  mutate(q_1.2 = fct_relevel(q_1.2, "BOP", "Otorohanga/Te Awamutu", "Putaruru", "Taranaki", "Manawatu", "Canterbury", "Otago")) %>%
  mutate(q_1.3 = ifelse(q_1.3=="EpiVets", "VetOra", q_1.3)) %>%
  mutate(q_1.5 = ifelse(q_1.5 %in% c("F10", "Fr", "Fr - FX", "HF"), "F", 
                        ifelse(q_1.5 %in% c("Kiwi X", "FX", "Jersey X", "Jersey, Kiwi X", "Kiwi X", "Mixed", "X Breed", "XB-Fr", "Fr, Kiwi X"), "XB", "J"))) %>%
  mutate(q_1.6 = case_when(q_1.6 =="forgotten" ~ NA,
                           q_1.6 =="510kg/cow" ~ 510*930,
                           q_1.6 =="238,000 (Last season), 255,440 (3yr Average)" ~ 238000,
                           q_1.6 == "400/cow" ~ 400*780,
                           q_1.6 == "330/cow" ~ 330*240,
                           q_1.6 == "1256Kg/day" ~ NA,
                           q_1.6 =="140-150,000" ~ 145000,
                           q_1.6 =="430 per cow (calculated 215,000)" ~ 125000,
                           q_1.6 =="420" ~ 420000,
                           q_1.6 =="448" ~ 448000,
                           q_1.6 == "126,000 (152,000 last season)" ~ 126000,
                           .default = as.numeric(as.character(q_1.6)))) %>%
  mutate(q_1.7 = case_when(q_1.7 =="Forgotten" ~ NA,
                           q_1.7 =="71% (Last season), 69% (3yr average)" ~ 0.71,
                           q_1.7 =="NA" ~ NA,
                           q_1.7 == "Low only 3 weeks AB" ~ NA,
                           q_1.7 == "Not sure" ~ NA,
                           q_1.7 == "65,68,63,63% (from 22/23 season to 19/20 season)" ~ 0.65,
                           q_1.7 =="Unknown (new herd)" ~ NA,
                           .default = as.numeric(as.character(q_1.7)))) %>%
  mutate(q_1.8 = case_when(q_1.8 =="NA" ~ NA,
                           q_1.8 == "10% (Last season), 13% (3yr average)" ~ 0.1,
                           q_1.8 == "14-17" ~ 0.155,
                           q_1.8 == "15-16%" ~ 0.155,
                           q_1.8 =="Unknown (new herd)" ~ NA,
                           .default = as.numeric(as.character(q_1.8)))) %>%
  mutate(q_1.9 = case_when(q_1.9 =="NA" ~ NA,
                           q_1.9 == "?" ~ NA,
                           q_1.9 == "140,000 (last season), 102,500 (3yr average)" ~ 140000,
                           q_1.9 == "120000-200000" ~ 160000,
                           q_1.9 =="164000-199000" ~ (164000+199000)/2,
                           q_1.9 == "165,000-167,000" ~ 166000,
                           q_1.9 == "229,000 (to date this season, no historic data)" ~ 229000,
                           .default = as.numeric(as.character(q_1.9)))) %>%
  mutate(q_1.7 = round(100*q_1.7, 0), q_1.8 = round(100*q_1.8, 0)) %>%
  mutate(q_2.1a = ifelse(q_2.1a=="Sometimes-depends" & q_2.1ax !="Maybe in mineral mix", "Yes", q_2.1a)) %>%
  mutate(q_2.8f = ifelse(q_2.8fx=="Fert reps advice", 1, q_2.8f)) %>%                     # This respondent should have selected "other".
  mutate(q_2.10 = ifelse(grepl("busy", q_2.10), "NA", q_2.10)) %>%                        # Changed the "too busy" to NA.
  mutate(q_2.17d = ifelse(q_2.17d=="NA" & q_2.17a=="1", "0", q_2.17d),
         q_2.17e = ifelse(q_2.17e=="NA" & q_2.17a=="1", "0", q_2.17e)) %>%                # Stacey accidentally entered "NA" instead of 0
  mutate(q_2.17e = ifelse(grepl("Feed at recommended rates", q_2.17ex), 1, q_2.17e)) %>%  # Should have been "other".
  mutate(q_2.21e = ifelse(grepl("ENL", q_2.21ex), 1, q_2.21e)) %>%                        # Should have been "other".
  mutate(q_2.13e = ifelse(grepl("muscle", q_2.13ex), 0, q_2.13e), q_2.13a = ifelse(grepl("muscle", q_2.13ex), 1, q_2.13a))  %>% # Deal with clinical sign entered as other.
  mutate(across(q_2.17a:q_2.17e, ~ifelse(grepl("periodically", q_2.17ex), "NA", .)))      # Should be a no-response as this person clearly tests.

completeness <- que %>%
  select(!ends_with("x")) %>%
  mutate(across(where(is.character), ~as.factor(na_if(., "NA")))) %>%
  rowwise() %>%
  mutate(incomplete = ifelse(sum(is.na(.)) >0, 1, 0))

# Plot functions
width <- 0.5
# Plot for "pick all that apply" questions with "other" option
multi_answer_plot_with_other <- function(data, questions_start, questions_end, width = 0.5) {
  
plot <- data %>% 
  select(c(questions_start:questions_end)) %>%
  select(!ends_with("x")) %>%
  mutate_if(is.character, list(~ ifelse(. == "NA", NA, .))) %>%
  mutate_if(is.character, as.numeric) %>%
  rowwise() %>%
  filter(if_any(everything(), ~ !is.na(.))) %>%
  ungroup() %>%
  rename_with(~ questions$q_text[match(., questions$q_code)], starts_with("Q")) %>%
  pivot_longer(cols=everything(), names_to = "option", values_to="value") %>% 
  group_by(option) %>%
  summarise(n = sum(value), prop = sum(value)/length(value)) %>% 
  ungroup() %>%
  ggplot(aes(x = fct_reorder(option, prop, .desc = T), y = prop, fill = option)) + 
  geom_bar(stat = "identity", position = position_dodge(), col="black", width=width) +  
  geom_text(aes(label = paste0(round(prop * 100, 0), "%"), y = prop + 0.05, x = as.numeric(factor(fct_reorder(option, prop, .desc = T))) + 0.05), 
            position  = position_dodge(width = width),   
            #vjust     = -3,
            size = 3,
            alpha = 1,
            #min.segment.length = 10,
            color = "black",
            #label.padding = 0.15,
            #box.padding = 1,
            #direction = "y",
            #ylim = c(prop+0.05, 1),
            show.legend = FALSE) +
  scale_y_continuous(labels = scales::percent, limits = c(0, 1.1), breaks = seq(0, 1, 0.25)) + 
  ylab("Percentage of responders") + 
  theme +
  theme(axis.title.x = element_blank(), axis.text.x = element_text(size=11, colour = "black", angle = 45, vjust = 1, hjust=1)) +
  guides(fill="none") +
  scale_x_discrete(labels = wrap_format(20))

return(plot)

}


# Response count function
multi_answer_counts_with_other <- function(data, questions_start, questions_end, width = 0.5) {
  
  n <- data %>% 
    select(c(questions_start:questions_end)) %>%
    select(!ends_with("x")) %>%
    mutate_if(is.character, list(~ ifelse(. == "NA", NA, .))) %>%
    mutate_if(is.character, as.numeric) %>%
    rowwise() %>%
    filter(if_any(everything(), ~ !is.na(.))) %>%
    ungroup() %>%
    rename_with(~ questions$q_text[match(., questions$q_code)], starts_with("Q")) %>%
    pivot_longer(cols=everything(), names_to = "option", values_to="value") %>% 
    group_by(option) %>%
    summarise(n = sum(value)) 
  
  stats <- data %>% 
    select(c(questions_start:questions_end)) %>%
    select(!ends_with("x")) %>%
    mutate_if(is.character, list(~ ifelse(. == "NA", NA, .))) %>%
    mutate_if(is.character, as.numeric) %>%
    rowwise() %>%
    filter(if_any(everything(), ~ !is.na(.))) %>%
    ungroup() %>%
    rename_with(~ questions$q_text[match(., questions$q_code)], starts_with("Q")) %>%
    mutate(n_options = rowSums(.)) %>%
    select(n_options)
  
  df <- data %>% 
    select(c(questions_start:questions_end)) %>%
    select(!ends_with("x")) %>%
    mutate_if(is.character, list(~ ifelse(. == "NA", NA, .))) %>%
    mutate_if(is.character, as.numeric)
  
  not_app <- data %>% 
    select(c(questions_start:questions_end)) %>%
    select(!ends_with("x")) %>%
    rowwise() %>%
    mutate(not_applicable = as.integer(all(c_across(everything()) == "NA"))) %>%
    filter(not_applicable==1)
  
  return(list(median    = median(stats$n_options),
              range     = noquote(paste0(min(stats$n_options), "-", max(stats$n_options))),
              n_skipped = sum(stats$n_options==0),
              n         = n,
              n_not_applicable = nrow(not_app),
              df        = df,
              df_na     = not_app,
              n_other   = n[n$option=="Other", ]$n))
}

