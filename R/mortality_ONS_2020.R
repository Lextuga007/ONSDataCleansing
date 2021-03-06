library(readxl)
library(janitor)
library(tidyverse)
library(stringi)
library(lubridate)
library(NHSRdatasets)
library(httr)
library(rvest)

# Source and licence acknowledgement

# This data has been made available through Office of National Statistics under the Open Government
# Licence http://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/

#https://www.ons.gov.uk/peoplepopulationandcommunity/birthsdeathsandmarriages/deaths/datasets/weeklyprovisionalfiguresondeathsregisteredinenglandandwales

# Download data -----------------------------------------------------------
# 2020 Format changed to xlsx from xls
download.file(
  "https://www.ons.gov.uk/file?uri=/peoplepopulationandcommunity/birthsdeathsandmarriages/deaths/datasets/weeklyprovisionalfiguresondeathsregisteredinenglandandwales/2020/publishedweek532020.xlsx",
  destfile = "spreadsheets/weekly/2020Mortality.xlsx",
  method = "wininet",
  mode = "wb")


# 2021 Format changed to xlsx from xls

# ext built from gist shared by rcatlord https://gist.github.com/rcatlord/1b44259e23bf8ef76bb54cc14d60d969

ext <- read_html("https://www.ons.gov.uk/peoplepopulationandcommunity/birthsdeathsandmarriages/deaths/datasets/weeklyprovisionalfiguresondeathsregisteredinenglandandwales") %>% 
  html_nodes("a") %>%
  html_attr("href") %>%
  str_subset("\\.xlsx") %>% 
  .[[1]]

download.file(
  paste0("https://www.ons.gov.uk", ext),
  destfile = "spreadsheets/weekly/2021Mortality.xlsx",
  method = "wininet",
  mode = "wb")

# Extract all worksheets to individual csv 2020 -------------------------------------------------------------

files_list <- list.files(path = "spreadsheets/weekly",
                         pattern = "*.xlsx",
                         full.names = TRUE)


read_then_csv <- function(sheet, path) {
  pathbase <- path %>%
    basename() %>%
    tools::file_path_sans_ext()
  path %>%
    read_excel(sheet = sheet) %>%
    write_csv(paste0("spreadsheets/weekly/", pathbase, "-", sheet, ".csv"))
}


for(j in 1:length(files_list)){
  
  path <- paste0(files_list[j])
  
  path %>%
    excel_sheets() %>%
    set_names() %>%
    map(read_then_csv, path = path)
}

# Reload just weekly figure worksheet -------------------------------------

# From 2010 to 2015 the tab name was Weekly Figures then it changed capitisation to Weekly figures

files_list_sheets <- list.files(path = "spreadsheets/weekly",
                                pattern = "Weekly figures",
                                full.names = TRUE
)

for(i in files_list_sheets) {
  
  x <- read_csv((i), col_types = cols(.default = col_character()))
  
  assign(i, x)
}

# Repeated code -----------------------------------------------------------

remove_lookup <- c('week over the previous five years1',
                   'Deaths by underlying cause2,3',
                   'Footnotes',
                   '1 This average is based on the actual number of death registrations recorded for each corresponding week over the previous five years. Moveable public holidays, when register offices are closed, affect the number of registrations made in the published weeks and in the corresponding weeks in previous years.',
                   '2 Counts of deaths by underlying cause exclude deaths at age under 28 days.',
                   '3 Coding of deaths by underlying cause for the latest week is not yet complete.',
                   "4Does not include deaths where age is either missing or not yet fully coded. For this reason counts of 'Persons', 'Males' and 'Females' may not sum to 'Total Deaths, all ages'.",
                   '5 Does not include deaths of those resident outside England and Wales or those records where the place of residence is either missing or not yet fully coded. For this reason counts for "Deaths by Region of usual residence" may not sum to "Total deaths, all ages".',
                   'Source: Office for National Statistics',
                   'Deaths by age group'
)

# Format data  -------------------------------------------------

# Format data skip line is 2
# Added formatting for age bands in line with historical data
# 2021 - note that there is a line in the original spreadsheet for 2020 deaths but as this is for total, 
# this code still relies on specific extraction of 2020 and 2021 data to get all areas for 2020

formatFunction2020 <- function(file){
  
  ONS <- file %>%
    clean_names %>%
    mutate(x2 = case_when(is.na(x2) ~ contents,
                          TRUE ~ x2),
           x2 = recode(x2, "<1" = "Under 1 year")) %>%
    remove_empty(c("rows","cols")) %>%
    select(-contents) %>%
    filter(!x2 %in% remove_lookup) %>%
    mutate(Category = case_when(#is.na(x3) & str_detect(x2, " 4") ~ str_replace(x2, " 4", ""),
      is.na(x3) & str_detect(x2, "region") ~ "Region",
      is.na(x3) & str_detect(x2, "Persons") ~ "Persons",
      is.na(x3) & str_detect(x2, "Females") ~ "Females",
      is.na(x3) & str_detect(x2, "Males") ~ "Males",
      TRUE ~ NA_character_)
    ) %>%
    select(x2, Category, everything()) %>%
    fill(Category) %>%
    filter(!str_detect(x2, 'Persons'),
           !str_detect(x2, 'People'), # used in 2021
           !str_detect(x2, 'Males'),
           !str_detect(x2, 'Females')) %>%
    unite("Categories", Category, x2) %>%
    filter(!is.na(x3)) %>% 
    mutate(Categories = case_when(str_detect(Categories, "NA_") ~ str_replace(Categories, "NA_", ""),
                                  str_detect(Categories, "Week ended") ~ "Week ended",
                                  str_detect(Categories, "Week number") ~ "Week number",
                                  TRUE ~ Categories),
           Categories = case_when(str_detect(Categories, "week over the previous 5") ~ str_c("Total deaths: average of corresponding", Categories, sep = " "),
                                  TRUE ~ Categories)
    ) 
  
  # Push date row to column names
  
  onsFormattedJanitor <- row_to_names(ONS, 2)
  
  x <- onsFormattedJanitor %>%
    pivot_longer(cols = -`Week ended`,
                 names_to = "allDates",
                 values_to = "counts") %>%
    mutate(realDate = dmy(allDates),
           ExcelSerialDate = case_when(stri_length(allDates) == 5 ~ excel_numeric_to_date(as.numeric(allDates), date_system = "modern")),
           date = case_when(is.na(realDate) ~ ExcelSerialDate,
                            TRUE ~ realDate)) %>%
    group_by(`Week ended`) %>%
    mutate(week_no = row_number()) %>%
    ungroup() %>%
    rename(Category = `Week ended`) %>%
    mutate(category_1 = case_when(str_detect(Category, ",") ~
                                    substr(Category,1,str_locate(Category, ",") -1),
                                  str_detect(Category, ":") ~
                                    substr(Category,1,str_locate(Category, ":") -1),
                                  str_detect(Category, "_") ~
                                    substr(Category,1,str_locate(Category, "_") -1),
                                  str_detect(Category, "respiratory")  ~
                                    "All respiratory diseases (ICD-10 J00-J99) ICD-10",
                                  TRUE ~ Category),
           category_2 = case_when(str_detect(Category, ",") ~
                                    substr(Category,str_locate(Category, ", ") +2, str_length(Category)),
                                  str_detect(Category, ":") ~
                                    substr(Category,str_locate(Category, ": ") +2, str_length(Category)),
                                  str_detect(Category, "_") ~
                                    substr(Category,str_locate(Category, "_") +1, str_length(Category)),
                                  str_detect(Category, "respiratory")  ~
                                    substr(Category,str_locate(Category, "v"), str_length(Category)),
                                  TRUE ~ NA_character_)
    ) %>%
    select(category_1,
           category_2,
           counts,
           date,
           week_no
    ) %>%
    filter(!is.na(counts),
           !is.na(date)) # In 2021 week 53 has no date - perhaps runs into 2022
  
  return(x)
  
}

Mortality2020 <- formatFunction2020(`spreadsheets/weekly/2020Mortality-Weekly figures 2020.csv`)
Mortality2021 <- formatFunction2020(`spreadsheets/weekly/2021Mortality-Weekly figures 2021.csv`)


# Checks

summary(Mortality2020$date)
summary(Mortality2021$date)

summary(Mortality2020$week_no)

# Bind together -----------------------------------------------------------
# taken from the NHSRdatasets GitHub but will be from the package in due course

ons_mortality <- NHSRdatasets::ons_mortality

ons_mortality <- do.call("rbind", list(ons_mortality,
                                       Mortality2020,
                                       Mortality2021))

# Save as rda file

save(ons_mortality, file = "data/ons_mortality_2020.rda")

# Save as a csv file

write_csv(ons_mortality, "data/ons_mortality_2020.csv")
