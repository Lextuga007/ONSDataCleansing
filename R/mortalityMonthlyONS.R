library(readxl)
library(janitor)
library(tidyverse)
library(stringi)
library(lubridate)
library(tidyxl)

# Source and licence acknowledgement

# This data has been made available through Office of National Statistics under the Open Government
# Licence http://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/

#https://www.ons.gov.uk/peoplepopulationandcommunity/birthsdeathsandmarriages/deaths/datasets/weeklyprovisionalfiguresondeathsregisteredinenglandandwales

# Create folders if do not exist. Folders not pushed to GitHub.

if(!dir.exists("spreadsheets/monthly")) {
  dir.create("spreadsheets/monthly", recursive = T)
}


# Download data -----------------------------------------------------------

# 2020
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2020/publishedoutputfebruary20202.xls",
  destfile = "spreadsheets/monthly/Monthly2020Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2019
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2019/publishedoutputdecember2019.xls",
  destfile = "spreadsheets/monthly/Monthly2019Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2018
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2018/publishedannual2018.xls",
  destfile = "spreadsheets/monthly/Monthly2018Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2017
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2017/publishedoutputannual2017final.xls",
  destfile = "spreadsheets/monthly/Monthly2017Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2016
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2016/publishedoutput2016final.xls",
  destfile = "spreadsheets/monthly/Monthly2016Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2015
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2015/publishedoutput2015final.xls",
  destfile = "spreadsheets/monthly/Monthly2015Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2014
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2014/publishedoutput2014finaltcm774115982.xls",
  destfile = "spreadsheets/monthly/Monthly2014Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2013
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2013/publishedoutput2013finaltcm773717241.xls",
  destfile = "spreadsheets/monthly/Monthly2013Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2012
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2012/publishedoutput2012finaltcm773197501.xls",
  destfile = "spreadsheets/monthly/Monthly2012Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2011
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2011/publishedoutput2011finaltcm772738151.xls",
  destfile = "spreadsheets/monthly/Monthly2011Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2010
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2010/publishedoutputfeb021tcm772274383.xls",
  destfile = "spreadsheets/monthly/Monthly2010Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2009
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2009/publishedoutput200tcm772274362.xls",
  destfile = "spreadsheets/monthly/Monthly2009Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2008
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2008/publishedoutput200tcm772274292.xls",
  destfile = "spreadsheets/monthly/Monthly2008Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2007
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2007/publishedoutput200tcm772274233.xls",
  destfile = "spreadsheets/monthly/Monthly2007Mortality.xls",
  method = "wininet",
  mode = "wb")

# 2006
download.file(
  "https://www.ons.gov.uk/file?uri=%2fpeoplepopulationandcommunity%2fbirthsdeathsandmarriages%2fdeaths%2fdatasets%2fmonthlyfiguresondeathsregisteredbyareaofusualresidence%2f2006/publishedoutput200tcm772274163.xls",
  destfile = "spreadsheets/monthly/Monthly2006Mortality.xls",
  method = "wininet",
  mode = "wb")


# Extract all worksheets to individual csv 2010-2019 -------------------------------------------------------------

files_list <- list.files(path = "spreadsheets/monthly",
                         pattern = "*.xls",
                         full.names = TRUE)


read_then_csv <- function(sheet, path) {
  pathbase <- path %>%
    basename() %>%
    tools::file_path_sans_ext()
  path %>%
    read_excel(sheet = sheet) %>%
    write_csv(paste0("spreadsheets/monthly/", pathbase, "-", sheet, ".csv"))
}


for(j in 1:length(files_list)){

  path <- paste0(files_list[j])

  path %>%
    excel_sheets() %>%
    set_names() %>%
    map(read_then_csv, path = path)
}


# Reload just Figures worksheet -------------------------------------

files_list_sheets <- list.files(path = "spreadsheets/monthly",
                                pattern = "Figures",
                                full.names = TRUE
)

for(i in files_list_sheets) {

  x <- read_csv((i), col_types = cols(.default = col_character()))

  assign(i, x)
}


# Format data 2011 -----------------------------------------------------------
# format out of date order as 2011 was unique in having old and new codes in the data. 
# Use this to create the tier for the earlier and later years' data

ONS <- `spreadsheets/monthly/Monthly2011Mortality-Figures for 2011.csv` %>%
  clean_names %>%
  rename_at(vars(starts_with("Monthly")), ~("codes")) %>% 
  mutate(x2 = case_when(x2 == "Former districts of:" ~ NA_character_,
                        TRUE ~ x2),
         x2 = str_remove(x2, ","),
         codes = coalesce(codes, x2),
         # Remove the footnote numbers in category
         codes = case_when(str_detect(codes, pattern = '[0-9]') ~ substr(codes, 1, str_locate(codes, pattern = '[0-9]') - 1),
                           TRUE ~ codes)
  ) %>% 
  remove_empty(c("rows","cols")) %>% 
  fill(codes, .direction = "down") %>% 
  filter(!is.na(x4)) %>%  # remove footnotes as no value in x4
  select(-x2) %>% 
  select(codes, everything()) %>% 
  mutate(codes = case_when(is.na(codes) ~ "tier",
                           TRUE ~ codes),
         x3 = case_when(codes == "tier" ~ "category",
                        TRUE ~ x3)) 

# Push date row to column names

onsFormattedJanitor <- row_to_names(ONS, 1)

Mortality2011 <- onsFormattedJanitor %>%
  pivot_longer(cols = -tier:-category,
               names_to = "dates",
               values_to = "counts") %>% 
  select(tier,
         category,
         dates,
         counts)


# Tier lookup -------------------------------------------------------------

tierLookup <- Mortality2011 %>% 
  select(tier, 
         category) %>% 
  unique() %>% 
  filter(!is.na(category))


# Format data 2006 - 2010 -----------------------------------------------------------

formatFunction <- function(file){

ONS <- file %>%
  clean_names %>%
  remove_empty(c("rows","cols")) %>%
  rename_at(vars(starts_with("Monthly")), ~("codes")) %>%
  mutate(category = coalesce(x2, x3, x4),
         # Remove the footnote numbers in category
         category = case_when(str_detect(category, pattern = '[0-9]') ~ substr(category, 1, str_locate(category, pattern = '[0-9]') - 1),
                              TRUE ~ category)) %>%
  select(category, everything()) %>%
  filter(!codes %in% c('Footnotes:',
                          '1')) %>%
  select(-x2, -x3, -x4) %>%
  mutate(category = case_when(is.na(category) ~ 'category',
                              TRUE ~ category),
         codes = case_when(str_detect(codes, "Area Codes") ~ "area_codes",
                           TRUE ~ codes))

  # Push date row to column names

  onsFormattedJanitor <- row_to_names(ONS, 1)

  x <- onsFormattedJanitor %>%
    pivot_longer(cols = -category:-area_codes,
                 names_to = "dates",
                 values_to = "counts") %>% 
    left_join(tierLookup) %>% 
    select(tier, 
           everything())

return(x)

}

Mortality2006 <- formatFunction(`spreadsheets/monthly/Monthly2006Mortality-Figures for 2006.csv`)
#1	Area Codes England and Wales were recoded in July 2007.
#2	TOTAL REGISTRATIONS This may include records where the place of usual residence is missing. All sub divisions of Total Registrations exclude these.

# Category regions are capitalised   filter(str_detect(category, "^[^[:lower:]]{2,}$"))
# ends_with() UA - unitary authority (single tier, no districts or parishes)
#  (Met County)	


Mortality2007 <- formatFunction(`spreadsheets/monthly/Monthly2007Mortality-Figures for 2007.csv`)
Mortality2008 <- formatFunction(`spreadsheets/monthly/Monthly2008Mortality-Figures for 2008.csv`)
Mortality2009 <- formatFunction(`spreadsheets/monthly/Monthly2009Mortality-Figures for 2009.csv`)
Mortality2010 <- formatFunction(`spreadsheets/monthly/Monthly2010Mortality-Figures for 2010.csv`)


# Extract area codes ------------------------------------------------------

areaCodes <- Mortality2006 %>% 
  select(category,
         area_codes) %>% 
  unique() %>% 
  filter(!is.na(area_codes))


# Mortality2011 -----------------------------------------------------------

Mortality2011 <- Mortality2011 %>% 
   left_join(areaCodes) 
  

# Format data 2012 - 2020 -----------------------------------------------------------
# This was a cross over year with old and new Authority names in the spreadsheet

formatFunction <- function(file){
  
  ONS <- file %>%
    clean_names %>%
    mutate(contents = coalesce(contents, x2, x3)) %>%
    remove_empty(c("rows","cols")) %>% 
    select(-x2, -x3) %>% 
    filter(!is.na(x4))    # remove footnotes as no value in x4
  
  # Push date row to column names
  
  onsFormattedJanitor <- row_to_names(ONS, 1)
  
  x <- onsFormattedJanitor %>%
    rename(category = `Area of usual residence`) %>% 
    pivot_longer(cols = -category,
                 names_to = "dates",
                 values_to = "counts") %>% 
    left_join(tierLookup) %>% 
    select(tier, 
           everything())
  
  
  return(x)
  
}


Mortality2012 <- formatFunction(`spreadsheets/monthly/Monthly2012Mortality-Figures for 2012.csv`)
Mortality2013 <- formatFunction(`spreadsheets/monthly/Monthly2013Mortality-Figures for 2013.csv`)
Mortality2014 <- formatFunction(`spreadsheets/monthly/Monthly2014Mortality-Figures for 2014.csv`)
Mortality2015 <- formatFunction(`spreadsheets/monthly/Monthly2015Mortality-Figures for 2015.csv`)
Mortality2016 <- formatFunction(`spreadsheets/monthly/Monthly2016Mortality-Figures for 2016.csv`)
Mortality2017 <- formatFunction(`spreadsheets/monthly/Monthly2017Mortality-Figures for 2017.csv`)
Mortality2018 <- formatFunction(`spreadsheets/monthly/Monthly2018Mortality-Figures for 2018.csv`)
Mortality2019 <- formatFunction(`spreadsheets/monthly/Monthly2019Mortality-Figures for 2019.csv`)
Mortality2020 <- formatFunction(`spreadsheets/monthly/Monthly2020Mortality-Figures for 2020.csv`)


# Bind together -----------------------------------------------------------

ons_mortality_monthly <- do.call("rbind", list(
  Mortality2006,
  Mortality2007,
  Mortality2008,
  Mortality2009,
  Mortality2010,
  Mortality2011))

#   Mortality2012,
#   Mortality2013,
#   Mortality2014,
#   Mortality2015,
#   Mortality2016,
#   Mortality2017,
#   Mortality2018,
#   Mortality2019,
#   Mortality2020))
# 
# 
#   # Save as rda file
# 
#   save(ons_mortality_monthly, file = "data/ons_mortality_monthly.rda")
# 
#   # Save as an csv file
# 
#   write_csv(ons_mortality_monthly, file = "`Working files`/ons_mortality_monthly.csv")
