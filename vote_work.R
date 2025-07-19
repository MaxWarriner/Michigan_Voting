library(readxl)
library(dplyr)
library(ggplot2)
# Raw Data Processing' ----------------------------------------------------

#upload 2022 counties first
setwd("~/Voter Data/Raw Data/2022")
counties2022 <- readxl::read_xlsx('counties2022.xlsx')

# 1998 raw data
setwd("~/Voter Data/Raw Data/1998")

dat1998 <- readxl::read_xlsx("votes1998.xlsx")

names1998 <- read_xlsx("names1998.xlsx")

cities1998 <- read_xlsx("cities1998.xlsx")

parties1998 <- read_xlsx("parties1998.xlsx") |>
  rename(Candidate = CandidateID,
         Party = PartyDescription)

dat1998 <- merge(dat1998, names1998, by = "Candidate" )

dat1998 <- merge(dat1998, counties2022, by = "County")

names1998 <- merge(names1998, parties1998, by = "Candidate", all.x = TRUE) |>
  distinct()

dat1998 <- merge(dat1998, names1998)

dat1998 <- merge(dat1998, cities1998)

dat1998 <- dat1998 |>
  select(-County, -City, -Candidate)


# 2000 raw data
setwd("~/Voter Data/Raw Data/2000")

dat2000 <- readxl::read_xlsx("votes2000.xlsx")

names2000 <- read_xlsx("names2000.xlsx")

cities2000 <- read_xlsx("cities2000.xlsx")

parties2000 <- read_xlsx("parties2000.xlsx") |>
  rename(Candidate = CandidateID,
         Party = PartyDescription)

dat2000 <- merge(dat2000, names2000, by = "Candidate" )

dat2000 <- merge(dat2000, counties2022, by = "County")

names2000 <- merge(names2000, parties2000, by = "Candidate", all.x = TRUE) |>
  distinct()

dat2000 <- merge(dat2000, names2000)

dat2000 <- merge(dat2000, cities2000)

dat2000 <- dat2000 |>
  select(-County, -City, -Candidate)


# 2002 raw data
setwd("~/Voter Data/Raw Data/2002")

dat2002 <- readxl::read_xlsx("votes2002.xlsx")

names2002 <- read_xlsx("names2002.xlsx")

cities2002 <- read_xlsx("cities2002.xlsx")

parties2002 <- read_xlsx("parties2002.xlsx") |>
  rename(Candidate = CandidateID,
         Party = PartyDescription)

dat2002 <- merge(dat2002, names2002, by = "Candidate" )

dat2002 <- merge(dat2002, counties2022, by = "County")

names2002 <- merge(names2002, parties2002, by = "Candidate", all.x = TRUE) |>
  distinct()

dat2002 <- merge(dat2002, names2002)

dat2002 <- merge(dat2002, cities2002)

dat2002 <- dat2002 |>
  select(-County, -City, -Candidate)


# 2004 raw data
setwd("~/Voter Data/Raw Data/2004")

dat2004 <- readxl::read_xlsx("votes2004.xlsx")

names2004 <- read_xlsx("names2004.xlsx")

cities2004 <- read_xlsx("cities2004.xlsx")

parties2004 <- read_xlsx("parties2004.xlsx") |>
  rename(Candidate = CandidateID,
         Party = PartyDescription)

dat2004 <- merge(dat2004, names2004, by = "Candidate" )

dat2004 <- merge(dat2004, counties2022, by = "County")

names2004 <- merge(names2004, parties2004, by = "Candidate", all.x = TRUE) |>
  distinct()

dat2004 <- merge(dat2004, names2004)

dat2004 <- merge(dat2004, cities2004)

dat2004 <- dat2004 |>
  select(-County, -City, -Candidate)



# 2006 raw data
setwd("~/Voter Data/Raw Data/2006")

dat2006 <- readxl::read_xlsx("votes2006.xlsx")

names2006 <- read_xlsx("names2006.xlsx")

cities2006 <- read_xlsx("cities2006.xlsx")|>
  select(-Election)

parties2006 <- read_xlsx("parties2006.xlsx")

dat2006 <- merge(dat2006, names2006, by = "Candidate" )

dat2006 <- merge(dat2006, counties2022, by = "County")

names2006 <- merge(names2006, parties2006) |>
  distinct()

dat2006 <- merge(dat2006, names2006)

dat2006 <- merge(dat2006, cities2006)

dat2006 <- dat2006 |>
  select(-County, -City, -Candidate)

# 2008 raw data
setwd("~/Voter Data/Raw Data/2008")

dat2008 <- readxl::read_xlsx("votes2008.xlsx")

names2008 <- read_xlsx("names2008.xlsx")

cities2008 <- read_xlsx("cities2008.xlsx")

dat2008 <- merge(dat2008, names2008, by = "Candidate" )

dat2008 <- merge(dat2008, counties2022, by = "County")

dat2008 <- merge(dat2008, cities2008)

dat2008 <- dat2008 |>
  select(-County, -City, -Candidate)


# 2010 raw data
setwd("~/Voter Data/Raw Data/2010")

dat2010 <- readxl::read_xlsx("votes2010.xlsx")

names2010 <- read_xlsx("names2010.xlsx")

cities2010 <- read_xlsx("cities2010.xlsx")

dat2010 <- merge(dat2010, names2010, by = "Candidate" )

dat2010 <- merge(dat2010, counties2022, by = "County")

dat2010 <- merge(dat2010, cities2010)

dat2010 <- dat2010 |>
  select(-County, -City, -Candidate)


# 2012 raw data
setwd("~/Voter Data/Raw Data/2012")

dat2012 <- readxl::read_xlsx("votes2012.xlsx")

names2012 <- read_xlsx("names2012.xlsx")

cities2012 <- read_xlsx("cities2012.xlsx")

dat2012 <- merge(dat2012, names2012, by = "Candidate" )

dat2012 <- merge(dat2012, counties2022, by = "County")

dat2012 <- merge(dat2012, cities2012)

dat2012 <- dat2012 |>
  select(-County, -City, -Candidate)


# 2014 raw data
setwd("~/Voter Data/Raw Data/2014")

dat2014 <- readxl::read_xlsx("votes2014.xlsx")

names2014 <- read_xlsx("names2014.xlsx")

cities2014 <- read_xlsx("cities2014.xlsx")

dat2014 <- merge(dat2014, names2014, by = "Candidate" )

dat2014 <- merge(dat2014, counties2022, by = "County")

dat2014 <- merge(dat2014, cities2014)

dat2014 <- dat2014 |>
  select(-County, -City, -Candidate)


# 2016 raw data
setwd("~/Voter Data/Raw Data/2016")

dat2016 <- readxl::read_xlsx("votes2016.xlsx")

names2016 <- read_xlsx("names2016.xlsx")

cities2016 <- read_xlsx("cities2016.xlsx")

dat2016 <- merge(dat2016, names2016, by = "Candidate" )

dat2016 <- merge(dat2016, counties2022, by = "County")

dat2016 <- merge(dat2016, cities2016)

dat2016 <- dat2016 |>
  select(-County, -City, -Candidate)


# 2018 raw data
setwd("~/Voter Data/Raw Data/2018")

dat2018 <- readxl::read_xlsx("votes2018.xlsx")

names2018 <- read_xlsx("names2018.xlsx")

cities2018 <- read_xlsx("cities2018.xlsx")

dat2018 <- merge(dat2018, names2018, by = "Candidate" )

dat2018 <- merge(dat2018, counties2022, by = "County")

dat2018 <- merge(dat2018, cities2018)

dat2018 <- dat2018 |>
  select(-County, -City, -Candidate)


# 2020 raw data
setwd("~/Voter Data/Raw Data/2020")

dat2020 <- readxl::read_xlsx("votes2020.xlsx")

names2020 <- read_xlsx("names2020.xlsx")

cities2020 <- read_xlsx("cities2020.xlsx")

dat2020 <- merge(dat2020, names2020, by = "Candidate" )

dat2020 <- merge(dat2020, counties2022, by = "County")

dat2020 <- merge(dat2020, cities2020)

dat2020 <- dat2020 |>
  select(-County, -City, -Candidate)


# 2022 raw data
setwd("~/Voter Data/Raw Data/2022")

dat2022 <- readxl::read_xlsx("votes2022.xlsx")

names2022 <- read_xlsx("names2022.xlsx")

counties2022 <- read_xlsx("counties2022.xlsx")

cities2022 <- read_xlsx("cities2022.xlsx")

dat2022 <- merge(dat2022, names2022, by = "Candidate" )

dat2022 <- merge(dat2022, counties2022, by = "County")

dat2022 <- merge(dat2022, cities2022)

dat2022 <- dat2022 |>
  select(-County, -City, -Candidate)


# 2024 Raw Data
setwd("~/Voter Data/Raw Data/2024")

dat2024 <- readxl::read_xlsx('votes2024.xlsx')

names2024 <- read_xlsx("names2024.xlsx")

cities2024 <- read_xlsx("cities2024.xlsx")

dat2024 <- merge(dat2024, names2024, by = "Candidate" )

dat2024 <- merge(dat2024, counties2022, by = "County")

dat2024 <- merge(dat2024, cities2024)

dat2024 <- dat2024 |>
  select(-County, -City, -Candidate)



# Merge Everything Together -----------------------------------------------

old_data <- bind_rows(list(dat1998, dat2000, dat2002, dat2004, dat2006)) |>
  mutate(Party = case_when(Party == "Republican" ~ "REP", 
                           Party == "No Affiliation" ~ "NPA", 
                           Party == "Natural Law" ~ "NLP", 
                           Party == "Democratic" ~ "DEM", 
                           Party == "Libertarian" ~ "LIB", 
                           Party == "Reform" ~ "REF", 
                           Party == "US Taxpayers" ~ "UST", 
                           Party == "Green" ~ "GRN"))

dat <- bind_rows(list(old_data, dat2008, dat2010, dat2012, dat2014, dat2016, dat2018, dat2020, dat2022, dat2024))



# Cleaning ----------------------------------------------------------------

dat <- dat |>
  mutate(Office = case_when(Office == 1 ~ "President", 
                            Office == 2 ~ "Governor", 
                            Office == 3 ~ "SOS", 
                            Office == 4 ~ "AG", 
                            Office == 5 ~ "U.S. Senator", 
                            Office == 6 ~ "U.S. Rep", 
                            Office == 7 ~ "State Senator", 
                            Office == 8 ~ "State Rep", 
                            Office == 9 ~ "State BOE", 
                            Office == 10 ~ "UofM Board", 
                            Office == 11 ~ "MSU Board", 
                            Office == 12 ~ "Wayne State Board", 
                            Office == 13 ~ "Supreme Court", 
                            Office == 90 ~ "Ballot Proposal", 
                            Office == 0 ~ "Poll Total"), 
         District = case_when(District == 0 ~ NA, 
                               District == 100 ~ "1st", 
                               District == 200 ~ "2nd", 
                               District == 300 ~ "3rd", 
                               District == 400 ~ "4th", 
                               District == 500 ~ "5th", 
                               District == 600 ~ "6th", 
                               District == 700 ~ "7th", 
                               District == 800 ~ "8th", 
                               District == 900 ~ "9th", 
                               District == 1000 ~ "10th", 
                               District == 1100 ~ "11th", 
                               District == 1200 ~ "12th", 
                               District == 1300 ~ "13th", 
                               District == 1400 ~ "14th", 
                               District == 1500 ~ "15th", 
                               District == 1600 ~ "16th", 
                               District == 1700 ~ "17th", 
                               District == 1800 ~ "18th", 
                               District == 1900 ~ "19th", 
                               District == 2000 ~ "20th",
                               District == 2100 ~ "21st", 
                               District == 2200 ~ "22nd", 
                               District == 2300 ~ "23rd", 
                               District == 2400 ~ "24th", 
                               District == 2500 ~ "25th", 
                               District == 2600 ~ "26th", 
                               District == 2700 ~ "27th", 
                               District == 2800 ~ "28th", 
                               District == 2900 ~ "29th", 
                               District == 3000 ~ "30th", 
                               District == 3100 ~ "31st", 
                               District == 3200 ~ "32nd", 
                               District == 3300 ~ "33rd", 
                               District == 3400 ~ "34th", 
                               District == 3500 ~ "35th", 
                               District == 3600 ~ "36th", 
                               District == 3700 ~ "37th", 
                               District == 3800 ~ "38th", 
                               District == 3900 ~ "39th", 
                               District == 4000 ~ "40th", 
                               District == 4100 ~ "41st", 
                               District == 4200 ~ "42nd", 
                               District == 4300 ~ "43rd", 
                               District == 4400 ~ "44th", 
                               District == 4500 ~ "45th", 
                               District == 4600 ~ "46th", 
                               District == 4700 ~ "47th", 
                               District == 4800 ~ "48th", 
                               District == 4900 ~ "49th", 
                               District == 5000 ~ "50th", 
                               District == 5100 ~ "51st", 
                               District == 5200 ~ "52nd", 
                               District == 5300 ~ "53rd", 
                               District == 5400 ~ "54th", 
                               District == 5500 ~ "55th", 
                               District == 5600 ~ "56th", 
                               District == 5700 ~ "57th", 
                               District == 5800 ~ "58th", 
                               District == 5900 ~ "59th", 
                               District == 6000 ~ "60th",
                               District == 6100 ~ "61st", 
                               District == 6200 ~ "62nd", 
                               District == 6300 ~ "63rd", 
                               District == 6400 ~ "64th", 
                               District == 6500 ~ "65th", 
                               District == 6600 ~ "66th", 
                               District == 6700 ~ "67th", 
                               District == 6800 ~ "68th", 
                               District == 6900 ~ "69th", 
                               District == 7000 ~ "70th", 
                               District == 7100 ~ "71st", 
                               District == 7200 ~ "72nd", 
                               District == 7300 ~ "73rd", 
                               District == 7400 ~ "74th", 
                               District == 7500 ~ "75th", 
                               District == 7600 ~ "76th", 
                               District == 7700 ~ "77th", 
                               District == 7800 ~ "78th", 
                               District == 7900 ~ "79th", 
                               District == 8000 ~ "80th",
                               District == 8100 ~ "81st", 
                               District == 8200 ~ "82nd", 
                               District == 8300 ~ "83rd", 
                               District == 8400 ~ "84th", 
                               District == 8500 ~ "85th", 
                               District == 8600 ~ "86th", 
                               District == 8700 ~ "87th", 
                               District == 8800 ~ "88th", 
                               District == 8900 ~ "89th", 
                               District == 9000 ~ "90th", 
                               District == 9100 ~ "91st", 
                               District == 9200 ~ "92nd", 
                               District == 9300 ~ "93rd", 
                               District == 9400 ~ "94th", 
                               District == 9500 ~ "95th", 
                               District == 9600 ~ "96th", 
                               District == 9700 ~ "97th", 
                               District == 9800 ~ "98th", 
                               District == 9900 ~ "99th", 
                               District == 10000 ~ "100th", 
                               District == 10100 ~ "101st", 
                               District == 10200 ~ "102nd", 
                               District == 10300 ~ "103rd", 
                               District == 10400 ~ "104th", 
                               District == 10500 ~ "105th", 
                               District == 10600 ~ "106th", 
                               District == 10700 ~ "107th", 
                               District == 10800 ~ "108th", 
                               District == 10900 ~ "109th", 
                               District == 11000 ~ "110th"), 
         Status = case_when(Status == 0 ~ "Regular Term", 
                            Status == 1 ~ "Non-Incumbent", 
                            Status >=2 & Status <= 4 ~ "Incumbent - Partial Term", 
                            Status >=5 & Status <= 7 ~ "Non-Incumbent - Partial Term", 
                            Status == 8 ~ "Partial Term", 
                            Status == 9 | Status == 10 ~ "New Judgeship"), 
         Election = ifelse(Election == "GEN", "General", "Primary"), 
         Name = ifelse(Name == "NPA", NA, Name)) |>
  dplyr::select(-Precinct_Label, -Election) |>
  filter(Votes > 0)


dat <- dat |>
  mutate(Year = as.factor(Year), 
         Office = as.factor(Office), 
         District = as.factor(District), 
         Status = as.factor(Status), 
         Party = as.factor(Party))


library(data.table)
setwd("~/Voter Data")
fwrite(dat, file = "1998_2024_data.csv")



# Analysis ----------------------------------------------------------------

hd31 <- dat |>
  filter(Office == "State Rep", 
         District == "31st", 
         Party == "DEM" | Party == "REP") |>
  group_by(Year, Party, Name) |>
  summarise(Votes = sum(Votes))

ggplot(data = hd31, aes(x = Year, y = Votes, fill = Party)) +
  geom_col(position = position_dodge(width = 0.9)) +  # Dodge bars
  geom_text(
    aes(label = Name, y = Votes + 100), 
    position = position_dodge(width = 0.9),  # Dodge labels to match bars
    size = 3, 
    vjust = -0.5  # Adjust vertical position of labels if needed
  ) +
  theme_minimal() +
  geom_hline(yintercept = 0) +
  scale_fill_manual(values = c("blue", "red")) +
  ggtitle("House District 31") +
  theme(plot.title = element_text(hjust = 0.5))

hd31_cities_DEM <- dat |>
  filter(Office == "State Rep", 
         District == "31st", 
         Party == "DEM") |>
  group_by(Year, City_Name) |>
  summarise(Votes = sum(Votes)) |>
  filter(Votes > 0)

hd31_cities_REP <- dat |>
  filter(Office == "State Rep", 
         District == "31st", 
         Party == "REP") |>
  group_by(Year, City_Name) |>
  summarise(Votes = sum(Votes)) |>
  filter(Votes > 0)


DEM_HD31 <- ggplot(data = hd31_cities_DEM, aes(x = Year, y = Votes, fill = City_Name)) +
  geom_col() +
  ggtitle("Democratic Votes: HD31")

REP_HD31 <- ggplot(data = hd31_cities_REP, aes(x = Year, y = Votes, fill = City_Name)) +
  geom_col() +
  ggtitle("Republican Votes: HD31")

library(patchwork)
DEM_HD31 + REP_HD31

library(ggthemes)
ggplot(data = hd31_cities, aes(x = Party, y = Votes, fill = City_Name)) +
  geom_col(color = "black") +
  facet_grid(.~Year) +
  ggtitle("Party Votes by Year and City") +
  theme(plot.title = element_text(hjust = 0.5)) +
  theme_economist() +
  geom_hline(yintercept = 0) + 
  theme(legend.title = element_text(size = 10), 
        legend.text = element_text(size = 10)) +
  theme(panel.border = element_rect(color = "black", 
                                    fill = NA, 
                                    size = 2))


