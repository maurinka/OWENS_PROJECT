
library(dplyr)
library(stringr)
library(stringi)
library(stringdist)
library(readxl)
library(writexl)
library(tidyverse)
library(conflicted)

orig <- read_excel("KnowldegeBase.xlsx", 
                            sheet = "02_10_26")

orig_csv <- read_csv("orig_csv.csv")

works <- read_csv("works.csv")


source("repair_knowledge_base.R")
head(orig)

j_csv <- match_rows(kb, csv)
kb <- update_rows(kb, csv, j_csv)

j_works <- match_rows(kb, works)
updated <- update_rows(kb, works, j_works)

full <- bind_rows(updated, new_rows) |> finish()

write_xlsx(full, "knowledge_base.xlsx")
