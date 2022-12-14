setwd('~/College/Nievergelt Lab')
library(readxl)
library(openxlsx)

original <- read_excel('biobank_longformat_10_25_22.xlsx', sheet = 'Database')
new <- read_excel('NEW_SHEET.xlsx', sheet = 'Sheet 1')