setwd('~/College/Nievergelt Lab')
library(readxl)
library(openxlsx)

old <- read_excel('biobank_longformat_10_25_22.xlsx', sheet = 'Database')
new <- read_excel('NEW_SHEET.xlsx', sheet = 'Sheet 1')

for (i in 1:nrow(old)) {
	for (j in 1:nrow(new)) {
		if (is.na(old[i,5]) == FALSE & old[i,5] == new[i,1]) {
			if (old[i,4] == new[i,3]) print ('match found')
		}
	}
}