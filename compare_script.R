setwd('~/College/Nievergelt Lab')
library(readxl)
library(openxlsx)

new <- unname(unlist(read_excel('compare_sheet.xlsx', sheet = 'new')[,10]))
old <- unname(unlist(read_excel('compare_sheet.xlsx', sheet = 'old')[,10]))

n <- 0
for (i in new) {
	if (i == 'no') {
		n <- n + 1
	}
}

m <- 0
for (i in old) {
	if (i == 'no') {
		m <- m + 1
	}
}