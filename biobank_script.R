setwd('~/College/Nievergelt Lab')
library(readxl)
library(openxlsx)

boxkey_nums <- c()
numberOfBoxkeys <- 0
ID_visit <- c()
row <- c()
column <- c()
tissue <- c()
LPS_concentration <- c()
box_letter <- c()

# make vector with the names of sheets in the sample map
sheet_names <- c('Plasma_V1',																										
                 'Plasma_V2',
                 'Plasma_V2_temp',
                 'Plasma_V3',
                 'Plasma_V4',
                 'Plasma_V5',
                 'Urine_V1',
                 'Urine_V2',
                 'Urine_V3',
                 'Urine_V4',
                 'Urine_V5',
                 'Saliva_V1',
                 'Saliva_V2',
                 'Saliva_V3',
                 'Saliva_V4',
                 'WB_V1',
                 'WB_V2',
                 'WB_V3',
                 'WB_V4',
                 'WB_V5',
                 'Aprotenin',
                 'LPS_V1',
                 'LPS_V2',
                 'LPS_V3',
                 'LPS_V4',
                 'LPS_V5',
                 'LPS_returned_aliquots')
								 
# iterate through each sheet in the sample map
for (x in sheet_names) {																												
	i <- 0
	
	# make new dataframe called x for each sheet
	assign(x, read_excel('biobank_boxes_10_25_22.xlsx', sheet = x))
	
	# iterate through each row of x
	while (i <= nrow(eval(parse(text = x)))) {
		level <- 0
	
		# locate each boxkey in x and define it as a 10x10 dataframe
		boxkey = eval(parse(text = x))[(i+3):(i+12), 2:11]
		
		# identify boxkey number, skip rows that don't have a boxkey
		boxkey_num <- as.numeric(unname(unlist(eval(parse(text = x))[(i+1),1])))
		
		# make sure there is a boxkey here
		if (is.na(boxkey_num) == FALSE) {
		numberOfBoxkeys <- numberOfBoxkeys + 1
			for (a in 1:nrow(boxkey)) {
				for (b in 1:ncol(boxkey)) {
				
					# ignore slots without samples
					if (is.na(unname(unlist(boxkey[a,b]))) == FALSE) {
						
						# add the sample name to a vector
						ID_visit <- append(ID_visit, unname(unlist(boxkey[a,b])))
						
						# add the row to a vector
						row <- append(row, a)
						
						# add the column to a vector
						column <- append(column, b)
						
						# add the boxkey number to a vector
						boxkey_nums <- append(boxkey_nums, boxkey_num)
						
						# add tissue type to a vector
						tissue <- append(tissue, x)
						
						# add box letter
						box_letter <- append(box_letter, unname(unlist(eval(parse(text = x))[(i+1),5])))
						
						# add concentration for LPS
						if (x == 'LPS_V1' | x == 'LPS_V2' | x == 'LPS_V3' | x == 'LPS_V4' | x == 'LPS_V5') {
							LPS_concentration <- append(LPS_concentration, level)
							if (level < 3) {
								level <- level + 1
							} else level <- 0
						} else LPS_concentration <- append(LPS_concentration, NA)
					}
				}
			}
		}
		# move to the next boxkey
		i = i + 14
	}
	i <- 0
  
	# only some sheets have another 'column' of boxkeys
	if (x != 'Plasma_V2_temp') {
		while (i < nrow(eval(parse(text = x)))) {
			level <- 0
			boxkey = eval(parse(text = x))[(i+3):(i+12), 14:23]
			boxkey_num <- as.numeric(unname(unlist(eval(parse(text = x))[(i+1),13])))
			if (is.na(boxkey_num) == FALSE) {
				numberOfBoxkeys <- numberOfBoxkeys + 1
				for (a in 1:nrow(boxkey)) {
					for (b in 1:ncol(boxkey)) {
						if (is.na(unname(unlist(boxkey[a,b]))) == FALSE) {
							ID_visit <- append(ID_visit, unname(unlist(boxkey[a,b])))
							row <- append(row, a)
							column <- append(column, b)
							boxkey_nums <- append(boxkey_nums, boxkey_num)
							tissue <- append(tissue, x)
							box_letter <- append(box_letter, unname(unlist(eval(parse(text = x))[(i+1),17])))
							if (x == 'LPS_V1' | x == 'LPS_V2' | x == 'LPS_V4' | x == 'LPS_V5') {
								LPS_concentration <- append(LPS_concentration, level)
								if (level < 3) {
									level <- level + 1
								} else level <- 0
							} else LPS_concentration <- append(LPS_concentration, NA)
						}
					}
				}
			}
			i = i + 14
		}
	}
}

cleanID <- ID_visit

# format ids to get rid of spaces and astericks
cleanID <- gsub(' - ', '-', cleanID)
cleanID <- gsub('  - ', '-', cleanID)
cleanID <- gsub('- ', '-', cleanID)
cleanID <- gsub('-  ', '-', cleanID)
cleanID <- gsub(' -  ', '-', cleanID)
cleanID <- gsub(' -', '-', cleanID)
cleanID <- gsub('-  ', '-', cleanID)
cleanID <- gsub('\\*', '', cleanID)

# format ids to get rid of aliquot information
cleanID <- gsub('-1', '', cleanID)
cleanID <- gsub('-2', '', cleanID)
cleanID <- gsub('-3', '', cleanID)
cleanID <- gsub('-4', '', cleanID)
cleanID <- gsub('-5', '', cleanID)
cleanID <- gsub('-6', '', cleanID)
cleanID <- gsub('-7', '', cleanID)
cleanID <- gsub('-8', '', cleanID)

# format tissue types
tissue_temp <- tissue
tissue_temp <- gsub('_V1', '', tissue_temp)
tissue_temp <- gsub('_V2', '', tissue_temp)
tissue_temp <- gsub('_V3', '', tissue_temp)
tissue_temp <- gsub('_V4', '', tissue_temp)
tissue_temp <- gsub('_V5', '', tissue_temp)
tissue_temp <- gsub('_V2_temp', '', tissue_temp)

# make superID
superID <- paste(boxkey_nums, cleanID, box_letter, tissue_temp, row, column, sep = '_')

# write a dataframe with the desired columns and make a new excel spreadsheet
df <- data.frame(boxkey_nums, box_letter, tissue, cleanID, LPS_concentration, row, column, superID)
names(df) = c('boxkey', 'box letter', 'tissue', 'ID_visit', 'LPS', 'row', 'column', 'superID')

write.xlsx(df, 'C:/Users/anura/Documents/College/Nievergelt Lab/new sheet.xlsx', colNames = TRUE)
#write.xlsx(df, 'F:/new sheet.xlsx', colNames = TRUE)