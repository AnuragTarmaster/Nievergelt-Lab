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
                 'LPS_V5')
								 
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
						box_letter <- append(box_letter, 'A')
						
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
	if (x != 'Plasma_V2_temp' & x != 'WB_V1' & x != 'LPS_V3') {
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
							box_letter <- append(box_letter, 'B')
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
# write a dataframe with the desired columns and make a new excel spreadsheet
df <- data.frame(boxkey_nums, box_letter, tissue, ID_visit, LPS_concentration, row, column)
names(df) = c('boxkey', 'box letter', 'tissue', 'ID_visit', 'LPS', 'row', 'column')
write.xlsx(df,'C:/Users/anura/Documents/College/Nievergelt Lab/new sheet.xlsx',colNames = TRUE)