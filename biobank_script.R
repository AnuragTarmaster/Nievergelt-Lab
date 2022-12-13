setwd('C:/Users/anura/Downloads/Work Stuff')
library(readxl)
library(openxlsx)

id <- c()																																				# make vectors corresponding to columns in the long format
visit <- c()
freeze <- c()
box <- c()
box.letter <- c()
check <- c()
pulled <- c()
pulled_data <- c()
Requestor <- c()
project <- c()
returned <- c()
returned_aliquots_boxkey <- c()
total_N_aliquot <- c()
excluded <- c()

boxkey_nums <- c()
numberOfBoxkeys <- 0
ID_visit <- c()
row <- c()
column <- c()
tissue <- c()

sheet_names <- c('Plasma_V1',																										# make vector with the names of sheets in the sample map
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

for (x in sheet_names) {																												# iterate through each sheet in the sample map
  i <- 0
  assign(x, read_excel('biobank_boxes_10_25_22.xlsx', sheet = x))								# make new dataframe called x for each sheet
  
  while (i <= nrow(eval(parse(text = x)))) { 																		# iterate through each row of x
    boxkey = eval(parse(text = x))[(i+3):(i+12), 2:11] 													# locate each boxkey in x and define it as a 10x10 dataframe
    boxkey_num <- as.numeric(unname(unlist(eval(parse(text = x))[(i+1),1]))) 		# identify boxkey number, skip rows that don't have a boxkey
    
		if (is.na(boxkey_num) == FALSE) { 																					# make sure there is a boxkey here
	  numberOfBoxkeys <- numberOfBoxkeys + 1
      
			for (a in 1:nrow(boxkey)) {
        for (b in 1:ncol(boxkey)) {
				
          if (is.na(unname(unlist(boxkey[a,b]))) == FALSE) { 										# ignore slots without samples
            ID_visit <- append(ID_visit, unname(unlist(boxkey[a,b]))) 					# add the sample name to a vector
            row <- append(row, a)																								# add the row to a vector
            column <- append(column, b)																					# add the column to a vector
            boxkey_nums <- append(boxkey_nums, boxkey_num)											# add the boxkey number to a vector
					  tissue <- append(tissue, x)																					# add tissue type to a vector
          }
        }
      }
    }
    i = i + 14																																	# move to the next boxkey
  }
  
  i <- 0
  
  if (x != 'Plasma_V2_temp' & x != 'WB_V1' & x != 'LPS_V3') {										# only some sheets have another 'column' of boxkeys
    while (i < nrow(eval(parse(text = x)))) {
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
						}
          }
        }
      }
      i = i + 14
    }
  }
}
df <- data.frame(boxkey_nums, tissue, ID_visit, row, column)
write.xlsx(df,'C:/Users/anura/Downloads/Work Stuff/NEW_SHEET.xlsx',colNames = TRUE)
print('done loading')