#' Creates a formatted data set in Excel.
#' @export
#Dataset = function(Dataset, Output)
{}

#' Creates a cross-tabulation table in Excel and Word.
#' @export
#Table = function(Dataset, Template, Output, Export)
  {
  
  #if(missing(Export)){Export = TRUE}
  #if(missing(Outputs)){Outputs = ""}
  
    {
      
      Group1 = GroupDetails1[1]
      Time1 = GroupDetails1[2]
      Weight1 = GroupDetails1[3]
      
      Group2 = GroupDetails2[1]
      Time2 = GroupDetails2[2]
      Weight2 = GroupDetails2[3]
      
      File = paste0(".", File)
      
      # Imports codebook.
      if(any(class(try((suppressMessages(read_excel(File, (which(readxl::excel_sheets(File) == "Codebook")), col_names = TRUE))),silent = TRUE)) == "try-error")){
        stop("No codebook found. Ensure there is a sheet named 'Codebook' in dataset file.")
      }
      Codebook = suppressMessages(read_excel(File, (which(readxl::excel_sheets(File) == "Codebook")), col_names = TRUE))
      
      # Gets the name of the datasets.
      Name_Group1 = paste(Group1, "at", Time1)
      Name_Group2 = paste(Group2, "at", Time2)
      
      # Checks if selected dataset exists.
      if(any(class(try((suppressMessages(read_excel(File, (which(readxl::excel_sheets(File) == Name_Group1)), col_names = TRUE))),silent = TRUE)) == "try-error")){
        stop(paste("No dataset with the name", Name_Group1, "found.", sep = " "))
      }
      if(any(class(try((suppressMessages(read_excel(File, (which(readxl::excel_sheets(File) == Name_Group2)), col_names = TRUE))),silent = TRUE)) == "try-error")){
        stop(paste("No dataset with the name", Name_Group2, "found.", sep = " "))
      }
      
      # Gets datasets.
      Dataset_Group1 = suppressMessages(read_excel(File, (which(readxl::excel_sheets(File) == Name_Group1)), col_names = TRUE))
      Dataset_Group2 = suppressMessages(read_excel(File, (which(readxl::excel_sheets(File) == Name_Group2)), col_names = TRUE))
      
      # Gets the name depending on the inputs.
      Name_Group = "Null"
      if(Group1 == Group2){
        Name_Group = paste(Group1, "at", Time1, "v.", Time2)
      }
      if(Time1 == Time2){
        Name_Group = paste(Group1, "v.", Group2, "at", Time1)
      }
      if((Group1 == Group2)&(Time1 == Time2)){
        Name_Group = paste(Group1, "at", Time1)
      }
      if((Group1!= Group2)&(Time1!= Time2)){
        Name_Group = paste(Group1, "at", Time1, "v.", Group2, "at", Time1)
      }
      
      # Returns a list of the dataframes and names.
      return(list(Codebook, Dataset_Group1, Dataset_Group2, Name_Group1, Name_Group2, Name_Group, Group1, Group2, Time1, Time2, Weight1, Weight2))
    }
  
  # Gets the codebook, datasets, and names in a list.
  List_Datasets_Names = Dataset_Import(Dataset, Group1, Group2)
  
  # Extracts the list elements.
  Codebook = (List_Datasets_Names[[1]])
  Dataset_Group1 = (List_Datasets_Names[[2]])
  Dataset_Group2 = (List_Datasets_Names[[3]])
  Name_Group1 = (List_Datasets_Names[[4]])
  Name_Group2 = (List_Datasets_Names[[5]])
  Name_Group = (List_Datasets_Names[[6]])
  Group1 = (List_Datasets_Names[[7]])
  Group2 = (List_Datasets_Names[[8]])
  Time1 = (List_Datasets_Names[[9]])
  Time2 = (List_Datasets_Names[[10]])
  Weight1 = (List_Datasets_Names[[11]])
  Weight2 = (List_Datasets_Names[[12]])
  
  # Adds overall column to codebook
  Overall = Codebook[1]
  Overall[] = NA
  Overall[1,] = "Ordinal"
  colnames(Overall) = "Overall"
  Codebook = cbind(Codebook, Overall)
  
  # Organizes the codebook by column value type (nominals, ordinals, etc.)
  Codebook = suppressWarnings(Codebook[,order(Codebook[1,])])
  
  # Adds overall column to group data
  Overall = Dataset_Group1[1]
  Overall[] = NA
  colnames(Overall) = "Overall"
  Dataset_Group1 = cbind(Dataset_Group1, Overall)
  Overall = Dataset_Group2[1]
  Overall[] = NA
  colnames(Overall) = "Overall"
  Dataset_Group2 = cbind(Dataset_Group2, Overall)
  
  Ordinals = which(Codebook[1,] == "Ordinal")
  
  for(Column in Columns){
    
  }
  
  # Values to stop crosstab creation once all nominals and ordinals have been analyzed
  OrdinalsCounter = Ordinal_Beg-1
  OrdinalsCounter_End = Ordinal_End
  
  # Values to stop crosstab creation once all nominals have been analyzed
  NominalsCounter_End = Nominal_End+1
  
  repeat{
    
    # Repeats crosstab creation if final ordinal category has not be completed.
    if(OrdinalsCounter == OrdinalsCounter_End){
      break
    }
    
    # Creates a dataset to contain crosstabs.
    Crosstabs = c()
    
    # Counts repetitions.
    OrdinalsCounter = OrdinalsCounter+1
    
    # Resets to first nominal.
    NominalsCounter = Nominal_Beg-1
    
    # Repeats crosstab creation for each nominal.
    repeat{
      
      # Counts repetitions.
      NominalsCounter = NominalsCounter+1
      
      # Ends crosstab creation if all nominals have been analyzed.
      if(NominalsCounter == NominalsCounter_End){
        # Ends repetitions.
        break
      }
      
      # Gets the column number of Group 1.
      ColumnNumber_Responses_Group1 = (which(names(Dataset_Group1) == colnames(Codebook[NominalsCounter])))
      
      # Gets the column number of the ordinals for Group 1.
      ColumnNumber_Ordinals_Group1 = (which(names(Dataset_Group1) == colnames(Codebook[OrdinalsCounter])))
      
      # Gets the response data from Group 1.
      Responses_Group1 = as.matrix(Dataset_Group1[ColumnNumber_Responses_Group1])
      
      # Gets the ordinal data for Group 1.
      Ordinals_Group1 = as.matrix(Dataset_Group1[ColumnNumber_Ordinals_Group1])
      
      # Gets numerical weights.
      if(Weight1 == "Unweighted"){
        Weight_Group1 = as.matrix(Dataset_Group1[(which(names(Dataset_Group1) == "Overall"))])
        Weight_Group1[] = 1
      }
      if(Weight1!= "Unweighted"){
        
        # Check to see if the weights exist
        Input_Test(Dataset_Group1, Weight1)
        
        # Gets the weight
        Weight_Group1 = as.matrix(Dataset_Group1[(which(names(Dataset_Group1) == Weight1))])
      }
      if(Weight2 == "Unweighted"){
        Weight_Group2 = as.matrix(Dataset_Group2[(which(names(Dataset_Group2) == "Overall"))])
        Weight_Group2[] = 1
      }
      if(Weight2!= "Unweighted"){
        
        # Check to see if the weights exist
        Input_Test(Dataset_Group2, Weight2)
        
        # Gets the weight
        Weight_Group2 = as.matrix(Dataset_Group2[(which(names(Dataset_Group2) == Weight2))])
      }
      
      # Gets the column number of Group 2.
      ColumnNumber_Responses_Group2 = (which(names(Dataset_Group2) == colnames(Codebook[NominalsCounter])))
      
      # Gets the column number of the ordinals for Group 2.
      ColumnNumber_Ordinals_Group2 = (which(names(Dataset_Group2) == colnames(Codebook[OrdinalsCounter])))
      
      # Gets the response data from Group 2.
      Responses_Group2 = as.matrix(Dataset_Group2[ColumnNumber_Responses_Group2])
      
      # Gets the ordinal data for Group 2.
      Ordinals_Group2 = as.matrix(Dataset_Group2[ColumnNumber_Ordinals_Group2])
      
      # Determines if the ID numbers are the same and, if they are, the data is set to paired.
      Paired = FALSE
      if(isTRUE(all.equal(Dataset_Group1$`Identification Number`, Dataset_Group2$`Identification Number`))){
        Paired = TRUE
      }
      
      # Gets the name of the ordinal category
      Ordinal_Category = colnames(Ordinals_Group1)
      
      # Gets the levels for the given nominal.
      Levels_Responses = unlist(Codebook[4:6,which(names(Codebook) == colnames(Responses_Group1))])
      
      # Gets the levels for the ordinal category.
      Levels_Ordinals = unlist(Codebook[which(names(Codebook) == Ordinal_Category)])
      Levels_Ordinals = seq(from = 1, to = length(Levels_Ordinals[complete.cases(Levels_Ordinals)])-1)
      if(isTRUE(all.equal(Levels_Ordinals, seq(from = 1, to = 0)))){
        Levels_Ordinals = NA
      }
      
      # Gets the nominal number for the corresponding responses.
      NominalNumber_Group1 = noquote(colnames(Dataset_Group1)[ColumnNumber_Responses_Group1])
      NominalNumber_Group2 = noquote(colnames(Dataset_Group2)[ColumnNumber_Responses_Group2])
      
      # Converts responses to text.
      Responses_Group1 = Responses_to_Text(Responses_Group1, NominalNumber_Group1, Codebook)
      Responses_Group2 = Responses_to_Text(Responses_Group2, NominalNumber_Group2, Codebook)
      
      # Assigns objects to global environment to accommodate svytable
      assign("Responses_Group1", Responses_Group1, envir = globalenv())
      assign("Responses_Group2", Responses_Group2, envir = globalenv())
      assign("Levels_Responses", Levels_Responses, envir = globalenv())
      assign("Ordinals_Group1", Ordinals_Group1, envir = globalenv())
      assign("Ordinals_Group2", Ordinals_Group2, envir = globalenv())
      assign("Levels_Ordinals", Levels_Ordinals, envir = globalenv())
      
      # Generates crosstabs.
      Crosstab_Group1 = 100*prop.table(svytable(~addNA(factor(Responses_Group1, ordered = T, levels = Levels_Responses))+factor(Ordinals_Group1, ordered = T, levels = Levels_Ordinals), svydesign(ids = ~1, data = as.data.frame(cbind(addNA(factor(Responses_Group1, ordered = T, levels = Levels_Responses)), factor(Ordinals_Group1, ordered = T, levels = Levels_Ordinals), Weight_Group1)), weights = ~Weight_Group1)), 2)
      Crosstab_Group2 = 100*prop.table(svytable(~addNA(factor(Responses_Group2, ordered = T, levels = Levels_Responses))+factor(Ordinals_Group2, ordered = T, levels = Levels_Ordinals), svydesign(ids = ~1, data = as.data.frame(cbind(addNA(factor(Responses_Group2, ordered = T, levels = Levels_Responses)), factor(Ordinals_Group2, ordered = T, levels = Levels_Ordinals), Weight_Group2)), weights = ~Weight_Group2)), 2)
      
      # If all responses are NA, then the crosstab is blanked
      if(all(na.exclude(Crosstab_Group1[nrow(Crosstab_Group1),] == 100))){Crosstab_Group1[] = NA}
      if(all(na.exclude(Crosstab_Group2[nrow(Crosstab_Group2),] == 100))){Crosstab_Group2[] = NA}
      
      # Gets the crosstab of the differences
      Crosstab_Difference = Crosstab_Group2-Crosstab_Group1
      
      # Formats crosstabs.
      Crosstab_Group1 = as.data.frame.matrix(format(round(Crosstab_Group1, digits = 1), nsmall = 1))
      Crosstab_Group2 = as.data.frame.matrix(format(round(Crosstab_Group2, digits = 1), nsmall = 1))
      Crosstab_Difference = as.data.frame.matrix(format(round(Crosstab_Difference, digits = 1), nsmall = 1))
      
      # Adds percentage symbol "%" to crosstab entries.
      if(ncol(Crosstab_Group1)>0){
        Excerpt1 = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, Crosstab_Group1, ""))
        Excerpt2 = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, Crosstab_Group2, ""))
        Excerpt3 = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, Crosstab_Difference, ""))
        rownames(Excerpt1) = rownames(Crosstab_Group1)
        rownames(Excerpt2) = rownames(Crosstab_Group2)
        rownames(Excerpt3) = rownames(Crosstab_Difference)
        Crosstab_Group1 = Excerpt1
        Crosstab_Group2 = Excerpt2
        Crosstab_Difference = Excerpt3
      }
      
      # Gets nominal number.
      NominalNumberText = noquote(as.character(NominalNumber_Group1))
      
      # Gets the text of the nominal.
      NominalText = as.character(Codebook[2, (which(names(Codebook) == as.character(NominalNumber_Group1)))])
      
      # Replaces "." with spaces.
      rownames(Crosstab_Group1) = gsub(".", " ", rownames(Crosstab_Group1), fixed = T)
      rownames(Crosstab_Group2) = gsub(".", " ", rownames(Crosstab_Group2), fixed = T)
      colnames(Crosstab_Group1) = gsub(".", " ", colnames(Crosstab_Group1), fixed = T)
      colnames(Crosstab_Group2) = gsub(".", " ", colnames(Crosstab_Group2), fixed = T)
      
      # Creates a vertical header.
      Column1 = c(NominalNumberText[1], 9999, 9999, 9999, 9999)
      Column2 = c(NominalText, (noquote(rownames(Crosstab_Group1)[1])), (noquote(rownames(Crosstab_Group1)[2])), (noquote(rownames(Crosstab_Group1)[3])), as.character(Codebook[7, (which(names(Codebook) == as.character(NominalNumber_Group1)))]))
      RowHeaders = data.frame(cbind(Column1, Column2))
      colnames(RowHeaders) = c("ID", "Prompt and Responses")
      
      # Generates total column.
      Total_Column_Group1 = matrix(data = "T", nrow = length(Responses_Group1), ncol = 1)
      Total_Column_Group2 = matrix(data = "T", nrow = length(Responses_Group2), ncol = 1)
      
      # Assigns objects to global environment to accommodate svytable
      assign("Total_Column_Group1", Total_Column_Group1, envir = globalenv())
      assign("Total_Column_Group2", Total_Column_Group2, envir = globalenv())
      
      # Generates crosstabs if there are responses
      Totals_Group1 = 100*prop.table(svytable(~addNA(factor(Responses_Group1, ordered = T, levels = Levels_Responses))+factor(Total_Column_Group1), svydesign(ids = ~1, data = as.data.frame(cbind(addNA(factor(Responses_Group1, ordered = T, levels = Levels_Responses)), factor(Total_Column_Group1), Weight_Group1)), weights = ~Weight_Group1)), 2)
      Totals_Group2 = 100*prop.table(svytable(~addNA(factor(Responses_Group2, ordered = T, levels = Levels_Responses))+factor(Total_Column_Group2), svydesign(ids = ~1, data = as.data.frame(cbind(addNA(factor(Responses_Group2, ordered = T, levels = Levels_Responses)), factor(Total_Column_Group2), Weight_Group2)), weights = ~Weight_Group2)), 2)
      
      # If all responses are NA, then the crosstab is blanked
      if(Totals_Group1[nrow(Totals_Group1)] == 100){Totals_Group1[] = NA}
      if(Totals_Group2[nrow(Totals_Group2)] == 100){Totals_Group2[] = NA}
      
      # Gets the crosstab of the differences
      Totals_Difference = Totals_Group2-Totals_Group1
      
      # Formats totals.
      Totals_Group1 = as.data.frame.matrix(format(round(Totals_Group1, digits = 1), nsmall = 1))
      Totals_Group2 = as.data.frame.matrix(format(round(Totals_Group2, digits = 1), nsmall = 1))
      Totals_Difference = as.data.frame.matrix(format(round(Totals_Difference, digits = 1), nsmall = 1))
      
      # Adds percentage symbol "%" to total column entries.
      Totals_Group1 = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, Totals_Group1, ""))
      Totals_Group2 = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, Totals_Group2, ""))
      Totals_Difference = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, Totals_Difference, ""))
      
      # Adds total column to crosstabs.
      Crosstab_Group1 = cbind(Totals_Group1, Crosstab_Group1)
      Crosstab_Group2 = cbind(Totals_Group2, Crosstab_Group2)
      Crosstab_Difference = cbind(Totals_Difference, Crosstab_Difference)
      
      # Replaces "." with spaces.
      rownames(Crosstab_Difference) = gsub(".", " ", rownames(Crosstab_Difference), fixed = T)
      colnames(Crosstab_Difference) = gsub(".", " ", colnames(Crosstab_Difference), fixed = T)
      
      # Ensures the number of rows of the crosstabs is 4.
      NumberofRowsis4 = TRUE
      if(nrow(Crosstab_Group1) == 3){
        Spacer = data.frame(matrix(data = 9999, nrow = 1, ncol = ncol(Crosstab_Group1)))
        rownames(Spacer) = 9999
        colnames(Spacer) = colnames(Crosstab_Group1)
        Crosstab_Group1 = rbind(Crosstab_Group1, Spacer)
        Crosstab_Group2 = rbind(Crosstab_Group2, Spacer)
        Crosstab_Difference = rbind(Crosstab_Difference, Spacer)
        NumberofRowsis4 = FALSE
      }
      
      # Adds a spacer between the crosstabs for both groups.
      Spacer = c(9999, 9999, 9999, 9999)
      Crosstab_Proportions = cbind(Crosstab_Group1, Spacer, Crosstab_Group2, Spacer, Crosstab_Difference)
      
      # Creates spacers to fill in with means and no nominal data.
      Spacer = data.frame(matrix(data = 9999, nrow = 1, ncol = 1))
      colnames(Spacer) = NA
      Means_Group1 = Means_Group2 = Means_Difference = NoNominals_Group1 = NoNominals_Group2 = NoNominals_Difference = data.frame(matrix(data = 9999, nrow = 1, ncol = length(Crosstab_Group1)))
      
      # Gets the numeric response data.
      Responses_Group1 = Dataset_Group1[(which(names(Dataset_Group1) == as.character(NominalNumber_Group1)))]
      Responses_Group2 = Dataset_Group2[(which(names(Dataset_Group2) == as.character(NominalNumber_Group2)))]
      
      # Gets the weighted NA raw percentage.
      {
        Responses_Group_NA1 = Responses_Group1
        Responses_Group_NA2 = Responses_Group2
        
        # Assigns objects to global environment to accommodate svytable
        assign("Responses_Group_NA1", Responses_Group_NA1, envir = globalenv())
        assign("Responses_Group_NA2", Responses_Group_NA2, envir = globalenv())
        
        # TRUE if responses are all NAs, indicating no responses for the given prompt.
        AllNAs_1 = (length((Responses_Group_NA1)[(is.na(as.matrix(Responses_Group_NA1)))]) == nrow((Responses_Group_NA1)))
        AllNAs_2 = (length((Responses_Group_NA2)[(is.na(as.matrix(Responses_Group_NA2)))]) == nrow((Responses_Group_NA2)))
        
        # Gets the percentage of responses that were NA.
        {
          # Gets the percentage of responses that were NA for group 1.
          {
            # If all entries are NAs, then the percentage of responses that are NA is marked as "NaN"
            if(AllNAs_1){
              NoNominal_Group1 = "NaN"
            }
            # If not all entries are NAs, then the percentage of responses that are NA is calculated.
            if(isFALSE(AllNAs_1)){
              
              # If no entries are NAs, then the percentage of responses that are NA is 0.
              if(isFALSE(sum(is.na(Responses_Group_NA1))>0)){
                NoNominal_Group1 = 0
              }
              
              # If some entries are NAs, then the percentage of responses that are NA is 0.
              if((sum(is.na(Responses_Group_NA1))>0)){
                Responses_Group_NA1[is.na(Responses_Group_NA1)] = -99
                
                # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                if(length(unlist(Responses_Group_NA1)) == 1){NoNominal_Group1 = 100}
                
                # Calculates the weighted percentage of responses that were NA
                if(isFALSE(length(unlist(Responses_Group_NA1)) == 1)){
                  NoNominal_Group1 = 100*prop.table(svytable(~unlist(Responses_Group_NA1), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA1), unlist(Weight_Group1))), weights = ~unlist(Weight_Group1))))[1]
                }}
            }
          }
          
          # Gets the percentage of responses that were NA for group 2.
          {
            # If all entries are NAs, then the percentage of responses that are NA is marked as "NaN"
            if(AllNAs_2){
              NoNominal_Group2 = "NaN"
            }
            # If not all entries are NAs, then the percentage of responses that are NA is calculated.
            if(isFALSE(AllNAs_2)){
              
              # If no entries are NAs, then the percentage of responses that are NA is 0.
              if(isFALSE(sum(is.na(Responses_Group_NA2))>0)){
                NoNominal_Group2 = 0
              }
              
              # If some entries are NAs, then the percentage of responses that are NA is 0.
              if((sum(is.na(Responses_Group_NA2))>0)){
                Responses_Group_NA2[is.na(Responses_Group_NA2)] = -99
                
                # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                if(length(unlist(Responses_Group_NA2)) == 1){NoNominal_Group2 = 100}
                
                # Calculates the weighted percentage of responses that were NA
                if(isFALSE(length(unlist(Responses_Group_NA2)) == 1)){
                  NoNominal_Group2 = 100*prop.table(svytable(~unlist(Responses_Group_NA2), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA2), unlist(Weight_Group2))), weights = ~unlist(Weight_Group2))))[1]
                }}
            }
          }
          
          # Gets the percentage of responses that were NA for the difference between group 1 and 2.
          {
            # If group 1 or 2 entries are all NA, then the difference in the percentage of responses that are NA is NaN.
            if(AllNAs_1|AllNAs_2){
              NoNominal_Difference = "NaN"
            }
            
            # If neither group 1 nor 2 entries are all NA, then the difference in the percentage of responses that are NA is calculated.
            if(isFALSE(AllNAs_1|AllNAs_2)){
              NoNominal_Difference = NoNominal_Group2-NoNominal_Group1
              NoNominal_Difference = (format(round(NoNominal_Difference, digits = 1), nsmall = 1))
            }
            
            if(isFALSE(AllNAs_1)){
              # Formats the percentage.
              NoNominal_Group1 = (format(round(NoNominal_Group1, digits = 1), nsmall = 1))
            }
            
            if(isFALSE(AllNAs_2)){
              # Formats the percentage.
              NoNominal_Group2 = (format(round(NoNominal_Group2, digits = 1), nsmall = 1))
            }
          }
        }
      }
      
      # If data is paired, removes NAs.
      if(Paired&isFALSE(AllNAs_1)&isFALSE(AllNAs_2)){
        # Removes entries for which there are missing responses in either groups.
        Responses_Group1_2 = cbind(Responses_Group1, Responses_Group2, Weight_Group1, Weight_Group2)
        Responses_Group1_2 = Responses_Group1_2[complete.cases(Responses_Group1_2),]
        
        # Inserts formatted data.
        Responses_Group1 = Responses_Group1_2[, 1]
        Responses_Group2 = Responses_Group1_2[, 2]
        Weight_Group1 = Responses_Group1_2[, 3]
        Weight_Group2 = Responses_Group1_2[, 4]
        
        # Unlists response data.
        Responses_Group1 = unlist(Responses_Group1)
        Responses_Group2 = unlist(Responses_Group2)
      }
      
      # Gets the means of the response data.
      Mean_Group1 = format(round(wtd.mean(Responses_Group1, Weight_Group1, na.rm = TRUE), digits = 3), nsmall = 3)
      Mean_Group2 = format(round(wtd.mean(Responses_Group2, Weight_Group2, na.rm = TRUE), digits = 3), nsmall = 3)
      Mean_Difference = format(round((wtd.mean(Responses_Group2, Weight_Group2, na.rm = TRUE)-wtd.mean(Responses_Group1, Weight_Group1, na.rm = TRUE)), digits = 3), nsmall = 3)
      
      # Performs a t-test if possible and adds data to crosstabs.
      Mean_Difference = Test_T(Mean_Difference, Responses_Group1, Responses_Group2, Weight_Group1, Weight_Group2, Paired)
      
      # Adds the no nominal data to dataframe.
      NoNominals_Group1[1] = noquote(as.character(NoNominal_Group1))
      NoNominals_Group2[1] = noquote(as.character(NoNominal_Group2))
      NoNominals_Difference[1] = noquote(as.character(NoNominal_Difference))
      
      # Adds percentage symbol "%" to no nominal entries.
      NoNominals_Group1[1] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoNominals_Group1[1], ""))
      NoNominals_Group2[1] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoNominals_Group2[1], ""))
      NoNominals_Difference[1] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoNominals_Difference[1], ""))
      
      # Adds the mean data to the dataframe.
      Means_Group1[1] = noquote(as.character(Mean_Group1))
      Means_Group2[1] = noquote(as.character(Mean_Group2))
      Means_Difference[1] = noquote(as.character(Mean_Difference))
      
      # Repeats creation of means and no nominal data for each ordinal category.
      CategoryCounter_End = length(Crosstab_Group1)
      CategoryCounter = 1
      repeat{
        CategoryCounter = 1+CategoryCounter
        if(CategoryCounter>CategoryCounter_End){
          break
        }
        count1001 = CategoryCounter
        
        # Extracts ordinal category.
        Category = colnames(Crosstab_Group1)[count1001]
        if(Category == "T"){
          Category = colnames(Crosstab_Group1)[count1001+1]
        }
        
        # Gets responses within each ordinal category.
        Selections_Group1 = Dataset_Group1[grepl((paste("^", Category, "$", sep = "")), Ordinals_Group1),]
        Selections_Group2 = Dataset_Group2[grepl((paste("^", Category, "$", sep = "")), Ordinals_Group2),]
        Responses_Group1 = Selections_Group1[(which(names(Selections_Group1) == as.character(NominalNumber_Group1)))]
        Responses_Group2 = Selections_Group2[(which(names(Selections_Group2) == as.character(NominalNumber_Group2)))]
        
        # Gets numerical weights.
        if(nrow(Responses_Group1)>0){
          if(Weight1 == "Unweighted"){
            Weight_Group1 = Selections_Group1[(which(names(Selections_Group1) == "Overall"))]
            Weight_Group1[] = 1
          }
          if(Weight1!= "Unweighted"){
            Weight_Group1 = Selections_Group1[(which(names(Selections_Group1) == as.character(Weight1)))]
          }
        } else {
          Weight_Group1 = Responses_Group1
        }
        if(nrow(Responses_Group2)>0){
          if(Weight2 == "Unweighted"){
            Weight_Group2 = Selections_Group2[(which(names(Selections_Group2) == "Overall"))]
            Weight_Group2[] = 1
          }
          if(Weight2!= "Unweighted"){
            Weight_Group2 = Selections_Group2[(which(names(Selections_Group2) == as.character(Weight2)))]
          }
        } else {
          Weight_Group2 = Responses_Group2
        }
        
        # TRUE if responses are all NAs (indicating respondents weren't surveyed).
        AllNAs_1 = (length(as.matrix(Responses_Group1)[(is.na(as.matrix(Responses_Group1)))]) == nrow(as.matrix(Responses_Group1)))
        AllNAs_2 = (length(as.matrix(Responses_Group2)[(is.na(as.matrix(Responses_Group2)))]) == nrow(as.matrix(Responses_Group2)))
        
        # Gets the weighted NA raw percentage.
        {
          Responses_Group_NA1 = Responses_Group1
          Responses_Group_NA2 = Responses_Group2
          
          # Assigns objects to global environment to accommodate svytable
          assign("Responses_Group_NA1", Responses_Group_NA1, envir = globalenv())
          assign("Responses_Group_NA2", Responses_Group_NA2, envir = globalenv())
          
          # TRUE if responses are all NAs, indicating no responses for the given prompt.
          AllNAs_1 = (length((Responses_Group_NA1)[(is.na(as.matrix(Responses_Group_NA1)))]) == nrow((Responses_Group_NA1)))
          AllNAs_2 = (length((Responses_Group_NA2)[(is.na(as.matrix(Responses_Group_NA2)))]) == nrow((Responses_Group_NA2)))
          
          # Gets the percentage of responses that were NA.
          {
            # Gets the percentage of responses that were NA for group 1.
            {
              # If all entries are NAs, then the percentage of responses that are NA is marked as "NaN"
              if(AllNAs_1){
                NoNominal_Group1 = "NaN"
              }
              # If not all entries are NAs, then the percentage of responses that are NA is calculated.
              if(isFALSE(AllNAs_1)){
                
                # If no entries are NAs, then the percentage of responses that are NA is 0.
                if(isFALSE(sum(is.na(Responses_Group_NA1))>0)){
                  NoNominal_Group1 = 0
                }
                
                # If some entries are NAs, then the percentage of responses that are NA is 0.
                if((sum(is.na(Responses_Group_NA1))>0)){
                  Responses_Group_NA1[is.na(Responses_Group_NA1)] = -99
                  
                  # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                  if(length(unlist(Responses_Group_NA1)) == 1){NoNominal_Group1 = 100}
                  
                  # Calculates the weighted percentage of responses that were NA
                  if(isFALSE(length(unlist(Responses_Group_NA1)) == 1)){
                    NoNominal_Group1 = 100*prop.table(svytable(~unlist(Responses_Group_NA1), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA1), unlist(Weight_Group1))), weights = ~unlist(Weight_Group1))))[1]
                  }}
              }
            }
            
            # Gets the percentage of responses that were NA for group 2.
            {
              # If all entries are NAs, then the percentage of responses that are NA is marked as "NaN"
              if(AllNAs_2){
                NoNominal_Group2 = "NaN"
              }
              # If not all entries are NAs, then the percentage of responses that are NA is calculated.
              if(isFALSE(AllNAs_2)){
                
                # If no entries are NAs, then the percentage of responses that are NA is 0.
                if(isFALSE(sum(is.na(Responses_Group_NA2))>0)){
                  NoNominal_Group2 = 0
                }
                
                # If some entries are NAs, then the percentage of responses that are NA is 0.
                if((sum(is.na(Responses_Group_NA2))>0)){
                  Responses_Group_NA2[is.na(Responses_Group_NA2)] = -99
                  
                  # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                  if(length(unlist(Responses_Group_NA2)) == 1){NoNominal_Group2 = 100}
                  
                  # Calculates the weighted percentage of responses that were NA
                  if(isFALSE(length(unlist(Responses_Group_NA2)) == 1)){
                    NoNominal_Group2 = 100*prop.table(svytable(~unlist(Responses_Group_NA2), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA2), unlist(Weight_Group2))), weights = ~unlist(Weight_Group2))))[1]
                  }}
              }
            }
            
            # Gets the percentage of responses that were NA for the difference between group 1 and 2.
            {
              # If group 1 or 2 entries are all NA, then the difference in the percentage of responses that are NA is NaN.
              if(AllNAs_1|AllNAs_2){
                NoNominal_Difference = "NaN"
              }
              
              # If neither group 1 nor 2 entries are all NA, then the difference in the percentage of responses that are NA is calculated.
              if(isFALSE(AllNAs_1|AllNAs_2)){
                NoNominal_Difference = NoNominal_Group2-NoNominal_Group1
                NoNominal_Difference = (format(round(NoNominal_Difference, digits = 1), nsmall = 1))
              }
              
              if(isFALSE(AllNAs_1)){
                # Formats the percentage.
                NoNominal_Group1 = (format(round(NoNominal_Group1, digits = 1), nsmall = 1))
              }
              
              if(isFALSE(AllNAs_2)){
                # Formats the percentage.
                NoNominal_Group2 = (format(round(NoNominal_Group2, digits = 1), nsmall = 1))
              }
            }
          }
        }
        
        # If data is paired, removes NAs.
        if(Paired&isFALSE(AllNAs_1)&isFALSE(AllNAs_2)){
          # Removes entries for which there are missing responses in either groups.
          Responses_Group1_2 = cbind(Responses_Group1, Responses_Group2, Weight_Group1, Weight_Group2)
          Responses_Group1_2 = Responses_Group1_2[complete.cases(Responses_Group1_2),]
          
          # Inserts formatted data.
          Responses_Group1 = Responses_Group1_2[, 1]
          Responses_Group2 = Responses_Group1_2[, 2]
          Weight_Group1 = Responses_Group1_2[, 3]
          Weight_Group2 = Responses_Group1_2[, 4]
          
          # Unlists response data.
          Responses_Group1 = unlist(Responses_Group1)
          Responses_Group2 = unlist(Responses_Group2)
        }
        
        # Gets the means of the response data.
        Mean_Group1 = format(round(wtd.mean(Responses_Group1, Weight_Group1, na.rm = TRUE), digits = 3), nsmall = 3)
        Mean_Group2 = format(round(wtd.mean(Responses_Group2, Weight_Group2, na.rm = TRUE), digits = 3), nsmall = 3)
        Mean_Difference = format(round((wtd.mean(Responses_Group2, Weight_Group2, na.rm = TRUE)-wtd.mean(Responses_Group1, Weight_Group1, na.rm = TRUE)), digits = 3), nsmall = 3)
        
        # Performs a t-test if possible and adds data to crosstabs.
        Mean_Difference = Test_T(Mean_Difference, Responses_Group1, Responses_Group2, Weight_Group1, Weight_Group2, Paired)
        
        # Adds the no nominal data to dataframe.
        NoNominals_Group1[count1001] = noquote(as.character(NoNominal_Group1))
        NoNominals_Group2[count1001] = noquote(as.character(NoNominal_Group2))
        NoNominals_Difference[count1001] = noquote(as.character(NoNominal_Difference))
        
        # Adds percentage symbol "%" to no nominal entries.
        NoNominalPercentageAdder = data.frame(matrix(data = "", nrow = 1, ncol = 1))
        NoNominals_Group1[count1001] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoNominals_Group1[count1001], NoNominalPercentageAdder))
        NoNominals_Group2[count1001] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoNominals_Group2[count1001], NoNominalPercentageAdder))
        NoNominals_Difference[count1001] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoNominals_Difference[count1001], NoNominalPercentageAdder))
        
        # Adds the mean data to the dataframe.
        Means_Group1[count1001] = noquote(as.character(Mean_Group1))
        Means_Group2[count1001] = noquote(as.character(Mean_Group2))
        Means_Difference[count1001] = noquote(as.character(Mean_Difference))
      }
      
      # Adds header to p-values dataframe.
      colnames(Means_Group1) = colnames(Means_Group2) = colnames(Means_Difference) = colnames(Crosstab_Group1)
      
      # Combines means and no nominal data.
      Crosstab_Means = cbind(Means_Group1, Spacer, Means_Group2, Spacer, Means_Difference)
      Crosstab_NoNominals = cbind(NoNominals_Group1, Spacer, NoNominals_Group2, Spacer, NoNominals_Difference)
      
      # Combines crosstab with means and no nominal.
      colnames(Crosstab_Proportions) = colnames(Crosstab_NoNominals) = colnames(Crosstab_Means)
      Crosstab = rbind(Crosstab_Means, Crosstab_Proportions)
      
      # Adds header to crosstab.
      Crosstab = cbind(RowHeaders, Crosstab)
      
      # Adds spacer for between different crosstabs.
      Spacer = data.frame(matrix(data = 9999, nrow = 1, ncol = length(Crosstab)))
      colnames(Spacer) = colnames(Crosstab)
      
      # Removes unneccesary row.
      if(isFALSE(NumberofRowsis4)){
        Crosstab[4,2] = Crosstab[5,2]
        Crosstab = Crosstab[-5,]
      }
      
      # Removes non-means and adds mean statistics if specified.
      if(Only_Means){
        Crosstab = Crosstab[-2:-nrow(Crosstab),]
        Spacer = c()
      }
      
      # Combines crosstabs to one dataframe.
      Crosstabs = rbind(Crosstabs, Spacer, Crosstab)
    }
    
    # Adds crosstab headers to dataframe.
    {
      ColumnHeaders = c(colnames(Crosstabs))
      Crosstabs = as.matrix(Crosstabs)
      Crosstabs = rbind(ColumnHeaders, Crosstabs)
      Crosstabs = data.frame(Crosstabs)
    }
    
    # Adds sample size statistics.
    {
      # Creates dataframes to hold sample size data.
      SampleSizes_Group1 = data.frame(matrix(data = 9999, nrow = 1, ncol = length(Crosstab_Group1)-1))
      SampleSizes_Group2 = data.frame(matrix(data = 9999, nrow = 1, ncol = length(Crosstab_Group2)-1))
      
      CategoryCounter_End = (length(Crosstab_Group1))
      CategoryCounter = 1
      repeat{
        CategoryCounter = 1+CategoryCounter
        if(CategoryCounter>CategoryCounter_End){
          break
        }
        
        count1001 = CategoryCounter
        
        # Gets a ordinal category.
        Category = colnames(Crosstab_Group1)[count1001]
        if(Category == "T"){
          Category = colnames(Crosstab_Group1)[count1001+1]
        }
        
        # Gets the responses corresponding with that ordinal category.
        Selections_Group1 = Dataset_Group1[grepl((paste("^", Category, "$", sep = "")), Ordinals_Group1),]
        Selections_Group2 = Dataset_Group2[grepl((paste("^", Category, "$", sep = "")), Ordinals_Group2),]
        Responses_Group1 = Selections_Group1[(which(names(Selections_Group1) == as.character(NominalNumber_Group1)))]
        Responses_Group2 = Selections_Group2[(which(names(Selections_Group2) == as.character(NominalNumber_Group2)))]
        
        # Unlists responses.
        Responses_Group1 = unlist(Responses_Group1)
        Responses_Group2 = unlist(Responses_Group2)
        
        # Gets the sample sizes and puts them in a text format with the proper format.
        SampleSizeNotation = noquote("n = ")
        SampleSize_Group1 = noquote(length(Responses_Group1))
        SampleSize_Group2 = noquote(length(Responses_Group2))
        SampleSize_Group1 = as.character(paste(SampleSizeNotation, SampleSize_Group1, sep = ""))
        SampleSize_Group2 = as.character(paste(SampleSizeNotation, SampleSize_Group2, sep = ""))
        
        # Puts entries into dataframe of sample sizes.
        SampleSizes_Group1[count1001-1] = SampleSize_Group1
        SampleSizes_Group2[count1001-1] = SampleSize_Group2
      }
      
      # Gets the sample sizes and puts them in a text format.
      SampleSizeNotation = noquote("n = ")
      SampleSize_Group1 = noquote(nrow(Dataset_Group1))
      SampleSize_Group2 = noquote(nrow(Dataset_Group2))
      SampleSize_Group1 = as.character(paste(SampleSizeNotation, SampleSize_Group1, sep = ""))
      SampleSize_Group2 = as.character(paste(SampleSizeNotation, SampleSize_Group2, sep = ""))
      
      # Puts entries into dataframe of sample sizes.
      TotalSize_Group1 = SampleSize_Group1
      TotalSize_Group2 = SampleSize_Group2
      
      # Adds sample size data to crosstabs.
      Spacer_1 = cbind(9999, 9999)
      Spacer_2 = cbind(9999)
      SampleSizes = cbind(Spacer_1, TotalSize_Group1, SampleSizes_Group1, Spacer_2, TotalSize_Group2, SampleSizes_Group2)
      Spacer_3 = data.frame(matrix(data = 9999, nrow = 1, ncol = (length(Crosstabs)-length(SampleSizes))))
      SampleSizes = cbind(SampleSizes, Spacer_3)
      Spacer = data.frame(matrix(data = 9999, nrow = 1, ncol = length(Crosstabs)))
      colnames(SampleSizes) = colnames(Crosstabs)
      colnames(Spacer) = colnames(Crosstabs)
      Crosstabs = rbind(SampleSizes, Spacer, Crosstabs)
      
      # Creates a header dataframe.
      ColumnHeaders = data.frame(matrix(data = 9999, nrow = 2, ncol = length(Crosstabs)))
      
      # Adds names of the groups to dataframe.
      ColumnHeaders[1, 3] = Name_Group1
      ColumnHeaders[1, (4+1*length(Crosstab_Group1))] = Name_Group2
      ColumnHeaders[1, (5+2*length(Crosstab_Group1))] = "Difference"
      
      # Adds dataframe to crosstabs.
      colnames(ColumnHeaders) = colnames(Crosstabs)
      Crosstabs = rbind(ColumnHeaders, Crosstabs)
    }
    
    # Creates the name and title of the files.
    {
      if(Weight1 == Weight2){Weights = Weight1}
      if(Weight1!= Weight2){Weights = paste(Weight1, Weight2, sep = " and ")}
      File_Name = paste("Responses", Name_Group, Weights, Ordinal_Category, sep = " - ")
      Document_Title = gsub(" by Overall", "", paste(paste(Name_Group, ", Weighted by ", Weights, sep = ""), Ordinal_Category, sep = " by "))
      Document_Title = gsub("Weighted by Unweighted", "Unweighted", Document_Title)
    }
    
    # Removes markers.
    {
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("matrix.data.....nrow...1..ncol...1.", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("NaN%", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("NaN", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("NA%", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("9999", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("Spacer", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("In.the.middle", "In the middle", x)}))
    }
    
    # Runs a pre-defined function that creates Word documents from the crosstabs.
    {
      Export_Word(Export, Outputs, Template, Crosstabs, Codebook, File_Name, Document_Title, Type = "Responses", Ordinal_Category)
    }
    
    # Adds a legend.
    {
      # Creates a legend.
      Legend = Codebook[-1, match(Ordinal_Category, names(Codebook))]
      Legend = as.data.frame(Legend[complete.cases(Legend)])
      Legend = rbind(c("T", "Total"), (cbind(rownames(Legend), Legend)))
      Legend = rbind(c("Legend", ""),Legend)
      
      # Adds the legend to the crosstabs
      Spacer = data.frame(matrix(data = 9999, ncol = (ncol(Crosstabs)-ncol(Legend)), nrow = nrow(Legend)))
      Legend = cbind(Legend, Spacer)
      Spacer = data.frame(matrix(data = 9999, ncol = ncol(Crosstabs), nrow = 1))
      colnames(Legend) = colnames(Spacer) = colnames(Crosstabs)
      Crosstabs = rbind(Legend, Spacer, Crosstabs)
    }
    
    # Removes markers.
    {
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("matrix.data.....nrow...1..ncol...1.", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("NaN%", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("NaN", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("NA%", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("9999", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("Spacer", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("In.the.middle", "In the middle", x)}))
    }
    
    # Runs a pre-defined function that creates Excel spreadsheets from the crosstabs.
    {
      Export_Excel(Outputs, Crosstabs, File_Name, Name_Group)
    }
  }
}