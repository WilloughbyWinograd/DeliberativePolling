#' Creates a crosstab in Excel and Word.
#' @export
#' #Make final dataset SPSS. Make code run off SPSS
Crosstab = function(Dataset, Template, Output, Export){
  
  if(missing(Export)){Export = TRUE}
  if(missing(Outputs)){Outputs = ""}
  
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
  Overall[1,] = "Demographic"
  colnames(Overall) = "Overall"
  Codebook = cbind(Codebook, Overall)
  
  # Organizes the codebook by column value type (opinions, demographics, etc.)
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
  
  # Gets the column number of the first and last demographic and questions.
  Demographics = (which(Codebook[1,] == "Demographic"))
  Questions = (which(Codebook[1,] == "Opinion"))
  Demographic_Beg = Demographics[1]
  Demographic_End = Demographics[length(Demographics)]
  Question_Beg = Questions[1]
  Question_End = Questions[length(Questions)]
  
  # Values to stop crosstab creation once all questions and demographics have been analyzed
  DemographicsCounter = Demographic_Beg-1
  DemographicsCounter_End = Demographic_End
  
  # Values to stop crosstab creation once all questions have been analyzed
  QuestionsCounter_End = Question_End+1
  
  repeat{
    
    # Repeats crosstab creation if final demographic category has not be completed.
    if(DemographicsCounter == DemographicsCounter_End){
      break
    }
    
    # Creates a dataset to contain crosstabs.
    Crosstabs = c()
    
    # Counts repetitions.
    DemographicsCounter = DemographicsCounter+1
    
    # Resets to first question.
    QuestionsCounter = Question_Beg-1
    
    # Repeats crosstab creation for each question.
    repeat{
      
      # Counts repetitions.
      QuestionsCounter = QuestionsCounter+1
      
      # Ends crosstab creation if all questions have been analyzed.
      if(QuestionsCounter == QuestionsCounter_End){
        # Ends repetitions.
        break
      }
      
      # Gets the column number of Group 1.
      ColumnNumber_Responses_Group1 = (which(names(Dataset_Group1) == colnames(Codebook[QuestionsCounter])))
      
      # Gets the column number of the demographics for Group 1.
      ColumnNumber_Demographics_Group1 = (which(names(Dataset_Group1) == colnames(Codebook[DemographicsCounter])))
      
      # Gets the response data from Group 1.
      Responses_Group1 = as.matrix(Dataset_Group1[ColumnNumber_Responses_Group1])
      
      # Gets the demographic data for Group 1.
      Demographics_Group1 = as.matrix(Dataset_Group1[ColumnNumber_Demographics_Group1])
      
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
      ColumnNumber_Responses_Group2 = (which(names(Dataset_Group2) == colnames(Codebook[QuestionsCounter])))
      
      # Gets the column number of the demographics for Group 2.
      ColumnNumber_Demographics_Group2 = (which(names(Dataset_Group2) == colnames(Codebook[DemographicsCounter])))
      
      # Gets the response data from Group 2.
      Responses_Group2 = as.matrix(Dataset_Group2[ColumnNumber_Responses_Group2])
      
      # Gets the demographic data for Group 2.
      Demographics_Group2 = as.matrix(Dataset_Group2[ColumnNumber_Demographics_Group2])
      
      # Determines if the ID numbers are the same and, if they are, the data is set to paired.
      Paired = FALSE
      if(isTRUE(all.equal(Dataset_Group1$`Identification Number`, Dataset_Group2$`Identification Number`))){
        Paired = TRUE
      }
      
      # Gets the name of the demographic category
      Demographic_Category = colnames(Demographics_Group1)
      
      # Gets the levels for the given question.
      Levels_Responses = unlist(Codebook[4:6,which(names(Codebook) == colnames(Responses_Group1))])
      
      # Gets the levels for the demographic category.
      Levels_Demographics = unlist(Codebook[which(names(Codebook) == Demographic_Category)])
      Levels_Demographics = seq(from = 1, to = length(Levels_Demographics[complete.cases(Levels_Demographics)])-1)
      if(isTRUE(all.equal(Levels_Demographics, seq(from = 1, to = 0)))){
        Levels_Demographics = NA
      }
      
      # Gets the question number for the corresponding responses.
      QuestionNumber_Group1 = noquote(colnames(Dataset_Group1)[ColumnNumber_Responses_Group1])
      QuestionNumber_Group2 = noquote(colnames(Dataset_Group2)[ColumnNumber_Responses_Group2])
      
      # Converts responses to text.
      Responses_Group1 = Responses_to_Text(Responses_Group1, QuestionNumber_Group1, Codebook)
      Responses_Group2 = Responses_to_Text(Responses_Group2, QuestionNumber_Group2, Codebook)
      
      # Assigns objects to global environment to accommodate svytable
      assign("Responses_Group1", Responses_Group1, envir = globalenv())
      assign("Responses_Group2", Responses_Group2, envir = globalenv())
      assign("Levels_Responses", Levels_Responses, envir = globalenv())
      assign("Demographics_Group1", Demographics_Group1, envir = globalenv())
      assign("Demographics_Group2", Demographics_Group2, envir = globalenv())
      assign("Levels_Demographics", Levels_Demographics, envir = globalenv())
      
      # Generates crosstabs.
      Crosstab_Group1 = 100*prop.table(svytable(~addNA(factor(Responses_Group1, ordered = T, levels = Levels_Responses))+factor(Demographics_Group1, ordered = T, levels = Levels_Demographics), svydesign(ids = ~1, data = as.data.frame(cbind(addNA(factor(Responses_Group1, ordered = T, levels = Levels_Responses)), factor(Demographics_Group1, ordered = T, levels = Levels_Demographics), Weight_Group1)), weights = ~Weight_Group1)), 2)
      Crosstab_Group2 = 100*prop.table(svytable(~addNA(factor(Responses_Group2, ordered = T, levels = Levels_Responses))+factor(Demographics_Group2, ordered = T, levels = Levels_Demographics), svydesign(ids = ~1, data = as.data.frame(cbind(addNA(factor(Responses_Group2, ordered = T, levels = Levels_Responses)), factor(Demographics_Group2, ordered = T, levels = Levels_Demographics), Weight_Group2)), weights = ~Weight_Group2)), 2)
      
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
      
      # Gets question number.
      QuestionNumberText = noquote(as.character(QuestionNumber_Group1))
      
      # Gets the text of the question.
      QuestionText = as.character(Codebook[2, (which(names(Codebook) == as.character(QuestionNumber_Group1)))])
      
      # Replaces "." with spaces.
      rownames(Crosstab_Group1) = gsub(".", " ", rownames(Crosstab_Group1), fixed = T)
      rownames(Crosstab_Group2) = gsub(".", " ", rownames(Crosstab_Group2), fixed = T)
      colnames(Crosstab_Group1) = gsub(".", " ", colnames(Crosstab_Group1), fixed = T)
      colnames(Crosstab_Group2) = gsub(".", " ", colnames(Crosstab_Group2), fixed = T)
      
      # Creates a vertical header.
      Column1 = c(QuestionNumberText[1], 9999, 9999, 9999, 9999)
      Column2 = c(QuestionText, (noquote(rownames(Crosstab_Group1)[1])), (noquote(rownames(Crosstab_Group1)[2])), (noquote(rownames(Crosstab_Group1)[3])), as.character(Codebook[7, (which(names(Codebook) == as.character(QuestionNumber_Group1)))]))
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
      
      # Creates spacers to fill in with means and no opinion data.
      Spacer = data.frame(matrix(data = 9999, nrow = 1, ncol = 1))
      colnames(Spacer) = NA
      Means_Group1 = Means_Group2 = Means_Difference = NoOpinions_Group1 = NoOpinions_Group2 = NoOpinions_Difference = data.frame(matrix(data = 9999, nrow = 1, ncol = length(Crosstab_Group1)))
      
      # Gets the numeric response data.
      Responses_Group1 = Dataset_Group1[(which(names(Dataset_Group1) == as.character(QuestionNumber_Group1)))]
      Responses_Group2 = Dataset_Group2[(which(names(Dataset_Group2) == as.character(QuestionNumber_Group2)))]
      
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
              NoOpinion_Group1 = "NaN"
            }
            # If not all entries are NAs, then the percentage of responses that are NA is calculated.
            if(isFALSE(AllNAs_1)){
              
              # If no entries are NAs, then the percentage of responses that are NA is 0.
              if(isFALSE(sum(is.na(Responses_Group_NA1))>0)){
                NoOpinion_Group1 = 0
              }
              
              # If some entries are NAs, then the percentage of responses that are NA is 0.
              if((sum(is.na(Responses_Group_NA1))>0)){
                Responses_Group_NA1[is.na(Responses_Group_NA1)] = -99
                
                # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                if(length(unlist(Responses_Group_NA1)) == 1){NoOpinion_Group1 = 100}
                
                # Calculates the weighted percentage of responses that were NA
                if(isFALSE(length(unlist(Responses_Group_NA1)) == 1)){
                  NoOpinion_Group1 = 100*prop.table(svytable(~unlist(Responses_Group_NA1), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA1), unlist(Weight_Group1))), weights = ~unlist(Weight_Group1))))[1]
                }}
            }
          }
          
          # Gets the percentage of responses that were NA for group 2.
          {
            # If all entries are NAs, then the percentage of responses that are NA is marked as "NaN"
            if(AllNAs_2){
              NoOpinion_Group2 = "NaN"
            }
            # If not all entries are NAs, then the percentage of responses that are NA is calculated.
            if(isFALSE(AllNAs_2)){
              
              # If no entries are NAs, then the percentage of responses that are NA is 0.
              if(isFALSE(sum(is.na(Responses_Group_NA2))>0)){
                NoOpinion_Group2 = 0
              }
              
              # If some entries are NAs, then the percentage of responses that are NA is 0.
              if((sum(is.na(Responses_Group_NA2))>0)){
                Responses_Group_NA2[is.na(Responses_Group_NA2)] = -99
                
                # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                if(length(unlist(Responses_Group_NA2)) == 1){NoOpinion_Group2 = 100}
                
                # Calculates the weighted percentage of responses that were NA
                if(isFALSE(length(unlist(Responses_Group_NA2)) == 1)){
                  NoOpinion_Group2 = 100*prop.table(svytable(~unlist(Responses_Group_NA2), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA2), unlist(Weight_Group2))), weights = ~unlist(Weight_Group2))))[1]
                }}
            }
          }
          
          # Gets the percentage of responses that were NA for the difference between group 1 and 2.
          {
            # If group 1 or 2 entries are all NA, then the difference in the percentage of responses that are NA is NaN.
            if(AllNAs_1|AllNAs_2){
              NoOpinion_Difference = "NaN"
            }
            
            # If neither group 1 nor 2 entries are all NA, then the difference in the percentage of responses that are NA is calculated.
            if(isFALSE(AllNAs_1|AllNAs_2)){
              NoOpinion_Difference = NoOpinion_Group2-NoOpinion_Group1
              NoOpinion_Difference = (format(round(NoOpinion_Difference, digits = 1), nsmall = 1))
            }
            
            if(isFALSE(AllNAs_1)){
              # Formats the percentage.
              NoOpinion_Group1 = (format(round(NoOpinion_Group1, digits = 1), nsmall = 1))
            }
            
            if(isFALSE(AllNAs_2)){
              # Formats the percentage.
              NoOpinion_Group2 = (format(round(NoOpinion_Group2, digits = 1), nsmall = 1))
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
      
      # Adds the no opinion data to dataframe.
      NoOpinions_Group1[1] = noquote(as.character(NoOpinion_Group1))
      NoOpinions_Group2[1] = noquote(as.character(NoOpinion_Group2))
      NoOpinions_Difference[1] = noquote(as.character(NoOpinion_Difference))
      
      # Adds percentage symbol "%" to no opinion entries.
      NoOpinions_Group1[1] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Group1[1], ""))
      NoOpinions_Group2[1] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Group2[1], ""))
      NoOpinions_Difference[1] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Difference[1], ""))
      
      # Adds the mean data to the dataframe.
      Means_Group1[1] = noquote(as.character(Mean_Group1))
      Means_Group2[1] = noquote(as.character(Mean_Group2))
      Means_Difference[1] = noquote(as.character(Mean_Difference))
      
      # Repeats creation of means and no opinion data for each demographic category.
      CategoryCounter_End = length(Crosstab_Group1)
      CategoryCounter = 1
      repeat{
        CategoryCounter = 1+CategoryCounter
        if(CategoryCounter>CategoryCounter_End){
          break
        }
        count1001 = CategoryCounter
        
        # Extracts demographic category.
        Category = colnames(Crosstab_Group1)[count1001]
        if(Category == "T"){
          Category = colnames(Crosstab_Group1)[count1001+1]
        }
        
        # Gets responses within each demographic category.
        Selections_Group1 = Dataset_Group1[grepl((paste("^", Category, "$", sep = "")), Demographics_Group1),]
        Selections_Group2 = Dataset_Group2[grepl((paste("^", Category, "$", sep = "")), Demographics_Group2),]
        Responses_Group1 = Selections_Group1[(which(names(Selections_Group1) == as.character(QuestionNumber_Group1)))]
        Responses_Group2 = Selections_Group2[(which(names(Selections_Group2) == as.character(QuestionNumber_Group2)))]
        
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
                NoOpinion_Group1 = "NaN"
              }
              # If not all entries are NAs, then the percentage of responses that are NA is calculated.
              if(isFALSE(AllNAs_1)){
                
                # If no entries are NAs, then the percentage of responses that are NA is 0.
                if(isFALSE(sum(is.na(Responses_Group_NA1))>0)){
                  NoOpinion_Group1 = 0
                }
                
                # If some entries are NAs, then the percentage of responses that are NA is 0.
                if((sum(is.na(Responses_Group_NA1))>0)){
                  Responses_Group_NA1[is.na(Responses_Group_NA1)] = -99
                  
                  # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                  if(length(unlist(Responses_Group_NA1)) == 1){NoOpinion_Group1 = 100}
                  
                  # Calculates the weighted percentage of responses that were NA
                  if(isFALSE(length(unlist(Responses_Group_NA1)) == 1)){
                    NoOpinion_Group1 = 100*prop.table(svytable(~unlist(Responses_Group_NA1), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA1), unlist(Weight_Group1))), weights = ~unlist(Weight_Group1))))[1]
                  }}
              }
            }
            
            # Gets the percentage of responses that were NA for group 2.
            {
              # If all entries are NAs, then the percentage of responses that are NA is marked as "NaN"
              if(AllNAs_2){
                NoOpinion_Group2 = "NaN"
              }
              # If not all entries are NAs, then the percentage of responses that are NA is calculated.
              if(isFALSE(AllNAs_2)){
                
                # If no entries are NAs, then the percentage of responses that are NA is 0.
                if(isFALSE(sum(is.na(Responses_Group_NA2))>0)){
                  NoOpinion_Group2 = 0
                }
                
                # If some entries are NAs, then the percentage of responses that are NA is 0.
                if((sum(is.na(Responses_Group_NA2))>0)){
                  Responses_Group_NA2[is.na(Responses_Group_NA2)] = -99
                  
                  # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                  if(length(unlist(Responses_Group_NA2)) == 1){NoOpinion_Group2 = 100}
                  
                  # Calculates the weighted percentage of responses that were NA
                  if(isFALSE(length(unlist(Responses_Group_NA2)) == 1)){
                    NoOpinion_Group2 = 100*prop.table(svytable(~unlist(Responses_Group_NA2), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA2), unlist(Weight_Group2))), weights = ~unlist(Weight_Group2))))[1]
                  }}
              }
            }
            
            # Gets the percentage of responses that were NA for the difference between group 1 and 2.
            {
              # If group 1 or 2 entries are all NA, then the difference in the percentage of responses that are NA is NaN.
              if(AllNAs_1|AllNAs_2){
                NoOpinion_Difference = "NaN"
              }
              
              # If neither group 1 nor 2 entries are all NA, then the difference in the percentage of responses that are NA is calculated.
              if(isFALSE(AllNAs_1|AllNAs_2)){
                NoOpinion_Difference = NoOpinion_Group2-NoOpinion_Group1
                NoOpinion_Difference = (format(round(NoOpinion_Difference, digits = 1), nsmall = 1))
              }
              
              if(isFALSE(AllNAs_1)){
                # Formats the percentage.
                NoOpinion_Group1 = (format(round(NoOpinion_Group1, digits = 1), nsmall = 1))
              }
              
              if(isFALSE(AllNAs_2)){
                # Formats the percentage.
                NoOpinion_Group2 = (format(round(NoOpinion_Group2, digits = 1), nsmall = 1))
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
        
        # Adds the no opinion data to dataframe.
        NoOpinions_Group1[count1001] = noquote(as.character(NoOpinion_Group1))
        NoOpinions_Group2[count1001] = noquote(as.character(NoOpinion_Group2))
        NoOpinions_Difference[count1001] = noquote(as.character(NoOpinion_Difference))
        
        # Adds percentage symbol "%" to no opinion entries.
        NoOpinionPercentageAdder = data.frame(matrix(data = "", nrow = 1, ncol = 1))
        NoOpinions_Group1[count1001] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Group1[count1001], NoOpinionPercentageAdder))
        NoOpinions_Group2[count1001] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Group2[count1001], NoOpinionPercentageAdder))
        NoOpinions_Difference[count1001] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Difference[count1001], NoOpinionPercentageAdder))
        
        # Adds the mean data to the dataframe.
        Means_Group1[count1001] = noquote(as.character(Mean_Group1))
        Means_Group2[count1001] = noquote(as.character(Mean_Group2))
        Means_Difference[count1001] = noquote(as.character(Mean_Difference))
      }
      
      # Adds header to p-values dataframe.
      colnames(Means_Group1) = colnames(Means_Group2) = colnames(Means_Difference) = colnames(Crosstab_Group1)
      
      # Combines means and no opinion data.
      Crosstab_Means = cbind(Means_Group1, Spacer, Means_Group2, Spacer, Means_Difference)
      Crosstab_NoOpinions = cbind(NoOpinions_Group1, Spacer, NoOpinions_Group2, Spacer, NoOpinions_Difference)
      
      # Combines crosstab with means and no opinion.
      colnames(Crosstab_Proportions) = colnames(Crosstab_NoOpinions) = colnames(Crosstab_Means)
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
        
        # Gets a demographic category.
        Category = colnames(Crosstab_Group1)[count1001]
        if(Category == "T"){
          Category = colnames(Crosstab_Group1)[count1001+1]
        }
        
        # Gets the responses corresponding with that demographic category.
        Selections_Group1 = Dataset_Group1[grepl((paste("^", Category, "$", sep = "")), Demographics_Group1),]
        Selections_Group2 = Dataset_Group2[grepl((paste("^", Category, "$", sep = "")), Demographics_Group2),]
        Responses_Group1 = Selections_Group1[(which(names(Selections_Group1) == as.character(QuestionNumber_Group1)))]
        Responses_Group2 = Selections_Group2[(which(names(Selections_Group2) == as.character(QuestionNumber_Group2)))]
        
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
      File_Name = paste("Responses", Name_Group, Weights, Demographic_Category, sep = " - ")
      Document_Title = gsub(" by Overall", "", paste(paste(Name_Group, ", Weighted by ", Weights, sep = ""), Demographic_Category, sep = " by "))
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
      Export_Word(Export, Outputs, Template, Crosstabs, Codebook, File_Name, Document_Title, Type = "Responses", Demographic_Category)
    }
    
    # Adds a legend.
    {
      # Creates a legend.
      Legend = Codebook[-1, match(Demographic_Category, names(Codebook))]
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

#' Creates a crosstab in Excel and Word comparing opinions.
#' @export
Crosstab.Opinions = function(Dataset, Template, Group1, Group2, Outputs, Export, Alpha, Only_Significant, Only_Means){

  if(missing(Export)){Export = TRUE}
  if(missing(Outputs)){Outputs = ""}
  if(missing(Alpha)){Alpha = 0.05}
  if(missing(Only_Significant)){Only_Significant = FALSE}
  if(missing(Only_Means)){Only_Means = FALSE}

  # Adds objects to global environment
  assign("Alpha", Alpha, envir = globalenv())
  assign("Only_Significant", Only_Significant, envir = globalenv())
  assign("Only_Means", Only_Means, envir = globalenv())

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
  Overall[1,] = "Demographic"
  colnames(Overall) = "Overall"
  Codebook = cbind(Codebook, Overall)

  # Organizes the codebook by column value type (opinions, demographics, etc.)
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

  # Gets the column number of the first and last demographic and questions.
  Demographics = (which(Codebook[1,] == "Demographic"))
  Questions = (which(Codebook[1,] == "Opinion"))
  Demographic_Beg = Demographics[1]
  Demographic_End = Demographics[length(Demographics)]
  Question_Beg = Questions[1]
  Question_End = Questions[length(Questions)]

  # Values to stop crosstab creation once all questions and demographics have been analyzed
  DemographicsCounter = Demographic_Beg-1
  DemographicsCounter_End = Demographic_End

  # Values to stop crosstab creation once all questions have been analyzed
  QuestionsCounter_End = Question_End+1

  repeat{

    # Repeats crosstab creation if final demographic category has not be completed.
    if(DemographicsCounter == DemographicsCounter_End){
      break
    }

    # Creates a dataset to contain crosstabs.
    Crosstabs = c()

    # Counts repetitions.
    DemographicsCounter = DemographicsCounter+1

    # Resets to first question.
    QuestionsCounter = Question_Beg-1

    # Repeats crosstab creation for each question.
    repeat{

      # Counts repetitions.
      QuestionsCounter = QuestionsCounter+1

      # Ends crosstab creation if all questions have been analyzed.
      if(QuestionsCounter == QuestionsCounter_End){
        # Ends repetitions.
        break
      }

      # Gets the column number of Group 1.
      ColumnNumber_Responses_Group1 = (which(names(Dataset_Group1) == colnames(Codebook[QuestionsCounter])))

      # Gets the column number of the demographics for Group 1.
      ColumnNumber_Demographics_Group1 = (which(names(Dataset_Group1) == colnames(Codebook[DemographicsCounter])))

      # Gets the response data from Group 1.
      Responses_Group1 = as.matrix(Dataset_Group1[ColumnNumber_Responses_Group1])

      # Gets the demographic data for Group 1.
      Demographics_Group1 = as.matrix(Dataset_Group1[ColumnNumber_Demographics_Group1])

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
      ColumnNumber_Responses_Group2 = (which(names(Dataset_Group2) == colnames(Codebook[QuestionsCounter])))

      # Gets the column number of the demographics for Group 2.
      ColumnNumber_Demographics_Group2 = (which(names(Dataset_Group2) == colnames(Codebook[DemographicsCounter])))

      # Gets the response data from Group 2.
      Responses_Group2 = as.matrix(Dataset_Group2[ColumnNumber_Responses_Group2])

      # Gets the demographic data for Group 2.
      Demographics_Group2 = as.matrix(Dataset_Group2[ColumnNumber_Demographics_Group2])

      # Determines if the ID numbers are the same and, if they are, the data is set to paired.
      Paired = FALSE
      if(isTRUE(all.equal(Dataset_Group1$`Identification Number`, Dataset_Group2$`Identification Number`))){
        Paired = TRUE
      }

      # Gets the name of the demographic category
      Demographic_Category = colnames(Demographics_Group1)

      # Gets the levels for the given question.
      Levels_Responses = unlist(Codebook[4:6,which(names(Codebook) == colnames(Responses_Group1))])

      # Gets the levels for the demographic category.
      Levels_Demographics = unlist(Codebook[which(names(Codebook) == Demographic_Category)])
      Levels_Demographics = seq(from = 1, to = length(Levels_Demographics[complete.cases(Levels_Demographics)])-1)
      if(isTRUE(all.equal(Levels_Demographics, seq(from = 1, to = 0)))){
        Levels_Demographics = NA
      }

      # Gets the question number for the corresponding responses.
      QuestionNumber_Group1 = noquote(colnames(Dataset_Group1)[ColumnNumber_Responses_Group1])
      QuestionNumber_Group2 = noquote(colnames(Dataset_Group2)[ColumnNumber_Responses_Group2])

      # Converts responses to text.
      Responses_Group1 = Responses_to_Text(Responses_Group1, QuestionNumber_Group1, Codebook)
      Responses_Group2 = Responses_to_Text(Responses_Group2, QuestionNumber_Group2, Codebook)

      # Assigns objects to global environment to accommodate svytable
      assign("Responses_Group1", Responses_Group1, envir = globalenv())
      assign("Responses_Group2", Responses_Group2, envir = globalenv())
      assign("Levels_Responses", Levels_Responses, envir = globalenv())
      assign("Demographics_Group1", Demographics_Group1, envir = globalenv())
      assign("Demographics_Group2", Demographics_Group2, envir = globalenv())
      assign("Levels_Demographics", Levels_Demographics, envir = globalenv())

      # Generates crosstabs.
      Crosstab_Group1 = 100*prop.table(svytable(~addNA(factor(Responses_Group1, ordered = T, levels = Levels_Responses))+factor(Demographics_Group1, ordered = T, levels = Levels_Demographics), svydesign(ids = ~1, data = as.data.frame(cbind(addNA(factor(Responses_Group1, ordered = T, levels = Levels_Responses)), factor(Demographics_Group1, ordered = T, levels = Levels_Demographics), Weight_Group1)), weights = ~Weight_Group1)), 2)
      Crosstab_Group2 = 100*prop.table(svytable(~addNA(factor(Responses_Group2, ordered = T, levels = Levels_Responses))+factor(Demographics_Group2, ordered = T, levels = Levels_Demographics), svydesign(ids = ~1, data = as.data.frame(cbind(addNA(factor(Responses_Group2, ordered = T, levels = Levels_Responses)), factor(Demographics_Group2, ordered = T, levels = Levels_Demographics), Weight_Group2)), weights = ~Weight_Group2)), 2)

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

      # Gets question number.
      QuestionNumberText = noquote(as.character(QuestionNumber_Group1))

      # Gets the text of the question.
      QuestionText = as.character(Codebook[2, (which(names(Codebook) == as.character(QuestionNumber_Group1)))])

      # Replaces "." with spaces.
      rownames(Crosstab_Group1) = gsub(".", " ", rownames(Crosstab_Group1), fixed = T)
      rownames(Crosstab_Group2) = gsub(".", " ", rownames(Crosstab_Group2), fixed = T)
      colnames(Crosstab_Group1) = gsub(".", " ", colnames(Crosstab_Group1), fixed = T)
      colnames(Crosstab_Group2) = gsub(".", " ", colnames(Crosstab_Group2), fixed = T)

      # Creates a vertical header.
      Column1 = c(QuestionNumberText[1], 9999, 9999, 9999, 9999)
      Column2 = c(QuestionText, (noquote(rownames(Crosstab_Group1)[1])), (noquote(rownames(Crosstab_Group1)[2])), (noquote(rownames(Crosstab_Group1)[3])), as.character(Codebook[7, (which(names(Codebook) == as.character(QuestionNumber_Group1)))]))
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

      # Creates spacers to fill in with means and no opinion data.
      Spacer = data.frame(matrix(data = 9999, nrow = 1, ncol = 1))
      colnames(Spacer) = NA
      Means_Group1 = Means_Group2 = Means_Difference = NoOpinions_Group1 = NoOpinions_Group2 = NoOpinions_Difference = data.frame(matrix(data = 9999, nrow = 1, ncol = length(Crosstab_Group1)))

      # Gets the numeric response data.
      Responses_Group1 = Dataset_Group1[(which(names(Dataset_Group1) == as.character(QuestionNumber_Group1)))]
      Responses_Group2 = Dataset_Group2[(which(names(Dataset_Group2) == as.character(QuestionNumber_Group2)))]

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
              NoOpinion_Group1 = "NaN"
            }
            # If not all entries are NAs, then the percentage of responses that are NA is calculated.
            if(isFALSE(AllNAs_1)){

              # If no entries are NAs, then the percentage of responses that are NA is 0.
              if(isFALSE(sum(is.na(Responses_Group_NA1))>0)){
                NoOpinion_Group1 = 0
              }

              # If some entries are NAs, then the percentage of responses that are NA is 0.
              if((sum(is.na(Responses_Group_NA1))>0)){
                Responses_Group_NA1[is.na(Responses_Group_NA1)] = -99

                # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                if(length(unlist(Responses_Group_NA1)) == 1){NoOpinion_Group1 = 100}

                # Calculates the weighted percentage of responses that were NA
                if(isFALSE(length(unlist(Responses_Group_NA1)) == 1)){
                  NoOpinion_Group1 = 100*prop.table(svytable(~unlist(Responses_Group_NA1), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA1), unlist(Weight_Group1))), weights = ~unlist(Weight_Group1))))[1]
                }}
            }
          }

          # Gets the percentage of responses that were NA for group 2.
          {
            # If all entries are NAs, then the percentage of responses that are NA is marked as "NaN"
            if(AllNAs_2){
              NoOpinion_Group2 = "NaN"
            }
            # If not all entries are NAs, then the percentage of responses that are NA is calculated.
            if(isFALSE(AllNAs_2)){

              # If no entries are NAs, then the percentage of responses that are NA is 0.
              if(isFALSE(sum(is.na(Responses_Group_NA2))>0)){
                NoOpinion_Group2 = 0
              }

              # If some entries are NAs, then the percentage of responses that are NA is 0.
              if((sum(is.na(Responses_Group_NA2))>0)){
                Responses_Group_NA2[is.na(Responses_Group_NA2)] = -99

                # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                if(length(unlist(Responses_Group_NA2)) == 1){NoOpinion_Group2 = 100}

                # Calculates the weighted percentage of responses that were NA
                if(isFALSE(length(unlist(Responses_Group_NA2)) == 1)){
                  NoOpinion_Group2 = 100*prop.table(svytable(~unlist(Responses_Group_NA2), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA2), unlist(Weight_Group2))), weights = ~unlist(Weight_Group2))))[1]
                }}
            }
          }

          # Gets the percentage of responses that were NA for the difference between group 1 and 2.
          {
            # If group 1 or 2 entries are all NA, then the difference in the percentage of responses that are NA is NaN.
            if(AllNAs_1|AllNAs_2){
              NoOpinion_Difference = "NaN"
            }

            # If neither group 1 nor 2 entries are all NA, then the difference in the percentage of responses that are NA is calculated.
            if(isFALSE(AllNAs_1|AllNAs_2)){
              NoOpinion_Difference = NoOpinion_Group2-NoOpinion_Group1
              NoOpinion_Difference = (format(round(NoOpinion_Difference, digits = 1), nsmall = 1))
            }

            if(isFALSE(AllNAs_1)){
              # Formats the percentage.
              NoOpinion_Group1 = (format(round(NoOpinion_Group1, digits = 1), nsmall = 1))
            }

            if(isFALSE(AllNAs_2)){
              # Formats the percentage.
              NoOpinion_Group2 = (format(round(NoOpinion_Group2, digits = 1), nsmall = 1))
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

      # Adds the no opinion data to dataframe.
      NoOpinions_Group1[1] = noquote(as.character(NoOpinion_Group1))
      NoOpinions_Group2[1] = noquote(as.character(NoOpinion_Group2))
      NoOpinions_Difference[1] = noquote(as.character(NoOpinion_Difference))

      # Adds percentage symbol "%" to no opinion entries.
      NoOpinions_Group1[1] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Group1[1], ""))
      NoOpinions_Group2[1] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Group2[1], ""))
      NoOpinions_Difference[1] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Difference[1], ""))

      # Adds the mean data to the dataframe.
      Means_Group1[1] = noquote(as.character(Mean_Group1))
      Means_Group2[1] = noquote(as.character(Mean_Group2))
      Means_Difference[1] = noquote(as.character(Mean_Difference))

      # Repeats creation of means and no opinion data for each demographic category.
      CategoryCounter_End = length(Crosstab_Group1)
      CategoryCounter = 1
      repeat{
        CategoryCounter = 1+CategoryCounter
        if(CategoryCounter>CategoryCounter_End){
          break
        }
        count1001 = CategoryCounter

        # Extracts demographic category.
        Category = colnames(Crosstab_Group1)[count1001]
        if(Category == "T"){
          Category = colnames(Crosstab_Group1)[count1001+1]
        }

        # Gets responses within each demographic category.
        Selections_Group1 = Dataset_Group1[grepl((paste("^", Category, "$", sep = "")), Demographics_Group1),]
        Selections_Group2 = Dataset_Group2[grepl((paste("^", Category, "$", sep = "")), Demographics_Group2),]
        Responses_Group1 = Selections_Group1[(which(names(Selections_Group1) == as.character(QuestionNumber_Group1)))]
        Responses_Group2 = Selections_Group2[(which(names(Selections_Group2) == as.character(QuestionNumber_Group2)))]

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
                NoOpinion_Group1 = "NaN"
              }
              # If not all entries are NAs, then the percentage of responses that are NA is calculated.
              if(isFALSE(AllNAs_1)){

                # If no entries are NAs, then the percentage of responses that are NA is 0.
                if(isFALSE(sum(is.na(Responses_Group_NA1))>0)){
                  NoOpinion_Group1 = 0
                }

                # If some entries are NAs, then the percentage of responses that are NA is 0.
                if((sum(is.na(Responses_Group_NA1))>0)){
                  Responses_Group_NA1[is.na(Responses_Group_NA1)] = -99

                  # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                  if(length(unlist(Responses_Group_NA1)) == 1){NoOpinion_Group1 = 100}

                  # Calculates the weighted percentage of responses that were NA
                  if(isFALSE(length(unlist(Responses_Group_NA1)) == 1)){
                    NoOpinion_Group1 = 100*prop.table(svytable(~unlist(Responses_Group_NA1), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA1), unlist(Weight_Group1))), weights = ~unlist(Weight_Group1))))[1]
                  }}
              }
            }

            # Gets the percentage of responses that were NA for group 2.
            {
              # If all entries are NAs, then the percentage of responses that are NA is marked as "NaN"
              if(AllNAs_2){
                NoOpinion_Group2 = "NaN"
              }
              # If not all entries are NAs, then the percentage of responses that are NA is calculated.
              if(isFALSE(AllNAs_2)){

                # If no entries are NAs, then the percentage of responses that are NA is 0.
                if(isFALSE(sum(is.na(Responses_Group_NA2))>0)){
                  NoOpinion_Group2 = 0
                }

                # If some entries are NAs, then the percentage of responses that are NA is 0.
                if((sum(is.na(Responses_Group_NA2))>0)){
                  Responses_Group_NA2[is.na(Responses_Group_NA2)] = -99

                  # If there is one entry that is NA, then the percentage of responses that are NA is 100.
                  if(length(unlist(Responses_Group_NA2)) == 1){NoOpinion_Group2 = 100}

                  # Calculates the weighted percentage of responses that were NA
                  if(isFALSE(length(unlist(Responses_Group_NA2)) == 1)){
                    NoOpinion_Group2 = 100*prop.table(svytable(~unlist(Responses_Group_NA2), svydesign(ids = ~1, data = as.data.frame(cbind(unlist(Responses_Group_NA2), unlist(Weight_Group2))), weights = ~unlist(Weight_Group2))))[1]
                  }}
              }
            }

            # Gets the percentage of responses that were NA for the difference between group 1 and 2.
            {
              # If group 1 or 2 entries are all NA, then the difference in the percentage of responses that are NA is NaN.
              if(AllNAs_1|AllNAs_2){
                NoOpinion_Difference = "NaN"
              }

              # If neither group 1 nor 2 entries are all NA, then the difference in the percentage of responses that are NA is calculated.
              if(isFALSE(AllNAs_1|AllNAs_2)){
                NoOpinion_Difference = NoOpinion_Group2-NoOpinion_Group1
                NoOpinion_Difference = (format(round(NoOpinion_Difference, digits = 1), nsmall = 1))
              }

              if(isFALSE(AllNAs_1)){
                # Formats the percentage.
                NoOpinion_Group1 = (format(round(NoOpinion_Group1, digits = 1), nsmall = 1))
              }

              if(isFALSE(AllNAs_2)){
                # Formats the percentage.
                NoOpinion_Group2 = (format(round(NoOpinion_Group2, digits = 1), nsmall = 1))
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

        # Adds the no opinion data to dataframe.
        NoOpinions_Group1[count1001] = noquote(as.character(NoOpinion_Group1))
        NoOpinions_Group2[count1001] = noquote(as.character(NoOpinion_Group2))
        NoOpinions_Difference[count1001] = noquote(as.character(NoOpinion_Difference))

        # Adds percentage symbol "%" to no opinion entries.
        NoOpinionPercentageAdder = data.frame(matrix(data = "", nrow = 1, ncol = 1))
        NoOpinions_Group1[count1001] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Group1[count1001], NoOpinionPercentageAdder))
        NoOpinions_Group2[count1001] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Group2[count1001], NoOpinionPercentageAdder))
        NoOpinions_Difference[count1001] = as.data.frame(mapply(function(x, y){paste(x, y, sep = "%")}, NoOpinions_Difference[count1001], NoOpinionPercentageAdder))

        # Adds the mean data to the dataframe.
        Means_Group1[count1001] = noquote(as.character(Mean_Group1))
        Means_Group2[count1001] = noquote(as.character(Mean_Group2))
        Means_Difference[count1001] = noquote(as.character(Mean_Difference))
      }

      # Adds header to p-values dataframe.
      colnames(Means_Group1) = colnames(Means_Group2) = colnames(Means_Difference) = colnames(Crosstab_Group1)

      # Combines means and no opinion data.
      Crosstab_Means = cbind(Means_Group1, Spacer, Means_Group2, Spacer, Means_Difference)
      Crosstab_NoOpinions = cbind(NoOpinions_Group1, Spacer, NoOpinions_Group2, Spacer, NoOpinions_Difference)

      # Combines crosstab with means and no opinion.
      colnames(Crosstab_Proportions) = colnames(Crosstab_NoOpinions) = colnames(Crosstab_Means)
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

        # Gets a demographic category.
        Category = colnames(Crosstab_Group1)[count1001]
        if(Category == "T"){
          Category = colnames(Crosstab_Group1)[count1001+1]
        }

        # Gets the responses corresponding with that demographic category.
        Selections_Group1 = Dataset_Group1[grepl((paste("^", Category, "$", sep = "")), Demographics_Group1),]
        Selections_Group2 = Dataset_Group2[grepl((paste("^", Category, "$", sep = "")), Demographics_Group2),]
        Responses_Group1 = Selections_Group1[(which(names(Selections_Group1) == as.character(QuestionNumber_Group1)))]
        Responses_Group2 = Selections_Group2[(which(names(Selections_Group2) == as.character(QuestionNumber_Group2)))]

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
      File_Name = paste("Responses", Name_Group, Weights, Demographic_Category, sep = " - ")
      Document_Title = gsub(" by Overall", "", paste(paste(Name_Group, ", Weighted by ", Weights, sep = ""), Demographic_Category, sep = " by "))
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
      Export_Word(Export, Outputs, Template, Crosstabs, Codebook, File_Name, Document_Title, Type = "Responses", Demographic_Category)
    }

    # Adds a legend.
    {
      # Creates a legend.
      Legend = Codebook[-1, match(Demographic_Category, names(Codebook))]
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

#' Creates a crosstab in Excel and Word comparing demographics
#' @export
Crosstab.Demographics = function(Dataset, Template, Group1, Group2, Outputs, Export, Alpha, Only_Significant){

  if(missing(Export)){Export = TRUE}
  if(missing(Outputs)){Outputs = ""}
  if(missing(Alpha)){Alpha = 0.05}
  if(missing(Only_Significant)){Only_Significant = FALSE}

  # Adds objects to global environment
  assign("Alpha", Alpha, envir = globalenv())
  assign("Only_Significant", Only_Significant, envir = globalenv())

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
  Overall[1,] = "Other"
  colnames(Overall) = "Overall"
  Codebook = cbind(Codebook, Overall)

  # Organizes the codebook by column value type (opinions, demographics, etc.)
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

  # Gets the column number of the first and last demographic and questions.
  Demographics = (which(Codebook[1,] == "Demographic"))
  Questions = (which(Codebook[1,] == "Opinion"))
  Demographic_Beg = Demographics[1]
  Demographic_End = Demographics[length(Demographics)]
  Question_Beg = Questions[1]
  Question_End = Questions[length(Questions)]

  # Values to stop crosstab creation once all questions and demographics have been analyzed
  DemographicsCounter = Demographic_Beg-1
  DemographicsCounter_End = Demographic_End

  # Creates a dataset to contain crosstabs
  Crosstabs = c()

  repeat{

    # Repeats crosstab creation if final demographic category has not be completed.
    if(DemographicsCounter == DemographicsCounter_End){
      break
    }

    # Values repetitions.
    DemographicsCounter = DemographicsCounter+1

    # Gets the demographic data from both groups
    Demographics_Group1 = as.matrix(Dataset_Group1[(which(names(Dataset_Group1) == colnames(Codebook[DemographicsCounter])))])
    Demographics_Group2 = as.matrix(Dataset_Group2[(which(names(Dataset_Group2) == colnames(Codebook[DemographicsCounter])))])

    # Gets the overall column
    Overall_Group1 = as.matrix(Dataset_Group1$Overall)
    Overall_Group2 = as.matrix(Dataset_Group2$Overall)

    # Gets the name of the demographic category
    Demographic_Category = colnames(Demographics_Group1)

    # Gets the levels for the demographic category.
    Levels_Demographics = unlist(Codebook[which(names(Codebook) == Demographic_Category)])[-1]
    Levels_Demographics = seq(from = 1, to = length(Levels_Demographics[complete.cases(Levels_Demographics)]))

    # Gets the demographic categories
    Demographic_Categories = unlist(Codebook[which(names(Codebook) == Demographic_Category)])
    Demographic_Categories = unlist(Demographic_Categories[complete.cases(Demographic_Categories)])[-1]

    # Gets numerical weights.
    if(Weight1 == "Unweighted"){
      Weight_Group1 = as.matrix(Dataset_Group1[(which(names(Dataset_Group1) == "Overall"))])
      Weight_Group1[] = 1
    }
    if(Weight2 == "Unweighted"){
      Weight_Group2 = as.matrix(Dataset_Group2[(which(names(Dataset_Group2) == "Overall"))])
      Weight_Group2[] = 1
    }
    if(Weight1!= "Unweighted"){

      # Check to see if the weights exist
      Input_Test(Dataset_Group1, Weight1)

      # Gets the weights
      Weight_Group1 = as.matrix(Dataset_Group1[(which(names(Dataset_Group1) == Weight1))])
    }
    if(Weight2!= "Unweighted"){

      # Check to see if the weights exist
      Input_Test(Dataset_Group2, Weight2)

      # Gets the weights
      Weight_Group2 = as.matrix(Dataset_Group2[(which(names(Dataset_Group2) == Weight2))])
    }

    # Assigns objects to global environment to accommodate svytable
    assign("Demographics_Group1", Demographics_Group1, envir = globalenv())
    assign("Demographics_Group2", Demographics_Group2, envir = globalenv())
    assign("Levels_Demographics", Levels_Demographics, envir = globalenv())

    # Gets the weighted frequencies of each demographic category
    Frequencies_Group1 = as.data.frame.matrix(as.matrix(svytable(~factor(Demographics_Group1, ordered = T, levels = Levels_Demographics), svydesign(ids = ~1, data = as.data.frame(cbind(Demographics_Group1, Weight_Group1)), weights = ~Weight_Group1))))
    Frequencies_Group2 = as.data.frame.matrix(as.matrix(svytable(~factor(Demographics_Group2, ordered = T, levels = Levels_Demographics), svydesign(ids = ~1, data = as.data.frame(cbind(Demographics_Group2, Weight_Group2)), weights = ~Weight_Group2))))

    # Gets the frequencies of each demographic category
    Crosstab_Frequencies_Group1 = round(Frequencies_Group1, digits = 0)
    Crosstab_Frequencies_Group2 = round(Frequencies_Group2, digits = 0)

    # Gets the relative frequencies of each demographic category
    Crosstab_RelativeFrequencies_Group1 = round(100*prop.table(Frequencies_Group1), digits = 1)
    Crosstab_RelativeFrequencies_Group2 = round(100*prop.table(Frequencies_Group2), digits = 1)

    # Combines frequencies and relative frequencies into one crosstab for each group
    Crosstab_Group1 = as.data.frame(mapply(function(a, b, c, d){paste(a, b, c, d, sep = "")}, "(", Crosstab_RelativeFrequencies_Group1, "%) ", Crosstab_Frequencies_Group1))
    Crosstab_Group2 = as.data.frame(mapply(function(a, b, c, d){paste(a, b, c, d, sep = "")}, "(", Crosstab_RelativeFrequencies_Group2, "%) ", Crosstab_Frequencies_Group2))

    # Inserts the name of the group
    colnames(Crosstab_Group1) = Name_Group1
    colnames(Crosstab_Group2) = Name_Group2

    # Inserts the row names
    rownames(Crosstab_Group1) = Demographic_Categories
    rownames(Crosstab_Group2) = Demographic_Categories

    # Combines the group crosstabs into one crosstab
    Crosstab = cbind(Crosstab_Group1, Crosstab_Group2)

    # Creates spacers
    VerticalSpacer1 = data.frame(matrix(data = 9999, nrow = nrow(Crosstab), ncol = 1))
    VerticalSpacer2 = data.frame(matrix(data = 9999, nrow = nrow(Crosstab), ncol = 1))

    # Performs a chi-squared test if possible and adds data to crosstabs
    if(sum(Frequencies_Group1[,1]+Frequencies_Group2[,1])>0){Demographic_Category = Test_Chi_Squared(Demographic_Category, Frequencies_Group1[,1], Frequencies_Group2[,1])}

    # Adds demographic number to crosstab
    VerticalSpacer1[1,] = Demographic_Category

    # Adds rownames to spacer
    VerticalSpacer2 = rownames(Crosstab)

    # Creates a vertical header and adds it to the crosstab
    VerticalHeader = cbind(VerticalSpacer1, VerticalSpacer2)
    colnames(VerticalHeader) = c("Demographic", "Category")
    Crosstab = cbind(VerticalHeader, Crosstab)

    # Creates a spacer to divide crosstabs
    HorizontalSpacer = data.frame(matrix(data = 9999, nrow = 1, ncol = ncol(Crosstab)))
    colnames(HorizontalSpacer) = colnames(Crosstab)
    Crosstab = rbind(HorizontalSpacer, Crosstab)

    # Adds crosstab to dataframe containing all crosstabs
    Crosstabs = rbind(Crosstabs, Crosstab)
  }

  # Adds the header of the crosstabs to the dataframe as entries
  Header = colnames(Crosstabs)
  Crosstabs = rbind(Header, Crosstabs)

  # Gets the sample sizes and puts them in a text format with the proper format
  SampleSizeNotation = noquote("n = ")
  SampleSize_Group1 = length(Overall_Group1)
  SampleSize_Group2 = length(Overall_Group2)
  SampleSize_Group1 = as.character(paste(SampleSizeNotation, SampleSize_Group1, sep = ""))
  SampleSize_Group2 = as.character(paste(SampleSizeNotation, SampleSize_Group2, sep = ""))

  # Creates a spacer
  Spacer = cbind(9999, 9999)

  # Creates a dataframe of sample size data
  SampleSizes = cbind(Spacer, SampleSize_Group1, SampleSize_Group2)

  # Adds sample size data to crosstabs with a spacer in-between
  colnames(SampleSizes) = colnames(Crosstabs)
  Spacer = data.frame(matrix(data = 9999, nrow = 1, ncol = length(Crosstabs)))
  colnames(Spacer) = colnames(Crosstabs)
  Crosstabs = rbind(Spacer, SampleSizes, Spacer, Crosstabs)

  # Replaces the column numbers of the demographic categories with the names of the categories
  Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("9999", "", x)}))
  Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("NaN%", "0%", x)}))

  # Creates the name of the file.
  {
    if(Weight1 == Weight2){Weights = Weight1}
    if(Weight1!= Weight2){Weights = paste(Weight1, Weight2, sep = " and ")}
    File_Name = paste("Demographics", Name_Group, Weights, sep = " - ")
    Document_Title = gsub(" by Overall", "", paste(paste(Name_Group, ", Weighted by ", Weights, sep = "")))
    Document_Title = gsub("Weighted by Unweighted", "Unweighted", Document_Title)
  }

  # Runs a pre-defined function that creates Excel spreadsheets from the crosstabs.
  {
    Export_Excel(Outputs, Crosstabs, File_Name, Name_Group)
  }

  # Runs a pre-defined function that creates Word documents from the crosstabs.
  {
    Export_Word(Export, Outputs, Template, Crosstabs, Codebook, File_Name, Document_Title, Type = "Demographics")
  }
}

#' Creates a crosstab in Excel and Word comparing opinions between two demographic categories.
#' @export
Crosstab.Opinions.Detailed = function(Dataset, Template, Group1, Group2, Demographics, Outputs, Export, Alpha, Only_Means, Only_Significant){

  Demographic_Category = Demographics[1]
  Category1_Name = Demographics[2]
  Category2_Name = Demographics[3]

  if(missing(Export)){Export = TRUE}
  if(missing(Outputs)){Outputs = ""}
  if(missing(Alpha)){Alpha = 0.05}
  if(missing(Only_Significant)){Only_Significant = FALSE}
  if(missing(Only_Means)){Only_Means = FALSE}

  # Adds objects to global environment
  assign("Alpha", Alpha, envir = globalenv())
  assign("Only_Significant", Only_Significant, envir = globalenv())
  assign("Only_Means", Only_Means, envir = globalenv())

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

  # Checks if the demographic categories selected exist
  if(isFALSE(any(na.exclude(Codebook[which(names(Codebook) == Demographic_Category)]) == Category1_Name))){
    stop(paste("There is no demographic category in the codebook named", Category1_Name, sep = " "))
  }
  if(isFALSE(any(na.exclude(Codebook[which(names(Codebook) == Demographic_Category)]) == Category2_Name))){
    stop(paste("There is no demographic category in the codebook named", Category2_Name, sep = " "))
  }

  # Adds overall column to codebook
  Overall = Codebook[1]
  Overall[] = NA
  Overall[1,] = "Demographic"
  colnames(Overall) = "Overall"
  Codebook = cbind(Codebook, Overall)

  # Organizes the codebook by column value type (opinions, demographics, etc.)
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

  # Gets the column number of the first and last questions.
  Questions = (which(Codebook[1,] == "Opinion"))
  Question_Beg = Questions[1]
  Question_End = Questions[length(Questions)]

  # Values to stop crosstab creation once all questions have been analyzed
  QuestionsCounter_End = Question_End+1

  # Gets the demographic categories
  Demographic_Categories = unlist(Codebook[which(names(Codebook) == Demographic_Category)])
  Demographic_Categories = unlist(Demographic_Categories[complete.cases(Demographic_Categories)])
  Demographic_Categories = Demographic_Categories[-1]

  # Gets the category numbers.
  Category1_Number = array(which(Demographic_Categories == Category1_Name))
  Category2_Number = array(which(Demographic_Categories == Category2_Name))

  # Creates datasets containing only the data for each demographic category
  {
    Dataset_Group1_Category1 = Dataset_Group1[grepl((paste("^", Category1_Number, "$", sep = "")), as.matrix(Dataset_Group1[(which(names(Dataset_Group1) == Demographic_Category))])),]
    Dataset_Group2_Category1 = Dataset_Group2[grepl((paste("^", Category1_Number, "$", sep = "")), as.matrix(Dataset_Group2[(which(names(Dataset_Group2) == Demographic_Category))])),]
    Dataset_Group1_Category2 = Dataset_Group1[grepl((paste("^", Category2_Number, "$", sep = "")), as.matrix(Dataset_Group1[(which(names(Dataset_Group1) == Demographic_Category))])),]
    Dataset_Group2_Category2 = Dataset_Group2[grepl((paste("^", Category2_Number, "$", sep = "")), as.matrix(Dataset_Group2[(which(names(Dataset_Group2) == Demographic_Category))])),]
  }

  # Capitalizes category names.
  Category1_Name = str_to_title(Category1_Name)
  Category2_Name = str_to_title(Category2_Name)

  # Repeats crosstab creation
  {
    # Creates counters to count repetitions
    QuestionsCounter = 0

    # Creates a dataset to contain crosstabs
    Crosstabs = c()

    # Resets to first question.
    QuestionsCounter = Question_Beg-1

    # Repeats crosstab creation for each question.
    repeat{

      # Counts repetitions
      QuestionsCounter = QuestionsCounter+1

      # Checks if there are responses in both Group1 and Group2 to compare
      repeat{
        ResponsesinBoth = isFALSE(identical((which(names(Dataset_Group2) == colnames(Dataset_Group1[QuestionsCounter]))), integer(0)))
        if(ResponsesinBoth){
          break
        }
        if(QuestionsCounter == QuestionsCounter_End){
          break
        }
        QuestionsCounter = QuestionsCounter+1
      }

      # Ends crosstab creation if all questions have been analyzed
      if(QuestionsCounter == QuestionsCounter_End){
        # Ends repetitions.
        break
      }

      # Gets the column number of Group 2
      ColumnNumber_Responses_Group1 = (which(names(Dataset_Group1) == colnames(Codebook[QuestionsCounter])))
      ColumnNumber_Responses_Group2 = (which(names(Dataset_Group2) == colnames(Codebook[QuestionsCounter])))

      # Gets the question number for the corresponding responses
      QuestionNumber_Group1 = noquote(colnames(Dataset_Group1)[ColumnNumber_Responses_Group1])
      QuestionNumber_Group2 = noquote(colnames(Dataset_Group2)[ColumnNumber_Responses_Group2])

      # Gets question number.
      QuestionNumberText = noquote(as.character(QuestionNumber_Group1))

      # Gets the text of the question for the corresponding reponses.
      QuestionText = as.character(Codebook[2, (which(names(Codebook) == as.character(QuestionNumber_Group1)))])

      # Creates a vertical header.
      Column1 = c(QuestionNumberText[1], 9999, 9999, 9999, 9999)
      Column2 = c(QuestionText, 9999, 9999, 9999, 9999)
      RowHeaders = data.frame(cbind(Column1, Column2))
      colnames(RowHeaders) = c("ID", "Prompt and Responses")

      # Creates spacers to fill in with means and no opinion data
      Spacer = data.frame(matrix(data = 9999, nrow = 1, ncol = 1))
      Spacer_Means = data.frame(matrix(data = 9999, nrow = 1, ncol = 3))

      # Creates a dataframe to contain mean data
      AllMeans = c(9999, 9999, 9999)
      Loop = 0
      repeat{

        # Breaks loop if all datasets have been analyzed
        Loop = Loop+1
        if(Loop == 4){
          break
        }

        # Selects datasets to compare and adds the name of the group to the header
        if(Loop == 1){
          Group1 = Dataset_Group1
          Group2 = Dataset_Group2
          RowHeaders[2, 2] = c("All Participants")
        }
        if(Loop == 2){
          Group1 = Dataset_Group1_Category1
          Group2 = Dataset_Group2_Category1
          RowHeaders[3, 2] = c(Category1_Name)
        }
        if(Loop == 3){
          Group1 = Dataset_Group1_Category2
          Group2 = Dataset_Group2_Category2
          RowHeaders[4, 2] = c(Category2_Name)
        }

        # Gets the numeric response data
        Responses_Group1 = Group1[(which(names(Group1) == as.character(QuestionNumber_Group1)))]
        Responses_Group2 = Group2[(which(names(Group2) == as.character(QuestionNumber_Group2)))]

        if(Weight1 == "Unweighted"){
          Weight_Group1 = as.matrix(Group1[(which(names(Group1) == "Overall"))])
          Weight_Group1[] = 1
        }
        if(Weight2 == "Unweighted"){
          Weight_Group2 = as.matrix(Group2[(which(names(Group2) == "Overall"))])
          Weight_Group2[] = 1
        }
        if(Weight1!= "Unweighted"){

          # Check to see if the weights exist
          Input_Test(Group1, Weight1)

          # Gets the weight
          Weight_Group1 = as.matrix(Group1[(which(names(Group1) == Weight1))])
        }
        if(Weight2!= "Unweighted"){

          # Check to see if the weights exist
          Input_Test(Group2, Weight2)

          # Gets the weight
          Weight_Group2 = as.matrix(Group2[(which(names(Group2) == Weight2))])
        }

        # TRUE if responses are all NAs (indicating respondents weren't surveyd)
        AllNAs_1 = (length(as.matrix(Responses_Group1)[(is.na(as.matrix(Responses_Group1)))]) == nrow(as.matrix(Responses_Group1)))
        AllNAs_2 = (length(as.matrix(Responses_Group2)[(is.na(as.matrix(Responses_Group2)))]) == nrow(as.matrix(Responses_Group2)))

        # If data is paired, removes NAs.
        if(isFALSE(AllNAs_1)&isFALSE(AllNAs_2)){
          # Removes entries for which there are missing responses in either groups.
          Responses_Group1_2 = cbind(Responses_Group1, Responses_Group2, Weight_Group1, Weight_Group2)
          Responses_Group1_2 = Responses_Group1_2[complete.cases(Responses_Group1_2),]

          # Inserts formatted data.
          Responses_Group1 = Responses_Group1_2[, 1]
          Responses_Group2 = Responses_Group1_2[, 2]
          Weight_Group1 = Responses_Group1_2[, 3]
          Weight_Group2 = Responses_Group1_2[, 4]
        }

        # Unlists data
        Responses_Group1 = unlist(Responses_Group1)
        Responses_Group2 = unlist(Responses_Group2)
        Weight_Group1 = unlist(Weight_Group1)
        Weight_Group2 = unlist(Weight_Group2)

        # Gets the means of the response data
        Mean_Group1 = wtd.mean(Responses_Group1, Weight_Group1, na.rm = TRUE)
        Mean_Group2 = wtd.mean(Responses_Group2, Weight_Group2, na.rm = TRUE)
        Mean_Difference = Mean_Group2-Mean_Group1

        # Formats the means
        Mean_Group1 = format(round(Mean_Group1, digits = 3), nsmall = 3)
        Mean_Group2 = format(round(Mean_Group2, digits = 3), nsmall = 3)
        Mean_Difference = format(round(Mean_Difference, digits = 3), nsmall = 3)

        # Performs a t-test if possible and adds data to crosstabs
        Mean_Difference = Test_T(Mean_Difference, Responses_Group1, Responses_Group2, Weight_Group1, Weight_Group2, Paired = TRUE)

        # Adds the means data to the dataframe
        Means = Spacer_Means
        Means[1] = noquote(as.character(Mean_Group1))
        Means[2] = noquote(as.character(Mean_Group2))
        Means[3] = noquote(as.character(Mean_Difference))
        AllMeans = rbind(AllMeans, Means)
      }

      # Adds description to header column
      RowHeaders[5, 2] = paste("Absolute Difference Between", Category1_Name, "and", Category2_Name, sep = " ")

      # Gets the numeric response data
      Responses_Group1_Category1 = Dataset_Group1_Category1[(which(names(Dataset_Group1_Category1) == as.character(QuestionNumber_Group1)))]
      Responses_Group1_Category2 = Dataset_Group1_Category2[(which(names(Dataset_Group1_Category2) == as.character(QuestionNumber_Group2)))]
      Responses_Group2_Category1 = Dataset_Group2_Category1[(which(names(Dataset_Group2_Category1) == as.character(QuestionNumber_Group1)))]
      Responses_Group2_Category2 = Dataset_Group2_Category2[(which(names(Dataset_Group2_Category2) == as.character(QuestionNumber_Group2)))]
      Responses_All_Category1 = as.numeric(as.matrix(Responses_Group2_Category1))-as.numeric(as.matrix(Responses_Group1_Category1))
      Responses_All_Category2 = as.numeric(as.matrix(Responses_Group2_Category2))-as.numeric(as.matrix(Responses_Group1_Category2))

      # Gets the numerical weights for Category 1
      if(Weight1 == "Unweighted"){
        Weight_Group1_Category1 = as.matrix(Dataset_Group1_Category1[(which(names(Dataset_Group1_Category1) == "Overall"))])
        Weight_Group1_Category1[] = 1
      }
      if(Weight2 == "Unweighted"){
        Weight_Group2_Category1 = as.matrix(Dataset_Group2_Category1[(which(names(Dataset_Group2_Category1) == "Overall"))])
        Weight_Group2_Category1[] = 1
      }
      if(Weight1!= "Unweighted"){
        Weight_Group1_Category1 = as.matrix(Dataset_Group1_Category1[(which(names(Dataset_Group1_Category1) == Weight1))])
      }
      if(Weight2!= "Unweighted"){
        Weight_Group2_Category1 = as.matrix(Dataset_Group2_Category1[(which(names(Dataset_Group2_Category1) == Weight2))])
      }

      # Gets the numerical weights for Category 2
      if(Weight1 == "Unweighted"){
        Weight_Group1_Category2 = as.matrix(Dataset_Group1_Category2[(which(names(Dataset_Group1_Category2) == "Overall"))])
        Weight_Group1_Category2[] = 1
      }
      if(Weight2 == "Unweighted"){
        Weight_Group2_Category2 = as.matrix(Dataset_Group2_Category2[(which(names(Dataset_Group2_Category2) == "Overall"))])
        Weight_Group2_Category2[] = 1
      }
      if(Weight1!= "Unweighted"){
        Weight_Group1_Category2 = as.matrix(Dataset_Group1_Category2[(which(names(Dataset_Group1_Category2) == Weight1))])
      }
      if(Weight2!= "Unweighted"){
        Weight_Group2_Category2 = as.matrix(Dataset_Group2_Category2[(which(names(Dataset_Group2_Category2) == Weight2))])
      }

      # Unlists response data
      Responses_Group1_Category1 = as.numeric(unlist(Responses_Group1_Category1))
      Responses_Group1_Category2 = as.numeric(unlist(Responses_Group1_Category2))
      Responses_Group2_Category1 = as.numeric(unlist(Responses_Group2_Category1))
      Responses_Group2_Category2 = as.numeric(unlist(Responses_Group2_Category2))
      Responses_All_Category1 = as.numeric(unlist(Responses_All_Category1))
      Responses_All_Category2 = as.numeric(unlist(Responses_All_Category2))

      # Gets the difference in the means of the response data
      Mean_Difference_Group1 = abs(as.numeric(AllMeans[3, 1])-as.numeric(AllMeans[4, 1]))
      Mean_Difference_Group2 = abs(as.numeric(AllMeans[3, 2])-as.numeric(AllMeans[4, 2]))
      Mean_Difference_All = Mean_Difference_Group2-Mean_Difference_Group1

      # Formats the means
      Mean_Difference_Group1 = format(round(Mean_Difference_Group1, digits = 3), nsmall = 3)
      Mean_Difference_Group2 = format(round(Mean_Difference_Group2, digits = 3), nsmall = 3)
      Mean_Difference_All = format(round(Mean_Difference_All, digits = 3), nsmall = 3)

      # Performs t-tests if possible and adds data to crosstabs
      Mean_Difference_Group1 = Test_T(Mean_Difference_Group1, Responses_Group1_Category1, Responses_Group1_Category2, Weight_Group1_Category1, Weight_Group1_Category2, Paired = FALSE)
      Mean_Difference_Group2 = Test_T(Mean_Difference_Group2, Responses_Group2_Category1, Responses_Group2_Category2, Weight_Group2_Category1, Weight_Group2_Category2, Paired = FALSE)
      Mean_Difference_All = Test_T(Mean_Difference_All, Responses_All_Category1, Responses_All_Category2, Weight_Group2_Category1, Weight_Group2_Category2, Paired = FALSE)

      # Adds the means data to the dataframe
      Means = Spacer_Means
      Means[1] = noquote(as.character(Mean_Difference_Group1))
      Means[2] = noquote(as.character(Mean_Difference_Group2))
      Means[3] = noquote(as.character(Mean_Difference_All))
      AllMeans = rbind(AllMeans, Means)

      # Adds header to crosstab
      Header = as.data.frame(t(c(Name_Group1, Name_Group2, "Difference")))
      colnames(AllMeans) = Header
      QuestionCrosstab = cbind(RowHeaders, AllMeans)

      # Adds spacer for between different crosstabs
      Spacer = data.frame(matrix(data = 9999, nrow = 1, ncol = length(QuestionCrosstab)))
      colnames(Spacer) = colnames(QuestionCrosstab)

      # Combines crosstabs to one dataframe
      Crosstabs = rbind(Crosstabs, Spacer, QuestionCrosstab)
    }

    # Adds crosstab headers to dataframe
    ColumnHeaders = c(colnames(Crosstabs))
    Crosstabs = as.matrix(Crosstabs)
    Crosstabs = rbind(ColumnHeaders, Crosstabs)
    Crosstabs = data.frame(Crosstabs)

    # Creates the name and title of the files.
    {
      if(Weight1 == Weight2){Weights = Weight1}
      if(Weight1!= Weight2){Weights = paste(Weight1, Weight2, sep = " and ")}
      File_Name = paste("Detailed Responses", Name_Group, Weights, paste(Category1_Name, Category2_Name, sep = " v. "), sep = " - ")
      Document_Title = paste(Name_Group, paste(Category1_Name, Category2_Name, sep = " v. "), sep = " by ")
      Document_Title = gsub("Weighted by Unweighted", "Unweighted", Document_Title)
    }

    # Removes markers.
    {
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("NaN%", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("NaN", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("9999", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("Spacer", "", x)}))
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("(NaN%)", "", x)}))
    }

    # Runs a pre-defined function that creates Excel spreadsheets from the crosstabs.
    {
      Export_Excel(Outputs, Crosstabs, File_Name, Name_Group)
    }

    # Runs a pre-defined function that creates Word documents from the crosstabs.
    {
      Export_Word(Export, Outputs, Template, Crosstabs, Codebook, File_Name, Document_Title, Type = "Detailed Responses")
    }
  }
}

#' Creates a crosstab in Excel and Word of correlations between opinions and changes in opinion.
#' @export
Crosstab.Correlations = function(Dataset, Template, Outputs, Question_Beg, Question_End, Demographic_Beg, Demographic_End){

  # Gets the codebook, datasets, and names in a list.
  List_Datasets_Names = Dataset_Import(Input, Group1, Time1, Group2, Time2)

  # Extracts the list elements.
  Codebook = (List_Datasets_Names[[1]])
  Dataset_Group1 = (List_Datasets_Names[[2]])
  Dataset_Group2 = (List_Datasets_Names[[3]])
  Name_Group1 = (List_Datasets_Names[[4]])
  Name_Group2 = (List_Datasets_Names[[5]])

  # Sets up a counter to record repetitions.
  ComparisonsCounter = 0

  Dataset_Group1[] = suppressWarnings(lapply(Dataset_Group1, as.numeric))
  Dataset_Group2[] = suppressWarnings(lapply(Dataset_Group2, as.numeric))

  repeat{

    # Counts repetitions.
    ComparisonsCounter = ComparisonsCounter+1

    # Ends crosstab creation if all questions have been analyzed.
    if(ComparisonsCounter == 7){
      # Ends repetitions.
      break
    }

    # Gets the dataframes and names.
    if(ComparisonsCounter == 1){
      Dataset1 = Dataset_Group1
      Dataset2 = Dataset_Group1
      Name_Group = paste(Group1, "at", Time1, "v.", Time1)
      Name1 = paste(Group1, "at", Time1)
      Name2 = paste(Group1, "at", Time1)
    }
    if(ComparisonsCounter == 2){
      Dataset1 = Dataset_Group2
      Dataset2 = Dataset_Group2
      Name_Group = paste(Group1, "at", Time2, "v.", Time2)
      Name1 = paste(Group1, "at", Time2)
      Name2 = paste(Group1, "at", Time2)
    }
    if(ComparisonsCounter == 3){
      Dataset1 = Dataset_Group1
      Dataset2 = Dataset_Group2
      Name_Group = paste(Group1, "at", Time1, "v.", Time2)
      Name1 = paste(Group1, "at", Time1)
      Name2 = paste(Group1, "at", Time2)
    }
    if(ComparisonsCounter == 4){
      Dataset1 = Dataset_Group1
      Dataset2 = Dataset_Group2-Dataset_Group1
      Name_Group = paste(Group1, "at", Time1, "v.", paste(Time2, Time1, sep = "-"))
      Name1 = paste(Group1, "at", Time1)
      Name2 = paste(Group1, "at", paste(Time2, Time1, sep = "-"))
    }
    if(ComparisonsCounter == 5){
      Dataset1 = Dataset_Group2
      Dataset2 = Dataset_Group2-Dataset_Group1
      Name_Group = paste(Group1, "at", Time2, "v.", paste(Time2, Time1, sep = "-"))
      Name1 = paste(Group1, "at", Time2)
      Name2 = paste(Group1, "at", paste(Time2, Time1, sep = "-"))
    }
    if(ComparisonsCounter == 6){
      Dataset1 = Dataset_Group2-Dataset_Group1
      Dataset2 = Dataset_Group2-Dataset_Group1
      Name_Group = paste(Group1, "at", paste(Time2, Time1, sep = "-"), "v.", paste(Time2, Time1, sep = "-"))
      Name1 = paste(Group1, "at", paste(Time2, Time1, sep = "-"))
      Name2 = paste(Group1, "at", paste(Time2, Time1, sep = "-"))
    }

    # Creates a dataframe of the selected questions.
    Dataset1_Questions = Dataset1[match(Question_Beg, names(Dataset1)):match(Question_End, names(Dataset1))]
    Dataset2_Questions = Dataset2[match(Question_Beg, names(Dataset2)):match(Question_End, names(Dataset2))]

    # Creates a dataframe of the selected demographics
    Dataset1_Demographics = Dataset1[match(Demographic_Beg, names(Dataset1)):match(Demographic_End, names(Dataset1))]
    Dataset2_Demographics = Dataset2[match(Demographic_Beg, names(Dataset2)):match(Demographic_End, names(Dataset2))]

    # Creates a dataframe of the selected questions and demographics.
    Dataset1 = cbind(Dataset1_Questions, Dataset1_Demographics)
    Dataset2 = cbind(Dataset2_Questions, Dataset2_Demographics)

    # Converts the dataframes to matrices.
    Dataset1 = as.matrix(Dataset1)
    Dataset2 = as.matrix(Dataset2)

    # Creates a matrix of correlations.
    Correlation = suppressWarnings(as.data.frame.matrix(format(round((rcorr(Dataset1, Dataset2)$r)^2, digits = 3), nsmall = 3)))
    Correlation = Correlation[(1+nrow(Correlation)/2):(nrow(Correlation)), 1:(ncol(Correlation)/2)]

    # Creates a matrix of p-values.
    Significance = suppressWarnings(as.data.frame.matrix(format(round(rcorr(Dataset1, Dataset2)$P, digits = 3), nsmall = 3)))
    Significance = Significance[(1+nrow(Significance)/2):(nrow(Significance)), 1:(ncol(Significance)/2)]

    # Combines the correlations and p-values into one matrix.
    Crosstabs = as.matrix(mapply(function(a, b, c, d){paste(a, b, c, d, sep = "")}, Correlation, " (", Significance, ")"))

    # Converts the matrix to a dataframe.
    Crosstabs = as.data.frame(Crosstabs)

    # Removes correlations and p-values when the p-value is greater than the alpha.
    if(Only_Significant){
      Crosstabs[Significance>Alpha] = NA
    }

    # Adds the column and row names to the dataframe.
    colnames(Crosstabs) = rownames(Crosstabs) = colnames(Dataset1)
    Crosstabs = cbind(rownames(Crosstabs), Crosstabs)
    Crosstabs = rbind(colnames(Crosstabs), Crosstabs)

    # Adds the names of the datasets being compared to the crosstabs.
    Crosstabs = cbind(rownames(Crosstabs), Crosstabs)
    Crosstabs = rbind(colnames(Crosstabs), Crosstabs)
    Crosstabs[1,] = NA
    Crosstabs[,1] = NA
    Crosstabs[1,2] = Name1
    Crosstabs[2,1] = Name2

    # Removes markers.
    {
      Crosstabs[Crosstabs == "rownames(Crosstabs)"] = Crosstabs[Crosstabs == " 1.000 (0.000)"] = Crosstabs[Crosstabs == " 1.000 (   NA)"] = Crosstabs[Crosstabs == "    NA (  NaN)"] = Crosstabs[Crosstabs == "   NaN (  NaN)"] = NA
      Crosstabs = data.frame(lapply(Crosstabs, function(x) {gsub("()", "", x)}))
    }

    # Removes rows that are all NA.
    Crosstabs = Crosstabs[!(apply(Crosstabs[, 2:nrow(Crosstabs)], 1, function(x) all(is.na(x)))), !(apply(Crosstabs[2:ncol(Crosstabs),], 2, function(x) all(is.na(x))))]

    # Creates the name of the file.
    {
      File_Name = paste("Regressions", Name_Group, sep = " - ")
      Document_Title = paste(Name_Group)
    }

    # Runs a pre-defined function that creates Excel spreadsheets from the crosstabs.
    {
      Export_Excel(Outputs, Crosstabs, File_Name, Name_Group)
    }

    # Runs a pre-defined function that creates Word documents from the crosstabs.
    {
      Export_Word(Export, Outputs, Template, Crosstabs, Codebook, File_Name, Document_Title, Type = "Regressions")
    }
  }

}

#' Creates a crosstab of regressions.
#' @export
Multiple.Regression=function(Dataset, Template, Model, Group1, Group2, Outputs, Export, Alpha, Only_Significant){

  if(missing(Export)){Export = TRUE}
  if(missing(Outputs)){Outputs = ""}
  if(missing(Alpha)){Alpha = 0.05}
  if(missing(Only_Significant)){Only_Significant = FALSE}

  # Adds objects to global environment
  assign("Alpha", Alpha, envir = globalenv())
  assign("Only_Significant", Only_Significant, envir = globalenv())

  # Gets the codebook, datasets, and names in a list.
  List_Datasets_Names = Dataset_Import(Dataset, Group1, Group2)

  # Extracts the list elements.
  Codebook = (List_Datasets_Names[[1]])
  Dataset_Group1 = (List_Datasets_Names[[2]])
  Dataset_Group2 = (List_Datasets_Names[[3]])
  Group1 = (List_Datasets_Names[[7]])
  Group2 = (List_Datasets_Names[[8]])
  Time1 = (List_Datasets_Names[[9]])
  Time2 = (List_Datasets_Names[[10]])

  # Gets the name of the difference in times.
  Time_Difference=paste(Time2, Time1, sep="-")

  # Gets the name of the group.
  Name_Group=paste(Group1, "at", paste(Time1, Time2, Time_Difference, sep=", "))

  # Converts the dataframes to all numeric values.
  Dataset_Group1[]=suppressWarnings(lapply(Dataset_Group1, as.numeric))
  Dataset_Group2[]=suppressWarnings(lapply(Dataset_Group2, as.numeric))

  # Gets the difference between the dataframes.
  Dataset_Difference=Dataset_Group2-Dataset_Group1

  # Gets the explanatory variable.
  Explanatory_Variable=as.character(Model[2])

  # Creates a dataframe.
  Crosstabs=c()

  # Sets up a counter to record repetitions.
  ComparisonsCounter=0

  repeat{

    # Counts repetitions.
    ComparisonsCounter=ComparisonsCounter+1

    # Ends crosstab creation if all questions have been analyzed.
    if(ComparisonsCounter==7){
      # Ends repetitions.
      break
    }

    # Iterates comparisons between the two datasets and their differences.
    if(ComparisonsCounter==1){
      # T1 v. T1
      Dataset=Dataset_Group1
      Time_Dependent_Variables=Time_Explanatory_Variable=Time1
    }
    if(ComparisonsCounter==2){
      # T2 v. T2
      Dataset=Dataset_Group2
      Time_Dependent_Variables=Time_Explanatory_Variable=Time2
    }
    if(ComparisonsCounter==3){
      # T2 v. T1
      Dataset=Dataset_Group1
      Time_Dependent_Variables=Time1

      Dataset[match(Explanatory_Variable, names(Dataset))]=Dataset_Group2[match(Explanatory_Variable, names(Dataset_Group2))]
      Time_Explanatory_Variable=Time2
    }
    if(ComparisonsCounter==4){
      # T2-1 v. T1
      Dataset=Dataset_Group1
      Time_Dependent_Variables=Time1

      Dataset[match(Explanatory_Variable, names(Dataset))]=Dataset_Difference[match(Explanatory_Variable, names(Dataset_Difference))]
      Time_Explanatory_Variable=Time_Difference
    }
    if(ComparisonsCounter==5){
      # T2-1 v. T2
      Dataset=Dataset_Difference
      Time_Dependent_Variables=Time_Difference

      Dataset[match(Explanatory_Variable, names(Dataset))]=Dataset_Group2[match(Explanatory_Variable, names(Dataset_Group2))]
      Time_Explanatory_Variable=Time2
    }
    if(ComparisonsCounter==6){
      # T2-1 v. T2-1
      Dataset=Dataset_Difference
      Time_Dependent_Variables=Time_Explanatory_Variable=Time_Difference
    }

    # Checks if a regression is possible.
    if(isFALSE(class(try((lm(Model, data=Dataset)),silent = TRUE))=="try-error")){

      # Gets the regression of the model.
      Regression=lm(Model, data=Dataset)

      # Gets the summary of the regression.
      Results=summary(Regression)

      # Gets the coefficients from the regression.
      Coefficients=format(round(as.data.frame(Results$coefficients)[-3], digits=3), nsmall=3)

      # Checks for multicollinearity in the regression.
      if(nrow(Coefficients)>2){
        Multicollinearity=vif(Regression)
        if(any(Multicollinearity>5)){
          print(Multicollinearity)
          print("Error: Too high multicollinearity.")
        }}

      # Gets p-values from the left and right tailed tests for the betas.
      Left_Tailed=format(round(as.data.frame(pt(coef(Results)[, 3], Regression$df, lower=TRUE)), digits=3), nsmall=3)
      Right_Tailed=format(round(as.data.frame(pt(coef(Results)[, 3], Regression$df, lower=FALSE)), digits=3), nsmall=3)
      Coefficients=cbind(Coefficients, Left_Tailed, Right_Tailed)

      # Removes non-significant p-values.
      if(Only_Significant){
        Significance=Coefficients[3:5]
        Significance[Significance>Alpha]=NA
        Coefficients[3:5]=Significance
      }

      # Gets the r-squared for the regression.
      RSquared=format(round(Results$adj.r.squared, digits=3), nsmall=3)

      if(isFALSE(class(try((format(round(pf(Results$fstatistic[1], Results$fstatistic[2], Results$fstatistic[3], lower.tail=F), digits=3), nsmall=3)),silent = TRUE))=="try-error")){

        # Gets the significance of the regression model.
        Significance=format(round(pf(Results$fstatistic[1], Results$fstatistic[2], Results$fstatistic[3], lower.tail=F), digits=3), nsmall=3)

        # Gets the sample size.
        Observations=Results$df[1]+Results$df[2]

        # Combines the regression summary statistics into one dataframes and adds column headers.
        Summary=rbind(RSquared, Significance, paste("n =", Observations))
        rownames(Summary)=c("R-Squared Adjusted", "Significance", "Observations")
        colnames(Summary)=NA

        # Adds column headers to the regression statistics for each variable.
        Coefficients=rbind(c("Beta", "Standard Error", "Significance (H0: B=0)", "Significance (H0: B<0)", "Significance (H0: B>0)"), Coefficients)

        # Adds row names to dataframe.
        Summary=cbind(rownames(Summary), Summary)

        # Creates a spacer.
        Spacer=matrix(data="", nrow=(nrow(Coefficients)-nrow(Summary)), ncol=2)

        # Adds spacer to summary statistics.
        Summary=rbind(Summary, Spacer)

        # Creates a spacer.
        Spacer=matrix(data="", nrow=nrow(Coefficients), ncol=1)

        # Combines the all the statistics together into one dataframe.
        Crosstab=cbind(Coefficients, Spacer, Summary)

        # Sets row names.
        rownames(Crosstab)[3:nrow(Coefficients)]=paste("(", Time_Dependent_Variables, ") ", rownames(Crosstab)[3:nrow(Coefficients)], sep="")

        # Adds row names to dataframe.
        Crosstab=cbind(rownames(Crosstab), Crosstab)

        # Adds the explanatory variable and the term "constant" to the regression.
        Crosstab[1,1]=paste("(", Time_Explanatory_Variable, ") ", as.character(Model[2]), sep="")
        Crosstab[2,1]="Constant"

        # Clears the column names.
        colnames(Crosstab)=NA

        # Adds regression to the dataframe.
        Crosstab=rbind(Crosstab, Crosstab[100,])
        Crosstabs=rbind(Crosstabs, Crosstab)
      }}
  }

  # Creates the name of the file.
  {
    Name=paste(as.character(Model[2]), "by", gsub("   ", ", ", str_replace_all(as.character(Model[3]), "[^[:alnum:]]", " ")), sep=" ")
    File_Name=paste("Multiple Regressions", Name_Group, Name, sep=" - ")
    Document_Title=paste(Name_Group, "at", Name)
  }

  # Runs a pre-defined function that creates Excel spreadsheets from the crosstabs.
  {
    Export_Excel(Outputs, Crosstabs, File_Name, Name_Group)
  }

  # Runs a pre-defined function that creates Word documents from the crosstabs.
  {
    Export_Word(Export, Outputs, Template, Crosstabs, Codebook, File_Name, Document_Title, Type="Multiple Regressions")
  }
}

#' Strips the dataset of all columns not present in the codebook.
#' @export
Dataset.Organize = function(Codebook, Datasets){
  
  library(stringr)
  library(dplyr)
  
  for(Dataset_Number in 1:length(Datasets)){
    Dataset = Datasets[[Dataset_Number]]
    OG_Names = names(Dataset)
    names(Dataset) = paste(seq(1,length(OG_Names)), OG_Names)
    Dataset = Dataset %>% mutate_all(funs(str_replace_all(., paste0("[", paste(c("&", "$", "£", "+"), collapse = ""), "]"), "_")))
    names(Dataset) = OG_Names
    Datasets[[Dataset_Number]] = Dataset
  }
  
  Codebook_Demographics = Codebook[which(Codebook[1,] == "Demographic")]
  
  for(Name in names(Codebook_Demographics)){
    
    Unique_Entries = c()
    
    for(Dataset_Number in 1:length(Datasets)){
      
      Dataset = Datasets[[Dataset_Number]]
      
      Dataset_Column = Dataset[which(names(Dataset) ==  Name)]
      
      if(length(Dataset_Column)>0){
        Dataset_Column_Unique_Entries = unique(Dataset_Column)
      } else {
        Dataset_Column_Unique_Entries = c()
      }
      
      Unique_Entries = append(Unique_Entries, Dataset_Column_Unique_Entries)
    }
    
    Unique_Entries = unique(unlist(Unique_Entries))
    if(length(Unique_Entries)>0){Unique_Entries = Unique_Entries[complete.cases(Unique_Entries)]}
    
    if(!is.null(Unique_Entries[1])){
      if(!is.na(Unique_Entries[1])){
        if(length(Unique_Entries)>0){
          
          Unique_Entries = Unique_Entries[complete.cases(Unique_Entries)]
          
          Unique_Entries = as.character(Unique_Entries[order(Unique_Entries)])
          
          Codebook[2:(length(Unique_Entries)+1), which(names(Codebook) == Name)] = Unique_Entries
          
        }}}
  }
  
  for(Dataset_Number in 1:length(Datasets)){
    
    Dataset = Datasets[[Dataset_Number]]
    
    Dataset_New = as.data.frame(matrix(nrow = nrow(Dataset), ncol = 0))
    
    for(Name in names(Codebook)){
      
      Codebook_Column = Codebook[which(names(Codebook) == Name)]
      
      Dataset_Column_Number = which(names(Dataset) == Name)[1] # Gets the column number in dataset corresponding to the column name from the codebook
      
      if((length(Dataset_Column_Number)>0)&(!is.na(Dataset_Column_Number))){
        
        Dataset_Column = Dataset[Dataset_Column_Number] # Gets the location of the matching column in the dataset if it exists
        
        Type = Codebook_Column[1,]
        
        if(Type == "Demographic"){
          
          Demographic_Categories = Codebook_Column[-1,]
          
          Demographic_Categories = Demographic_Categories[complete.cases(Demographic_Categories),]
          
          for(Category_Number in 1:nrow(Demographic_Categories)){
            
            Category_Text = Demographic_Categories[Category_Number,][[1]]
            
            Category_Number = as.character(Category_Number)
            
            Dataset_Column = stack(as_tibble(lapply(unlist(Dataset_Column), function(x)
              if(isTRUE(nchar(x[1])==nchar(Category_Text))){
                gsub(Category_Text, Category_Number, x)
              }
              else {x})))[1]
            
            colnames(Dataset_Column) = Name
          }
        }
        if(Type == "Opinion"){
          
          Scale = Codebook_Column[3,]
          Scale = unlist(strsplit(unlist(Scale), " to ")[])
          
          Lower_Bound = as.numeric(Scale[1])
          Upper_Bound = as.numeric(Scale[2])
          
          Dataset_Column = suppressWarnings(as.numeric(as.matrix(Dataset_Column)))
          
          if(length(Dataset_Column[Dataset_Column < Lower_Bound])>0){Dataset_Column[Dataset_Column < Lower_Bound] = NA}
          if(length(Dataset_Column[Dataset_Column > Upper_Bound])>0){Dataset_Column[Dataset_Column > Upper_Bound] = NA}
        }
      } else {
        
        Dataset_Column = Dataset[1]
        Dataset_Column[] = NA
        colnames(Dataset_Column) = Name # Creates a blank column if there is no matching column in the dataset
        
      }
      
      Dataset_Column = suppressWarnings(as.numeric(as.matrix(Dataset_Column))) # Forces numeric
      Dataset_Column = as_tibble(Dataset_Column)
      colnames(Dataset_Column) = Name
      Dataset_New = cbind(Dataset_New, Dataset_Column)
    }
    
    Datasets[[Dataset_Number]] = Dataset_New[order(Dataset_New$`Identification Number`),]
    
  }
  
  return(list(Datasets, Codebook))
}

#' Opens datasets.
Dataset_Import = function(File, GroupDetails1, GroupDetails2){

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
  Dataset_Group1 = suppressMessages(read_excel(File, (which(readxl::excel_sheets(File) == Name_Group1)), col_names = TRUE, col_types = "numeric"))
  Dataset_Group2 = suppressMessages(read_excel(File, (which(readxl::excel_sheets(File) == Name_Group2)), col_names = TRUE, col_types = "numeric"))

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
    Name_Group = paste(Group1, "at", Time1, "v.", Group2, "at", Time2)
  }

  # Returns a list of the dataframes and names.
  return(list(Codebook, Dataset_Group1, Dataset_Group2, Name_Group1, Name_Group2, Name_Group, Group1, Group2, Time1, Time2, Weight1, Weight2))
}

#' Tests column presence in dataset.
Input_Test = function(Dataset, Name){

  # Check to see if the named column exists in the dataset
  if(isFALSE(any(names(Dataset) == Name))){

    # If the column does not exist in the dataset, returns an error
    stop(paste("There is no data named ", Name, ".", sep = ""))
  }
}

#' Converts numeric responses to text responses.
Responses_to_Text = function(Responses, ColumnName, Codebook){

  # Converts column name to character.
  ColumnName = as.character(ColumnName)

  # Gets the text value
  Response_Negative = as.character(Codebook[4, match(ColumnName, names(Codebook))])
  Response_Neutral = as.character(Codebook[5, match(ColumnName, names(Codebook))])
  Response_Positive = as.character(Codebook[6, match(ColumnName, names(Codebook))])

  Scale = as.character(Codebook[3, match(ColumnName, names(Codebook))])

  # Condenses the 0-100 scale.
  if(Scale == "0 to 100"){

    # Ensures there are no responses outside of the scale.
    if(suppressWarnings((max(na.omit(Responses)))>100)|suppressWarnings(min(na.omit(Responses)))<0){stop(paste("There is a response outside of the", Scale, "scale specified in the codebook for", ColumnName, sep = " "))}

    # Replaces the numeric responses with text responses.
    Responses[Responses>66] = Response_Positive
    Responses[Responses < 67 & Responses > 33] = Response_Neutral
    Responses[Responses<33] = Response_Negative
  }

  # Condenses the 0-10 scale.
  if(Scale == "0 to 10"){

    # Ensures there are no responses outside of the scale.
    if(suppressWarnings((max(na.omit(Responses)))>10)|suppressWarnings(min(na.omit(Responses)))<0){stop(paste("There is a response outside of the", Scale, "scale specified in the codebook for", ColumnName, sep = " "))}

    # Replaces the numeric responses with text responses.
    Responses[Responses>5] = Response_Positive
    Responses[Responses == 5] = Response_Neutral
    Responses[Responses<5] = Response_Negative
  }

  # Condenses the 1-5 scale.
  if(Scale == "1 to 5"){

    # Ensures there are no responses outside of the scale.
    if(suppressWarnings((max(na.omit(Responses)))>5)|suppressWarnings(min(na.omit(Responses)))<1){stop(paste("There is a response outside of the", Scale, "scale specified in the codebook for", ColumnName, sep = " "))}

    # Replaces the numeric responses with text responses.
    Responses[Responses>3] = Response_Positive
    Responses[Responses == 3] = Response_Neutral
    Responses[Responses<3] = Response_Negative
  }

  # Condenses the 1-3 scale.
  if(Scale == "1 to 3"){

    # Ensures there are no responses outside of the scale.
    if(suppressWarnings((max(na.omit(Responses)))>3)|suppressWarnings(min(na.omit(Responses)))<1){stop(paste("There is a response outside of the", Scale, "scale specified in the codebook for", ColumnName, sep = " "))}

    # Replaces the numeric responses with text responses.
    Responses[Responses == 1] = Response_Negative
    Responses[Responses == 2] = Response_Neutral
    Responses[Responses == 3] = Response_Positive
  }

  # Condenses the 0-1 scale.
  if(Scale == "0 to 1"){

    # Ensures there are no responses outside of the scale.
    if(suppressWarnings((max(na.omit(Responses)))>1)|suppressWarnings(min(na.omit(Responses)))<0){stop(paste("There is a response outside of the", Scale, "scale specified in the codebook for", ColumnName, sep = " "))}

    # Replaces the numeric responses with text responses.
    Responses[Responses == 1] = Response_Positive
    Responses[Responses == 0] = Response_Negative
  }

  # Returns the formatted responses.
  return(Responses)
}

#' Defines a function that runs a two sample t-test and adds significance markers.
Test_T = function(ValuetoMark, Group1, Group2, Weight1, Weight2, Paired){

  # Unlists data
  Group1 = unlist(Group1)
  Group2 = unlist(Group2)
  Weight1 = unlist(Weight1)
  Weight2 = unlist(Weight2)

  # Performs a t-test if possible
  Test_Validator = (((length(na.exclude(Group1))>4)&(length(na.exclude(Group2))>4)))

  if(Test_Validator){

    if(Paired){
      # Gets the p-value from the t-test.
      PValue = suppressWarnings(as.vector(wtd.t.test((Group2-Group1), weight = Weight1)$coefficients[3]))
    }

    if(isFALSE(Paired)){
      # Gets the p-value from the t-test.
      PValue = suppressWarnings(as.vector(wtd.t.test(Group1, Group2, weight = Weight1, weighty = Weight2, alternative = "two.tailed")$coefficients[3]))
    }

    if(is.na(PValue) == FALSE){

      # If specified, always adds p-values.
      if(isFALSE(Only_Significant)){
        # Adds p-values
        ValuetoMark = noquote(paste(ValuetoMark, " ", "(", format(round((PValue), digits = 3), nsmall = 3), ")", sep = ""))
      }

      # If specified, only adds p-values less than or equal to alpha specified.
      if(Only_Significant&(PValue<= Alpha)){
        # Adds p-values
        ValuetoMark = noquote(paste(ValuetoMark, " ", "(", format(round((PValue), digits = 3), nsmall = 3), ")", sep = ""))
      }
    }
  }
  return(ValuetoMark)
}

#' Defines a function that runs a chi-squared test.
Test_Chi_Squared = function(ValuetoMark, Group1, Group2){

  # Organizes the inputs.
  Observed_Expected = cbind(Group1, Group2)

  # Removes all zero categories.
  Observed_Expected = Observed_Expected[!(apply(Observed_Expected, 1, function(y) all(y ==  0))),]

  # Performs a chi-squared test and gets the p-value.
  PValue = suppressWarnings(chisq.test(Observed_Expected, simulate.p.value = FALSE)$p.value)

  if(is.na(PValue) == FALSE){

    # If specified, always adds p-values.
    if(isFALSE(Only_Significant)){
      # Adds p-values
      ValuetoMark = noquote(paste(ValuetoMark, " ", "(", format(round((PValue), digits = 3), nsmall = 3), ")", sep = ""))
    }

    # If specified, only adds p-values less than or equal to alpha specified.
    if(Only_Significant&(PValue<= Alpha)){
      # Adds p-values
      ValuetoMark = noquote(paste(ValuetoMark, " ", "(", format(round((PValue), digits = 3), nsmall = 3), ")", sep = ""))
    }

    # If a warning is generated due to an expect number less than 5, a note is added to the p-value.
    if(any(Group2<5)){
      ValuetoMark = noquote(paste(ValuetoMark, "[Warning: P-value may be incorrect, as at least one expected value is less than 5.]", sep = " "))
    }
  }

  return(ValuetoMark)
}

#' Exports crosstab in Excel.
Export_Excel = function(Outputs, Crosstabs, File_Name, Sheet_Name){

  # Gets the file name.
  File_Name = paste(File_Name, ".xlsx", sep = "")

  # Ensures the sheet name is the proper length.
  if(nchar(Sheet_Name)>31){
    Sheet_Name = paste(strtrim(Sheet_Name,28), "...", sep = "")
  }

  # Exports crosstab to Excel file
  write.xlsx(Crosstabs, paste0(getwd(), Outputs, "/", File_Name), sheetName = Sheet_Name, showNA = F, colNames = F, rownames = F, overwrite = T)

  # Notifies of document export
  print(noquote(paste("Exported:", File_Name)))
}

#' Exports crosstab in Word.
Export_Word = function(Export, Outputs, Template, Crosstabs, Codebook, File_Name, Document_Title, Type, Demographic_Category){

  if(Export){
    if(ncol(Crosstabs)<14){

  File_Name = paste(File_Name, ".docx", sep = "")

  # Gets the number of columns
  ColumnNumbers = ncol(Crosstabs)
  RowNumbers = nrow(Crosstabs)

  if(Type == "Responses"){

    # Creates a legend.
    Legend = Codebook[-1, match(Demographic_Category, names(Codebook))]
    Legend = as.data.frame(Legend[complete.cases(Legend)])
    Legend = rbind(c("T", "Total"), (cbind(rownames(Legend), Legend)))

    # Gets the column names of the crosstabs
    ColumnNames = Crosstabs[5,]

    # Gets parts of the crosstab to print
    Titles = Crosstabs[1,]
    SampleSizes = Crosstabs[3,]
    Crosstabs = Crosstabs[5:nrow(Crosstabs),]

    # Gets the headers and sets them as the column names
    colnames(Titles) = Titles[1,]
    colnames(Crosstabs) = Crosstabs[1,]
    Crosstabs = Crosstabs[-(1:2),]

    # Gets sequential numbers for the column names
    colnames(Titles) = seq(1, ncol(Crosstabs), 1)
    colnames(SampleSizes) = seq(1, ncol(Crosstabs), 1)
    colnames(Crosstabs) = seq(1, ncol(Crosstabs), 1)

    # Gets the number of columns
    ColumnNumbers = ncol(Crosstabs)
    Increment = ((ColumnNumbers-4)/3)-1

    # Converts the dataframes to flextables
    Legend = flextable(Legend)
    Titles = flextable(Titles)
    SampleSizes = flextable(SampleSizes)
    Crosstabs = flextable(Crosstabs)

    # Deletes excess header.
    Crosstabs = delete_part(Crosstabs, part = "header")

    # Removes the excess headers
    Legend = suppressWarnings(delete_part(Legend, part = "header"))
    Titles = delete_part(Titles, part = "header")
    SampleSizes = delete_part(SampleSizes, part = "header")

    # Adds correct header row
    Legend = add_header_row(Legend, top = TRUE, values = c("Legend", ""))
    Crosstabs = add_header_row(Crosstabs, top = TRUE, values = ColumnNames)

    # Sets the padding for the tables
    Legend = padding(Legend, padding.bottom = 2, part = "header")
    Legend = padding(Legend, padding = 0, part = "all")

    # Sets the widths of the columns and heights of the rows
    Legend = width(Legend, j = 1, width = 0.2)
    Legend = width(Legend, j = 2, width = 2.5)
    Titles = width(Titles, j = 1, width = 1.0)
    SampleSizes = width(SampleSizes, j = 1, width = 1.0)
    Crosstabs = width(Crosstabs, j = 1, width = 1.0)
    Titles = width(Titles, j = 2, width = 2.5)
    SampleSizes = width(SampleSizes, j = 2, width = 2.5)
    Crosstabs = width(Crosstabs, j = 2, width = 2.5)

    # Sets the widths of the columns
    Titles = width(Titles, j = 3:(ColumnNumbers), width = 0.6)
    SampleSizes = width(SampleSizes, j = 3:(ColumnNumbers), width = 0.6)
    Crosstabs = width(Crosstabs, j = 3:(ColumnNumbers), width = 0.6)
    if(ColumnNumbers<10){
      Titles = width(Titles, j = 3:(ColumnNumbers), width = 1.2)
      SampleSizes = width(SampleSizes, j = 3:(ColumnNumbers), width = 1.2)
      Crosstabs = width(Crosstabs, j = 3:(ColumnNumbers), width = 1.2)
    }

    # Sets font size
    Legend = fontsize(Legend, size = 8, part = "all")
    Titles = fontsize(Titles, size = 8, part = "all")
    SampleSizes = fontsize(SampleSizes, size = 8, part = "all")
    Crosstabs = fontsize(Crosstabs, size = 8, part = "all")

    # Sets font type
    Legend = font(Legend, fontname = "Arial", part = "all")
    Titles = font(Titles, fontname = "Arial", part = "all")
    SampleSizes = font(SampleSizes, fontname = "Arial", part = "all")
    Crosstabs = font(Crosstabs, fontname = "Arial", part = "all")

    # Sets the borders on the tables
    Legend = border_remove(Legend)
    Titles = border_remove(Titles)
    SampleSizes = border_remove(SampleSizes)
    Crosstabs = border_remove(Crosstabs)
    Crosstabs = border(Crosstabs, border.bottom = fp_border(color = "black", style = "solid", width = 1), part = "header")
    Crosstabs = border(Crosstabs, border.right = fp_border(color = "black", style = "solid", width = 1), part = "body")

    # Bolds the text
    Legend = bold(Legend, bold = TRUE, part = "header")
    Titles = bold(Titles, bold = TRUE, part = "all")
    SampleSizes = bold(SampleSizes, bold = TRUE, part = "all")
    Crosstabs = bold(Crosstabs, j = 1, bold = TRUE, part = "body")
    Crosstabs = bold(Crosstabs, bold = TRUE, part = "header")

    # Merges cells
    Legend = merge_at(Legend, part = "header")
    Titles = merge_at(Titles, j = (3+2*2+2*Increment):(3+2*2+3*Increment), part = "body")
    Titles = merge_at(Titles, j = (3+1*2+1*Increment):(3+1*2+2*Increment), part = "body")
    Titles = merge_at(Titles, j = (3+0*2+0*Increment):(3+0*2+1*Increment), part = "body")

    # Aligns the text
    Legend = align(Legend, align = "left", part = "all")
    Titles = align(Titles, align = "center", part = "all")
    SampleSizes = align(SampleSizes, align = "center", part = "all")
    Crosstabs = align(Crosstabs, j = 1:2, align = "left", part = "all")
    Crosstabs = align(Crosstabs, j = 3:ColumnNumbers, align = "center", part = "all")
    Crosstabs = align(Crosstabs, align = "center", part = "header")

    # Sets the description on the document footer.
    Description = "T1 is before deliberation. T2 is after deliberation. Numbers not in parentheses are mean responses or percentages. Numbers in parentheses are p-values resulting from a t-test."

    # Creates the Word document for export
    if(!class(try((suppressWarnings(read_docx(paste0(getwd(), Template)) %>%
                                    body_add_par("", pos = "after") %>%
                                    body_add_flextable(Legend, align = "left") %>%
                                    body_add_par("", pos = "after") %>%
                                    body_add_flextable(Titles, align = "left") %>%
                                    body_add_par("", pos = "after") %>%
                                    body_add_flextable(SampleSizes, align = "left") %>%
                                    body_add_par("", pos = "after") %>%
                                    body_add_flextable(Crosstabs, align = "left") %>%
                                    headers_replace_all_text("Type", Type) %>%
                                    headers_replace_all_text("Title", Document_Title) %>%
                                    footers_replace_all_text("Description", Description))),silent = TRUE)) == "try-error"){
      WordDocument = suppressWarnings(read_docx(paste0(getwd(), Template)) %>%
                                      body_add_par("", pos = "after") %>%
                                      body_add_flextable(Legend, align = "left") %>%
                                      body_add_par("", pos = "after") %>%
                                      body_add_flextable(Titles, align = "left") %>%
                                      body_add_par("", pos = "after") %>%
                                      body_add_flextable(SampleSizes, align = "left") %>%
                                      body_add_par("", pos = "after") %>%
                                      body_add_flextable(Crosstabs, align = "left") %>%
                                      headers_replace_all_text("Type", Type) %>%
                                      headers_replace_all_text("Title", Document_Title) %>%
                                      footers_replace_all_text("Description", Description))
    }
    else {print("Cannot export Word version due to large file size or template error.")}
  }

  if(Type == "Detailed Responses"){

    # Gets the headers and sets them as the column names
    ColumnNames = Crosstabs[1,]
    Crosstabs = Crosstabs[-(1),]

    # Converts the dataframes to flextables
    Crosstabs = flextable(Crosstabs)

    # Deletes excess header.
    Crosstabs = delete_part(Crosstabs, part = "header")

    # Adds correct header row
    Crosstabs = add_header_row(Crosstabs, top = TRUE, values = ColumnNames)

    # Sets the widths of the columns and heights of the rows
    Crosstabs = width(Crosstabs, j = 1, width = 1.0)
    Crosstabs = width(Crosstabs, j = 2, width = 2.5)

    # Sets the widths of the columns depending on the number of columns
    Crosstabs = width(Crosstabs, j = 3:(ColumnNumbers), width = 1.2)

    # Sets font size
    Crosstabs = fontsize(Crosstabs, size = 8, part = "all")

    # Sets font type
    Crosstabs = font(Crosstabs, fontname = "Arial", part = "all")

    # Sets the borders on the tables
    Crosstabs = border_remove(Crosstabs)
    Crosstabs = border(Crosstabs, border.bottom = fp_border(color = "black", style = "solid", width = 1), part = "header")
    Crosstabs = border(Crosstabs, border.right = fp_border(color = "black", style = "solid", width = 1), part = "body")

    # Bolds the text
    Crosstabs = bold(Crosstabs, j = 1, bold = TRUE, part = "body")
    Crosstabs = bold(Crosstabs, bold = TRUE, part = "header")

    # Aligns the text
    Crosstabs = align(Crosstabs, j = 1:2, align = "left", part = "all")
    Crosstabs = align(Crosstabs, j = 3:5, align = "center", part = "all")
    Crosstabs = align(Crosstabs, align = "center", part = "header")

    # Sets the description on the document footer.
    Description = "T1 is before deliberation. T2 is after deliberation. Numbers not in parentheses are mean responses or percentages. Numbers in parentheses are p-values resulting from a t-test."

    # Creates the Word document for export
    WordDocument = suppressWarnings(read_docx(paste0(getwd(), Template)) %>%
                                    body_add_par("", pos = "after") %>%
                                    body_add_flextable(Crosstabs, align = "left") %>%
                                    headers_replace_all_text("Type", Type) %>%
                                    headers_replace_all_text("Title", Document_Title) %>%
                                    footers_replace_all_text("Description", Description))
  }

  if(Type == "Regressions"){

    # Gets the headers and sets them as the column names
    ColumnNames = Crosstabs[1,]
    Crosstabs = Crosstabs[-(1),]

    # Converts the dataframes to flextables
    Crosstabs = flextable(Crosstabs)

    # Deletes excess header.
    Crosstabs = delete_part(Crosstabs, part = "header")

    # Adds correct header row
    Crosstabs = add_header_row(Crosstabs, top = TRUE, values = ColumnNames)

    # Sets font size
    Crosstabs = fontsize(Crosstabs, size = 5, part = "all")

    # Sets font type
    Crosstabs = font(Crosstabs, fontname = "Arial", part = "all")

    # Sets the borders on the tables
    Crosstabs = border_remove(Crosstabs)
    Crosstabs = border(Crosstabs, j = 3:ColumnNumbers, i = 1, border.top = fp_border(color = "grey", style = "solid", width = 0.5), part = "body")
    Crosstabs = border(Crosstabs, i = 2:(RowNumbers-1), border.right = fp_border(color = "grey", style = "solid", width = 0.5), part = "body")
    Crosstabs = border(Crosstabs, j = 2:ColumnNumbers, border.bottom = fp_border(color = "grey", style = "solid", width = 0.5), border.right = fp_border(color = "grey", style = "solid", width = 0.5), part = "body")
    Crosstabs = border(Crosstabs, j = 3:ColumnNumbers, i = 1, border.bottom = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, j = 2, i = 2:(RowNumbers-1), border.right = fp_border(color = "black", style = "solid", width = 1), part = "body")

    # Bolds the text
    Crosstabs = bold(Crosstabs, j = 1, bold = TRUE, part = "body")
    Crosstabs = bold(Crosstabs, bold = TRUE, part = "header")

    # Aligns the text
    Crosstabs = align(Crosstabs, j = 1, align = "left", part = "all")
    Crosstabs = align(Crosstabs, j = 2:ColumnNumbers, align = "center", part = "all")

    # Merges cells
    Crosstabs = merge_at(Crosstabs, j = 2:ColumnNumbers, part = "header")
    Crosstabs = merge_at(Crosstabs, j = 1, i = 1:(RowNumbers-1), part = "body")

    # Rotates cell
    Crosstabs = rotate(Crosstabs, j = 1, i = 1:(RowNumbers-1), rotation = "btlr", align = "center", part = "body")

    # Sets the heights and widths of the columns and heights of the rows
    Crosstabs = width(Crosstabs, width = 0.4)
    Crosstabs = width(Crosstabs, j = 1, width = 0.4)
    Crosstabs = height(Crosstabs, height = 0.4, part = "header")
    Crosstabs = height(Crosstabs, height = 0.4, part = "body")

    # Sets the description on the document footer.
    Description = "T1 is before deliberation. T2 is after deliberation. T2-T1 is the difference between before and after deliberation. Numbers not in parentheses are correlation coefficients. Numbers in parentheses are p-values resulting from an F-test or t-test."

    # Creates the Word document for export
    WordDocument = suppressWarnings(read_docx(paste0(getwd(), Template)) %>%
                                    body_add_par("", pos = "after") %>%
                                    body_add_flextable(Crosstabs, align = "left") %>%
                                    headers_replace_all_text("Type", Type) %>%
                                    headers_replace_all_text("Title", Document_Title) %>%
                                    footers_replace_all_text("Description", Description))
  }

  if(Type == "Multiple Regressions"){

    # Gets sequential numbers for the column names
    colnames(Crosstabs) = seq(1, ncol(Crosstabs), 1)

    # Converts the dataframes to flextables
    Crosstabs = flextable(Crosstabs)

    # Deletes excess header.
    Crosstabs = delete_part(Crosstabs, part = "header")

    # Sets the heights and widths of the columns and heights of the rows
    Crosstabs = width(Crosstabs, width = 1.0)
    Crosstabs = width(Crosstabs, j = 3, width = 1.2)

    # Sets font size
    Crosstabs = fontsize(Crosstabs, size = 8, part = "all")

    # Sets font type
    Crosstabs = font(Crosstabs, fontname = "Arial", part = "all")

    # Sets the borders on the tables
    Crosstabs = border_remove(Crosstabs)
    Crosstabs = border(Crosstabs, i = (1+0*RowNumbers/6), j = 1:6, border.bottom = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, i = (1+1*RowNumbers/6), j = 1:6, border.bottom = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, i = (1+2*RowNumbers/6), j = 1:6, border.bottom = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, i = (1+3*RowNumbers/6), j = 1:6, border.bottom = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, i = (1+4*RowNumbers/6), j = 1:6, border.bottom = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, i = (1+5*RowNumbers/6), j = 1:6, border.bottom = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, i = (1+0*RowNumbers/6):(1*RowNumbers/6-1), j = 1, border.right = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, i = (1+1*RowNumbers/6):(2*RowNumbers/6-1), j = 1, border.right = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, i = (1+2*RowNumbers/6):(3*RowNumbers/6-1), j = 1, border.right = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, i = (1+3*RowNumbers/6):(4*RowNumbers/6-1), j = 1, border.right = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, i = (1+4*RowNumbers/6):(5*RowNumbers/6-1), j = 1, border.right = fp_border(color = "black", style = "solid", width = 1), part = "body")
    Crosstabs = border(Crosstabs, i = (1+5*RowNumbers/6):(6*RowNumbers/6-1), j = 1, border.right = fp_border(color = "black", style = "solid", width = 1), part = "body")

    # Bolds the text
    Crosstabs = bold(Crosstabs, j = 1, bold = TRUE, part = "body")
    Crosstabs = bold(Crosstabs, j = 8, bold = TRUE, part = "body")
    Crosstabs = bold(Crosstabs, i = (1+0*RowNumbers/6), j = 1:6, bold = TRUE, part = "body")
    Crosstabs = bold(Crosstabs, i = (1+1*RowNumbers/6), j = 1:6, bold = TRUE, part = "body")
    Crosstabs = bold(Crosstabs, i = (1+2*RowNumbers/6), j = 1:6, bold = TRUE, part = "body")
    Crosstabs = bold(Crosstabs, i = (1+3*RowNumbers/6), j = 1:6, bold = TRUE, part = "body")
    Crosstabs = bold(Crosstabs, i = (1+4*RowNumbers/6), j = 1:6, bold = TRUE, part = "body")
    Crosstabs = bold(Crosstabs, i = (1+5*RowNumbers/6), j = 1:6, bold = TRUE, part = "body")

    # Aligns the text
    Crosstabs = align(Crosstabs, align = "center", part = "all")
    Crosstabs = align(Crosstabs, j = 1, align = "left", part = "all")
    Crosstabs = align(Crosstabs, j = 8:9, align = "left", part = "all")

    # Sets the description on the document footer.
    Description = "T1 is before deliberation. T2 is after deliberation. T2-T1 is the difference between before and after deliberation."

    # Creates the Word document for export
    WordDocument = suppressWarnings(read_docx(paste0(getwd(), Template)) %>%
                                    body_add_par("", pos = "after") %>%
                                    body_add_flextable(Crosstabs, align = "left") %>%
                                    headers_replace_all_text("Type", Type) %>%
                                    headers_replace_all_text("Title", Document_Title) %>%
                                    footers_replace_all_text("Description", Description))
  }

  if(Type == "Demographics"){

    # Gets the column names of the crosstabs
    ColumnNames = Crosstabs[4,]

    # Gets parts of the crosstab to print
    SampleSizes = Crosstabs[2,]
    Crosstabs = Crosstabs[5:nrow(Crosstabs),]

    # Gets the headers and sets them as the column names
    colnames(Crosstabs) = Crosstabs[1,]

    # Gets sequential numbers for the column names
    colnames(SampleSizes) = seq(1, ncol(Crosstabs), 1)
    colnames(Crosstabs) = seq(1, ncol(Crosstabs), 1)

    # Converts the dataframes to flextables
    SampleSizes = flextable(SampleSizes)
    Crosstabs = flextable(Crosstabs)

    # Deletes excess header.
    Crosstabs = delete_part(Crosstabs, part = "header")

    # Removes the excess headers
    SampleSizes = delete_part(SampleSizes, part = "header")

    # Adds correct header row
    Crosstabs = add_header_row(Crosstabs, top = TRUE, values = ColumnNames)

    # Sets the widths of the columns
    SampleSizes = width(SampleSizes, width = 1.5)
    Crosstabs = width(Crosstabs, width = 1.5)
    SampleSizes = width(SampleSizes, j = 2, width = 2.5)
    Crosstabs = width(Crosstabs, j = 2, width = 2.5)

    # Sets font size
    SampleSizes = fontsize(SampleSizes, size = 8, part = "all")
    Crosstabs = fontsize(Crosstabs, size = 8, part = "all")

    # Sets font type
    SampleSizes = font(SampleSizes, fontname = "Arial", part = "all")
    Crosstabs = font(Crosstabs, fontname = "Arial", part = "all")

    # Sets the borders on the tables
    SampleSizes = border_remove(SampleSizes)
    Crosstabs = border_remove(Crosstabs)
    Crosstabs = border(Crosstabs, border.bottom = fp_border(color = "black", style = "solid", width = 1), part = "header")
    Crosstabs = border(Crosstabs, border.right = fp_border(color = "black", style = "solid", width = 1), part = "body")

    # Bolds the text
    SampleSizes = bold(SampleSizes, bold = TRUE, part = "all")
    Crosstabs = bold(Crosstabs, j = 1, bold = TRUE, part = "body")
    Crosstabs = bold(Crosstabs, bold = TRUE, part = "header")

    # Aligns the text
    SampleSizes = align(SampleSizes, align = "center", part = "all")
    Crosstabs = align(Crosstabs, j = 1:2, align = "left", part = "all")
    Crosstabs = align(Crosstabs, j = 3:ColumnNumbers, align = "center", part = "all")
    Crosstabs = align(Crosstabs, align = "center", part = "header")

    # Sets the description on the document footer.
    Description = "T1 is before deliberation. T2 is after deliberation. Numbers not in parentheses are rounded counts. Numbers in parentheses are percentages or p-values resulting from a chi-squared test."

    # Creates the Word document for export
    if(!class(try((suppressWarnings(read_docx(paste0(getwd(), Template)) %>%
                                    body_add_par("", pos = "after") %>%
                                    body_add_flextable(SampleSizes, align = "left") %>%
                                    body_add_par("", pos = "after") %>%
                                    body_add_flextable(Crosstabs, align = "left") %>%
                                    headers_replace_all_text("Type", Type) %>%
                                    headers_replace_all_text("Title", Document_Title) %>%
                                    footers_replace_all_text("Description", Description))),silent = TRUE)) == "try-error"){
      WordDocument = suppressWarnings(read_docx(paste0(getwd(), Template)) %>%
                                      body_add_par("", pos = "after") %>%
                                      body_add_flextable(SampleSizes, align = "left") %>%
                                      body_add_par("", pos = "after") %>%
                                      body_add_flextable(Crosstabs, align = "left") %>%
                                      headers_replace_all_text("Type", Type) %>%
                                      headers_replace_all_text("Title", Document_Title) %>%
                                      footers_replace_all_text("Description", Description))
    }
    else {print("Cannot export Word version due to large file size or template error.")}
  }

  # Exports the Word document
  if(exists("WordDocument")){
    print(WordDocument, target = paste0(getwd(), Outputs, "/", File_Name))
    print(noquote(paste("Exported:", File_Name)))
  }}}}
