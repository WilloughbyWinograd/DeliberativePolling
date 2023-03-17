#' Returns tables and reports in Microsoft Word and Excel.
Results = function(
    Dataset,
    Exports,
    Group_1 = c(Treatment, Time, Weight),
    Group_2 = c(Treatment, Time, Weight),
    Template,
    API,
    Significance
    ){
  
  #' Returns tables and reports in Microsoft Word and Excel.
  #' 
  #' @description
  #' This function returns three types of outputs.
  #' 1. Tables on all nominal data.
  #' 2. Tables on all ordinal data by all nominal data.
  #' 3. A report on all ordinal data by select nominal data.
  #' This function required a dataset formatted using the "Dataset" function.
  #' 
  #' @param Dataset is the dataset file.
  #' @param Exports is the optional location of where files should be printed to.
  #' @param Group_1 is a vector containing the treatment group "Participants", time "T1", and weighting variable "Weight 1".
  #' @param Group_2 is a vector containing the treatment group "Nonparticipants", time "T2", and weighting variable "Weight 2".
  #' @param Template is the Word document that serves as the base from which documents will be printed.
  #' @param API is an optional API key to call ChatGPT to improve report phrasing.
  #' @param Significant is an optional value for alpha above which significant tests will not be reported.
  #' 
  #' @export
  
  if(missing(Exports)){Exports = ""}
  if(missing(Significance)){
    Only_Significant = FALSE
    Alpha = 0.05
    } else {
    Only_Significant = TRUE
    Alpha = Significance
  }
  if(missing(API)){API = "None"}
  
  Ordinal(Dataset, Template, Outputs = Exports, Group1 = Group_1, Group2 = Group_2, Alpha, Only_Significant)
  Nominal(Dataset, Template, Outputs = Exports, Group1 = Group_1, Group2 = Group_2, Alpha, Only_Significant, Only_Means = FALSE, API_Key = API)
}

#' Returns list of data frames.
Dataset = function(
    Codebook,
    Datasets
    ){
  
  #' Returns list of data frames.
  #' 
  #' @description
  #' This function returns a list of data frames and the codebook formatted for use in the "Results" function.
  #' 
  #' @param Codebook is a sheet containing a correctly formatted codebook for the dataset file.
  #' @param Datasets are the unformatted datasets to be formatted to match the codebook.
  #' 
  #' @export
  
  for(Dataset_Number in 1:length(Datasets)){
    Dataset = Datasets[[Dataset_Number]]
    OG_Names = names(Dataset)
    names(Dataset) = paste(seq(1,length(OG_Names)), OG_Names)
    Dataset = Dataset %>% mutate_all(funs(str_replace_all(., paste0("[", paste(c("&", "$", "£", "+"), collapse = ""), "]"), "_")))
    names(Dataset) = OG_Names
    Datasets[[Dataset_Number]] = Dataset
  }
  
  Codebook_Demographics = Codebook[which(Codebook[1,] == "Ordinal")]
  
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
        
        if(Type == "Ordinal"){
          
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
        if(Type == "Nominal"){
          
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
      
      Dataset_Column = suppressWarnings((as.matrix(Dataset_Column)))
      if(Type != "Text"){Dataset_Column = suppressWarnings(as.numeric(as.matrix(Dataset_Column)))} # Forces numeric
      Dataset_Column = as_tibble(Dataset_Column)
      colnames(Dataset_Column) = Name
      Dataset_New = cbind(Dataset_New, Dataset_Column)
    }
    
    Datasets[[Dataset_Number]] = Dataset_New[order(Dataset_New$`Identification Number`),]
    
  }
  
  return(list(Datasets, Codebook))
}



#' Creates a crosstab in Excel and Word comparing opinions.
Nominal = function(Dataset, Template, Outputs, Group1, Group2, Alpha, Only_Significant, Only_Means, API_Key){
  
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
  Overall[1,] = "Ordinal"
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
  Demographics = (which(Codebook[1,] == "Ordinal"))
  Questions = (which(Codebook[1,] == "Nominal"))
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
      File_Name = paste("Tables", "Nominal", Name_Group, Weights, Demographic_Category, sep = " - ")
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
      Export_Word(Crosstabs, Outputs, File_Name, Name_Group, Template, Document_Title, Type = "Nominal", Demographic_Category, Codebook)
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
      Export_Excel(Crosstabs, Outputs, File_Name, Name_Group)
    }
    
    # Runs a pre-defined function that creates a report from the crosstabs.
    {
      if(Group1 == Group2){
        Export_Report(Crosstabs, Outputs, File_Name, Name_Group, Template, Document_Title, Type = "Report", Demographic_Category, API_Key, Group1)
      }
    }
    
  }
}

#' Creates a crosstab in Excel and Word comparing demographics
Ordinal = function(Dataset, Template, Outputs, Group1, Group2, Alpha, Only_Significant){
  
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
  Demographics = (which(Codebook[1,] == "Ordinal"))
  Questions = (which(Codebook[1,] == "Nominal"))
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
    colnames(VerticalHeader) = c("Category", "Group")
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
    File_Name = paste("Tables", "Ordinal", Name_Group, Weights, sep = " - ")
    Document_Title = gsub(" by Overall", "", paste(paste(Name_Group, ", Weighted by ", Weights, sep = "")))
    Document_Title = gsub("Weighted by Unweighted", "Unweighted", Document_Title)
  }
  
  # Runs a pre-defined function that creates Excel spreadsheets from the crosstabs.
  {
    Export_Excel(Crosstabs, Outputs, File_Name, Name_Group)
  }
  
  # Runs a pre-defined function that creates Word documents from the crosstabs.
  {
    Export_Word(Crosstabs, Outputs, File_Name, Name_Group, Template, Document_Title, Type = "Ordinal", Demographic_Category, Codebook)
  }
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

#' Exports a crosstab in Excel.
Export_Excel = function(Crosstabs, Outputs, File_Name, Name_Group){
  
  # Gets the file name.
  File_Name = paste(File_Name, ".xlsx", sep = "")
  
  # Ensures the sheet name is the proper length.
  if(nchar(Name_Group)>31){
    Name_Group = paste(strtrim(Name_Group,28), "...", sep = "")
  }
  
  # Exports crosstab to Excel file
  write.xlsx(Crosstabs, paste0(getwd(), Outputs, "/", File_Name), sheetName = Name_Group, showNA = F, colNames = F, rownames = F, overwrite = T)
  
  # Notifies of document export
  print(noquote(paste("Exported:", File_Name)))
}

#' Exports a crosstab in Word.
Export_Word = function(Crosstabs, Outputs, File_Name, Name_Group, Template, Document_Title, Type, Demographic_Category, Codebook){
  
  {
    if(ncol(Crosstabs)<14){
      
      File_Name = paste(File_Name, ".docx", sep = "")
      
      # Gets the number of columns
      ColumnNumbers = ncol(Crosstabs)
      RowNumbers = nrow(Crosstabs)
      
      if(Type == "Nominal"){
        
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
        Description = "Numbers not in parentheses are mean responses or percentages. Numbers in parentheses are p-values resulting from a t-test."
        
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
                                            headers_replace_all_text("Title", Type) %>%
                                            headers_replace_all_text("Subtitle", Document_Title) %>%
                                            footers_replace_all_text("Details", Description))
        }
        else {print("Cannot export Word version due to large file size or template error.")}
      }
      
      if(Type == "Ordinal"){
        
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
        Description = "Numbers not in parentheses are rounded counts. Numbers in parentheses are percentages or p-values resulting from a chi-squared test."
        
        # Creates the Word document for export
        if(!class(try((suppressWarnings(read_docx(paste0(getwd(), Template)) %>%
                                        body_add_par("", pos = "after") %>%
                                        body_add_flextable(SampleSizes, align = "left") %>%
                                        body_add_par("", pos = "after") %>%
                                        body_add_flextable(Crosstabs, align = "left") %>%
                                        headers_replace_all_text("Title", Type) %>%
                                        headers_replace_all_text("Subtitle", Document_Title) %>%
                                        footers_replace_all_text("Details", Description))),silent = TRUE)) == "try-error"){
          WordDocument = suppressWarnings(read_docx(paste0(getwd(), Template)) %>%
                                            body_add_par("", pos = "after") %>%
                                            body_add_flextable(SampleSizes, align = "left") %>%
                                            body_add_par("", pos = "after") %>%
                                            body_add_flextable(Crosstabs, align = "left") %>%
                                            headers_replace_all_text("Title", Type) %>%
                                            headers_replace_all_text("Subtitle", Document_Title) %>%
                                            footers_replace_all_text("Details", Description))
        }
        else {print("Cannot export Word version due to large file size or template error.")}
      }
      
      # Exports the Word document
      if(exists("WordDocument")){
        print(WordDocument, target = paste0(getwd(), Outputs, "/", File_Name))
        print(noquote(paste("Exported:", File_Name)))
      }}}}

#' Exports a report on crosstab in Word.
Export_Report = function(Crosstabs, Outputs, File_Name, Name_Group, Template, Document_Title, Type, Demographic_Category, API_Key, Group1){
  
  ReturnText = function(Change){
    if(Change<0){
      return("decreased")
    }
    if(Change>0){
      return("increased")
    }
    if(Change==0){
      return("did not change")
    }
  }
  
  Stance = function(support_percent, opposition_percent, Stance_Negative, Stance_Positive){
    if (support_percent >= 0.67) {
      return(paste("a supermajority selected", Stance_Positive))
    } else if (support_percent > 0.5) {
      return(paste("a majority selected", Stance_Positive))
    } else if (opposition_percent > 0.67) {
      return(paste("a supermajority selected", Stance_Negative))
    } else if (opposition_percent >= 0.5) {
      return(paste("a majority selected", Stance_Negative))
    } else {
      return("there was no majority")
    }
  }
  
  Stance_Percentage = function(support_percent, opposition_percent, Stance_Negative, Stance_Positive){
    if (support_percent >= 0.67) {
      return(paste("a supermajority selected", Stance_Positive, paste0("(",FormatPercentage(support_percent),")")))
    } else if (support_percent > 0.5) {
      return(paste("a majority selected", Stance_Positive, paste0("(",FormatPercentage(support_percent),")")))
    } else if (opposition_percent > 0.67) {
      return(paste("a supermajority selected", Stance_Negative, paste0("(",FormatPercentage(opposition_percent),")")))
    } else if (opposition_percent >= 0.5) {
      return(paste("a majority selected", Stance_Negative, paste0("(",FormatPercentage(opposition_percent),")")))
    } else {
      return(paste("there was no majority for", Stance_Positive, paste0("(",FormatPercentage(support_percent),")"), "or", Stance_Negative, paste0("(",FormatPercentage(opposition_percent),")")))
    }
  }
  
  Convert_Percentage = function(Percentage){
    return(as.numeric(sub("%", "", Percentage))/100)
  }
  
  FormatPercentage = function(Percent){
    return(paste0(abs(Percent*100),"%"))
  }
  
  
  # Function to add "P=" inside parenthesis
  add_P_in_parenthesis <- function(s) {
    # Check if there are parentheses in the string
    if (grepl("\\(.*\\)", s)) {
      # Add "P=" inside the parentheses
      return(gsub("(\\()(.*)(\\))", "\\1P = \\2\\3", s))
    } else {
      # Return the original string if no parentheses are present
      return(s)
    }
  }
  
  extract_first_numeric <- function(s) {
    # Split the string using space or opening parenthesis as the delimiter
    split_string <- strsplit(s, split = " |\\(")
    
    # Extract the first part of the split string (numeric value)
    first_part <- split_string[[1]][1]
    
    # Convert the first part to a numeric value
    numeric_value <- as.numeric(first_part)
    
    return(numeric_value)
  }
  
  Legend = Crosstabs[2:(which(Crosstabs[2] == "Prompt and Responses")-6), 2]
  SampleSizes = Crosstabs[(which(Crosstabs[2] == "Prompt and Responses")-2), 3:(2+length(Legend))]
  
  if(length(Legend) < 6){
  Tabs = Crosstabs[(2+which(Crosstabs[2] == "Prompt and Responses")):nrow(Crosstabs), ]
  
  Questions = unique(Tabs$ID)
  Questions = Questions[nzchar(Questions)]
  
  Subject = Group1
  
  WordDocument <- suppressWarnings(read_docx(paste0(getwd(), Template)))
  
  for(Question in Questions){
    
    Tab = Tabs[which(Tabs$ID == Question):(which(Tabs$ID == Question)+4),]
    
    Stance_Negative = paste0("\"", Tab[2,2], "\"")
    Stance_Positive = paste0("\"", Tab[4,2], "\"")
    
    Text = (Crosstabs[which(Crosstabs == Question), 2])
    
    Lines = paste0(tools::toTitleCase(Subject), " were asked to respond to the statement, \"", gsub("\\.", "", Text), "\"", ".")
    Subject = tolower(Subject)
    
    # Create the matrix with the specified headings
    Data <- matrix(ncol=3, nrow=0, dimnames=list(NULL, c("Group", paste("Selected", Stance_Positive, "Before Deliberation"), paste("Selected", Stance_Positive, "After Deliberation"))))
    
    for(Group in Legend){
      
      Index = which(Legend == Group)
      
      Group_Original = Group
      
      if(Group == "Total"){
        Group = Subject
      } else {
        Group = paste("those who selected", Group)
      }
      
      Group_Results = Tab[seq(Index + 2, length.out = 3, by = 1+length(Legend))]
      
      Beg_Mean = Group_Results[1,1]
      End_Mean = Group_Results[1,2]
      Change_Mean = Group_Results[1,3]
      
      Beg_Support = Convert_Percentage(Group_Results[4,1])
      End_Support = Convert_Percentage(Group_Results[4,2])
      Change_Support = Convert_Percentage(Group_Results[4,3])
      
      Beg_Opposition = Convert_Percentage(Group_Results[2,1])
      End_Opposition = Convert_Percentage(Group_Results[2,2])
      Change_Opposition = Convert_Percentage(Group_Results[2,3])
      
      if(!(is.na(Change_Mean) || Change_Mean == "")){
        
        Data = rbind(Data, c(Group_Original, Beg_Support, End_Support))
        
        if(extract_first_numeric(Change_Mean) != 0){ Line1 = paste("Among", paste0(Group),  paste0("(", SampleSizes[Index], ")", ","), "the mean rating", ReturnText(Change_Mean), "by", add_P_in_parenthesis(Change_Mean))} else {
          Line1 = paste("Among", paste0(Group,","), "the mean rating did not change")}
        
        if(Stance(End_Support, End_Opposition, Stance_Negative, Stance_Positive) == Stance(Beg_Support, Beg_Opposition, Stance_Negative, Stance_Positive)){
          Line2 = paste0("After deliberation ", Stance_Percentage(End_Support, End_Opposition, Stance_Negative, Stance_Positive), " among this group, similar to before deliberation.")
        } else {
          Line2 = paste0("Before deliberation ", Stance_Percentage(Beg_Support, Beg_Opposition, Stance_Negative, Stance_Positive), " among this group, while after deliberation ", Stance_Percentage(End_Support, End_Opposition, Stance_Negative, Stance_Positive), ".")
        }
        
        Lines = paste(Lines, paste0(Line1, ". ", Line2))
      }}
    
    if(API_Key != "None"){
      # Calls the ChatGPT API with the given prompt and returns the answer
      ask_chatgpt <- function(prompt) {
        response <- POST(
          url = "https://api.openai.com/v1/chat/completions", 
          add_headers(Authorization = paste("Bearer", API_Key)),
          content_type_json(),
          encode = "json",
          body = list(
            model = "gpt-3.5-turbo",
            messages = list(list(
              role = "user", 
              content = prompt
            ))
          )
        )
        str_trim(content(response)$choices[[1]]$message$content)
      }
      
      Lines <- ask_chatgpt(paste("Rephase the following to sound less structured and more human, but do not change the numbers or conclusions. Keep the numbers the same. Don't remove the word deliberation:", Lines))
    }
    if(length(Lines) == 0){Lines = "API error. Ensure there is a valid API key for ChatGPT."}
    
    Data = as.data.frame(Data)
    Names = Data[,1]
    Data=Data[,-1]
    Data = apply(Data, 2, as.numeric)
    Data = as.data.frame(Data)
    
    if(Demographic_Category != "Overall"){
      rownames(Data) = Names
    } else {
      Data = t(Data)
      rownames(Data) = Names
    }
    
    Data = t(Data)
    
    Data <- as.data.frame(t(Data))
    
    Data1 = Data[1]
    Data2 = Data[2]
    
    names(Data1)[1] = paste("Selected", Stance_Positive, "(%)")
    names(Data2)[1] = paste("Selected", Stance_Positive, "(%)")
    
    rownames(Data1) = paste(rownames(Data1), "(Before)")
    rownames(Data2) = paste(rownames(Data2), "(After)")
    
    # Interlace the data frames
    Combined <- rbind(Data1, Data2)
    
    Combined = cbind(rownames(Combined), Combined)
    
    # Interlace rows based on row index
    Combined = Combined[rev(order(Combined$`rownames(Combined)`)), ]
    
    Data = Combined
    Data = Data[2]
    
    Group = row.names(Data)
    Data = as.data.frame(cbind(Group, Data))
    
    table <- flextable(Data)
    
    table = width(table, width = 2)
    table = width(table, j = 2, width = 4.5)
    
    table <- align(table, align = "right", part = "body", j = 1)
    table <- align(table, align = "left", part = "body", j = 2)
    table <- align(table, align = "center", part = "header")
    
    # Add minibars to the table
    table <- compose(table, j = 2, value = as_paragraph(minibar(as.vector(unlist(Data[2])), barcol = "black", width = 4.5)))

    Plot = table
    
    Description = ""
    
    WordDocument <- suppressWarnings(WordDocument %>% 
      body_add_par(Question, style = "heading 1", pos = "after") %>%
      body_add_par(Lines, pos = "after") %>%
      body_add_par("", pos = "after") %>%
      body_add_flextable(Plot, align = "center") %>%
      headers_replace_all_text("Title", Type) %>%
      headers_replace_all_text("Subtitle", Document_Title) %>%
      footers_replace_all_text("Details", Description) %>%
      body_add_break())
    
  }
  
  File_Name = paste0(gsub("Tables", "Report", File_Name), ".docx")
  
  print(WordDocument, target = paste0(getwd(), Outputs, "/", File_Name))
  print(noquote(paste("Exported:", File_Name)))
  
}}