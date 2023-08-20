
#Test

import pandas as pd

dataset = pd.read_excel("Python/Dataset.xlsx")

results(dataset="/Inputs/Dataset.xlsx",
        group_1=["Treatment", "T1", "weight_a"],
        group_2=["Control", "T2", "weight_a"],
        template="Template.docx" )

excel_comparison = pd.read_excel("Python/(Original) Tables - Ordinal - Treatment at T1 v. Control at T2 - weight_a - Gender.xlsx")
excel_new = pd.read_excel("Python/(Tables - Ordinal - Treatment at T1 v. Control at T2 - weight_a - Gender.xlsx")

# Check if both dataframes are of the same shape
if excel_comparison.shape != excel_new.shape:
    print("Dataframes are of different shapes!")
else:
    # Compare the two dataframes and get a boolean dataframe
    df_diff = excel_comparison.eq(excel_new)

    # Find where the dataframes are not equal
    rows, cols = np.where(df_diff == False)
    
    for row, col in zip(rows, cols):
        print(f"Difference at Row {row+1}, Column {col+1}")
        print(f"Original: {excel_comparison.iloc[row, col]}")
        print(f"New: {excel_new.iloc[row, col]}")
        print("-" * 40)

######## Above is for testing ########

import numpy as np
from scipy.stats import chi2_contingency

def Test_Chi_Squared(ValuetoMark, Group1, Group2):
    # Organizes the inputs.
    Observed_Expected = np.column_stack((Group1, Group2))

    # Removes all zero categories.
    Observed_Expected = Observed_Expected[~np.all(Observed_Expected == 0, axis=1)]

    # Performs a chi-squared test and gets the p-value.
    _, PValue, _, _ = chi2_contingency(Observed_Expected)

    if not np.isnan(PValue):
        # If specified, always adds p-values.
        if not Only_Significant:
            # Adds p-values
            ValuetoMark = f"{ValuetoMark} ({round(PValue, 3):.3f})"

        # If specified, only adds p-values less than or equal to alpha specified.
        if Only_Significant and PValue <= Alpha:
            # Adds p-values
            ValuetoMark = f"{ValuetoMark} ({round(PValue, 3):.3f})"

        # If a warning is generated due to an expected number less than 5, a note is added to the p-value.
        if np.any(Group2 < 5):
            ValuetoMark = f"{ValuetoMark} [Warning: P-value may be incorrect, as at least one expected value is less than 5.]"

    return ValuetoMark

import numpy as np


def Input_Test(dataset, name):
    # Check to see if the named column exists in the dataset
    if name not in dataset.columns:
        # If the column does not exist in the dataset, raise an error
        raise ValueError(f"There is no data named {name}.")


def Responses_to_Text(Responses, ColumnName, Codebook):
    # Converts column name to string.
    ColumnName = str(ColumnName)

    # Gets the text value
    Response_Negative = str(Codebook[4, Codebook.columns.get_loc(ColumnName)])
    Response_Neutral = str(Codebook[5, Codebook.columns.get_loc(ColumnName)])
    Response_Positive = str(Codebook[6, Codebook.columns.get_loc(ColumnName)])

    Scale = str(Codebook[3, Codebook.columns.get_loc(ColumnName)])

    # Condenses the 0-100 scale.
    if Scale == "0 to 100":

        # Ensures there are no responses outside of the scale.
        if (np.nanmax(np.nan_to_num(Responses)) > 100) or (np.nanmin(np.nan_to_num(Responses)) < 0):
            raise ValueError(f"There is a response outside of the {Scale} scale specified in the codebook for {ColumnName}")

        # Replaces the numeric responses with text responses.
        Responses[Responses > 66] = Response_Positive
        Responses[(Responses < 67) & (Responses > 33)] = Response_Neutral
        Responses[Responses < 33] = Response_Negative
    elif Scale == "0 to 10":

        # Ensures there are no responses outside of the scale.
        if (np.nanmax(np.nan_to_num(Responses)) > 10) or (np.nanmin(np.nan_to_num(Responses)) < 0):
            raise ValueError(f"There is a response outside of the {Scale} scale specified in the codebook for {ColumnName}")

        # Replaces the numeric responses with text responses.
        Responses[Responses > 5] = Response_Positive
        Responses[Responses == 5] = Response_Neutral
        Responses[Responses < 5] = Response_Negative
    elif Scale == "1 to 5":

        # Ensures there are no responses outside of the scale.
        if (np.nanmax(np.nan_to_num(Responses)) > 5) or (np.nanmin(np.nan_to_num(Responses)) < 1):
            raise ValueError(f"There is a response outside of the {Scale} scale specified in the codebook for {ColumnName}")

        # Replaces the numeric responses with text responses.
        Responses[Responses > 3] = Response_Positive
        Responses[Responses == 3] = Response_Neutral
        Responses[Responses < 3] = Response_Negative
    elif Scale == "Difference":

        # Ensures there are no responses outside of the scale.
        if (np.nanmax(np.nan_to_num(Responses)) > 10) or (np.nanmin(np.nan_to_num(Responses)) < -10):
            raise ValueError(f"There is a response outside of the {Scale} scale specified in the codebook for {ColumnName}")

        # Replaces the numeric responses with text responses.
        Responses[Responses > 0] = Response_Positive
        Responses[Responses == 0] = Response_Neutral
        Responses[Responses < 0] = Response_Negative
    elif Scale == "1 to 3":

        # Ensures there are no responses outside of the scale.
        if (np.nanmax(np.nan_to_num(Responses)) > 3) or (np.nanmin(np.nan_to_num(Responses)) < 1):
            raise ValueError(f"There is a response outside of the {Scale} scale specified in the codebook for {ColumnName}")

        # Replaces the numeric responses with text responses.
        Responses[Responses == 1] = Response_Negative
        Responses[Responses == 2] = Response_Neutral
        Responses[Responses == 3] = Response_Positive
    elif Scale == "1 to 4":

        # Ensures there are no responses outside of the scale.
        if (np.nanmax(np.nan_to_num(Responses)) > 4) or (np.nanmin(np.nan_to_num(Responses)) < 1):
            raise ValueError(f"There is a response outside of the {Scale} scale specified in the codebook for {ColumnName}")

        # Replaces the numeric responses with text responses.
        Responses[Responses == 1] = Response_Negative
        Responses[Responses == 2] = Response_Negative
        Responses[Responses == 3] = Response_Positive
        Responses[Responses == 4] = Response_Positive
    elif Scale == "0 to 1":

        # Ensures there are no responses outside of the scale.
        if (np.nanmax(np.nan_to_num(Responses)) > 1) or (np.nanmin(np.nan_to_num(Responses)) < 0):
            raise ValueError(f"There is a response outside of the {Scale} scale specified in the codebook for {ColumnName}")

        # Replaces the numeric responses with text responses.
        Responses[Responses == 1] = Response_Positive
        Responses[Responses == 0] = Response_Negative
    else:
        raise ValueError(f"Scale {Scale} for {ColumnName} not recognized. Please pick a pre-defined scale such as 0 to 100, 0 to 10, 1 to 5, 1 to 3, or 0 to 1.")

    # Returns the formatted responses.
    return Responses
import numpy as np
from scipy.stats import chi2_contingency

def Test_Chi_Squared(ValuetoMark, Group1, Group2):
    # Organizes the inputs.
    Observed_Expected = np.column_stack((Group1, Group2))

    # Removes all zero categories.
    Observed_Expected = Observed_Expected[~np.apply_along_axis(lambda y: np.all(y == 0), 1, Observed_Expected)]

    # Performs a chi-squared test and gets the p-value.
    _, PValue, _, _ = chi2_contingency(Observed_Expected, simulate_p_value=False)

    if not np.isnan(PValue):
        # If specified, always adds p-values.
        if not Only_Significant:
            # Adds p-values
            ValuetoMark = f"{ValuetoMark} ({PValue:.3f})"

        # If specified, only adds p-values less than or equal to alpha specified.
        if Only_Significant and PValue <= Alpha:
            # Adds p-values
            ValuetoMark = f"{ValuetoMark} ({PValue:.3f})"

        # If a warning is generated due to an expected number less than 5, a note is added to the p-value.
        if np.any(Group2 < 5):
            ValuetoMark = f"{ValuetoMark} [Warning: P-value may be incorrect, as at least one expected value is less than 5.]"

    return ValuetoMark

def Export_Excel(Crosstabs, Outputs, File_Name, Name_Group):
    # Gets the file name.
    File_Name = File_Name + ".xlsx"

    # Ensures the sheet name is the proper length.
    if len(Name_Group) > 31:
        Name_Group = Name_Group[:28] + "..."

    # Exports crosstab to Excel file
    Crosstabs.to_excel(Outputs + "/" + File_Name, sheet_name=Name_Group, index=False, header=False)

    # Notifies of document export
    print(f"Exported: {File_Name}")
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def Export_Word(Crosstabs, Outputs, File_Name, Name_Group, Template, Document_Title, Type, Demographic_Category, Codebook):
    if Crosstabs.shape[1] < 14:
        File_Name = File_Name + ".docx"
        ColumnNumbers = Crosstabs.shape[1]
        RowNumbers = Crosstabs.shape[0]

        if Type == "Ordinal":
            # Creates a legend.
            Legend = Codebook.iloc[3:, Codebook.columns.get_loc(Demographic_Category)]
            Legend = Legend.dropna()
            Legend = pd.DataFrame(Legend)
            Legend = pd.concat([pd.DataFrame([["T", "Total"]]), Legend], ignore_index=True)

            # Gets the column names of the crosstabs
            ColumnNames = list(Crosstabs.iloc[4])

            # Gets parts of the crosstab to print
            Titles = Crosstabs.iloc[0]
            SampleSizes = Crosstabs.iloc[2]
            Crosstabs = Crosstabs.iloc[4:]

            # Converts the dataframes to tables
            Legend = Legend.to_html(index=False, header=False)
            Titles = Titles.to_frame().T.to_html(index=False, header=False)
            SampleSizes = SampleSizes.to_frame().T.to_html(index=False, header=False)
            Crosstabs = Crosstabs.to_html(index=False, header=False)

            # Creates the Word document for export
            document = Document(Template)
            document.add_paragraph()
            document.add_paragraph(Legend)
            document.add_paragraph()
            document.add_paragraph(Titles)
            document.add_paragraph()
            document.add_paragraph(SampleSizes)
            document.add_paragraph()
            document.add_paragraph(Crosstabs)
            document.save(Outputs + "/" + File_Name)

            # Notifies of document export
            print("Exported:", File_Name)

        elif Type == "Nominal":
            # Gets the column names of the crosstabs
            ColumnNames = list(Crosstabs.iloc[3])

            # Gets parts of the crosstab to print
            SampleSizes = Crosstabs.iloc[1]
            Crosstabs = Crosstabs.iloc[4:]

            # Converts the dataframes to tables
            SampleSizes = SampleSizes.to_frame().T.to_html(index=False, header=False)
            Crosstabs = Crosstabs.to_html(index=False, header=False)

            # Creates the Word document for export
            document = Document(Template)
            document.add_paragraph()
            document.add_paragraph(SampleSizes)
            document.add_paragraph()
            document.add_paragraph(Crosstabs)
            document.save(Outputs + "/" + File_Name)

            # Notifies of document export
            print("Exported:", File_Name)

    else:
        print("Cannot export Word version due to large file size or template error.")
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.oxml.shared import OxmlElement
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
import pandas as pd
import requests
import urllib.parse

def export_report(Crosstabs, Outputs, File_Name, Name_Group, Template, Document_Title, Type, Demographic_Category, API_Key, Group1):

    def return_text(change):
        if change < 0:
            return "decreased"
        elif change > 0:
            return "increased"
        elif change == 0:
            return "did not change"

    def stance(support_percent, opposition_percent, stance_negative, stance_positive):
        if support_percent >= 0.67:
            return f"a supermajority selected {stance_positive.lower()}"
        elif support_percent > 0.5:
            return f"a majority selected {stance_positive.lower()}"
        elif opposition_percent > 0.67:
            return f"a supermajority selected {stance_negative.lower()}"
        elif opposition_percent >= 0.5:
            return f"a majority selected {stance_negative.lower()}"
        else:
            return "there was no majority"

    def stance_percentage(support_percent, opposition_percent, stance_negative, stance_positive):
        if support_percent >= 0.67:
            return f"a supermajority selected {stance_positive.lower()} ({format_percentage(support_percent)})"
        elif support_percent > 0.5:
            return f"a majority selected {stance_positive.lower()} ({format_percentage(support_percent)})"
        elif opposition_percent > 0.67:
            return f"a supermajority selected {stance_negative.lower()} ({format_percentage(opposition_percent)})"
        elif opposition_percent >= 0.5:
            return f"a majority selected {stance_negative.lower()} ({format_percentage(opposition_percent)})"
        else:
            return f"there was no majority for {stance_positive.lower()} ({format_percentage(support_percent)}) or {stance_negative.lower()} ({format_percentage(opposition_percent)})"

    def convert_percentage(percentage):
        return float(percentage.strip('%')) / 100

    def format_percentage(percent):
        return f"{abs(percent * 100)}%"

    def add_p_in_parenthesis(s):
        if '(' in s:
            return re.sub(r'\(([^)]+)\)', r'(P = \1)', s)
        else:
            return s

    def extract_first_numeric(s):
        parts = re.split(r' |\(', s)
        return float(parts[0])

    legend = Crosstabs[2:(Crosstabs.index("Prompt and Responses")-6)][1]
    sample_sizes = Crosstabs[Crosstabs.index("Prompt and Responses")-2][2:]

    if 2 <= len(legend) < 7:
        if Demographic_Category == "Overall":
            all_n_zero = False  # test
        all_n_zero = all(all(re.search("n = 0", x) for x in sample_sizes[1:]))
        if not all_n_zero:
            tabs = Crosstabs[Crosstabs.index("Prompt and Responses"):, :]

            questions = tabs[:, 0].unique()
            questions = [q for q in questions if q]

            subject = Group1

            document = Document(Template)

            for question in questions:
                tab = tabs[tabs[:, 0] == question]
                tab = tab[:5]

                stance_negative = tab[1, 1]
                stance_positive = tab[3, 1]

                text = Crosstabs[Crosstabs == question][1]

                lines = [f"{tools.toTitleCase(subject)} were asked to respond to the statement, \"{text}\"."]
                subject = subject.lower()

                data = []
                for group in legend:
                    index = legend.index(group)
                    group_original = group
                    if group == "Total":
                        group = subject
                    else:
                        group = f"those who selected \"{group}\""

                    group_results = tab[2:, index:index+3]
                    beg_mean = group_results[0, 0]
                    end_mean = group_results[0, 1]
                    change_mean = group_results[0, 2]
                    beg_support = convert_percentage(group_results[3, 0])
                    end_support = convert_percentage(group_results[3, 1])
                    change_support = convert_percentage(group_results[3, 2])
                    beg_opposition = convert_percentage(group_results[1, 0])
                    end_opposition = convert_percentage(group_results[1, 1])
                    change_opposition = convert_percentage(group_results[1, 2])

                    if beg_support == 0 or end_support == 0:
                        beg_support = convert_percentage(group_results[2, 0])
                        end_support = convert_percentage(group_results[2, 1])
                        change_support = convert_percentage(group_results[2, 2])
                        stance_positive = tab[2, 1]

                    if not pd.isnull(change_mean) and change_mean != "":
                        data.append((group_original, beg_support, end_support))

                        if extract_first_numeric(change_mean) != 0:
                            line1 = f"Among {group} ({sample_sizes[index]}), the mean {return_text(change_mean)} by {add_p_in_parenthesis(change_mean)}."
                        else:
                            line1 = f"Among {group}, the mean did not change."

                        if stance(end_support, end_opposition, stance_negative, stance_positive) == stance(beg_support, beg_opposition, stance_negative, stance_positive):
                            line2 = f"After deliberation, {stance_percentage(end_support, end_opposition, stance_negative, stance_positive)} among this group, similar to before deliberation."
                        else:
                            line2 = f"Before deliberation, {stance_percentage(beg_support, beg_opposition, stance_negative, stance_positive)} among this group, while after deliberation, {stance_percentage(end_support, end_opposition, stance_negative, stance_positive)}."

                        lines.append(f"{line1}. {line2}")

                if API_Key != "None":
                    def ask_chatgpt(prompt):
                        headers = {
                            "Authorization": f"Bearer {API_Key}",
                            "Content-Type": "application/json"
                        }
                        data = {
                            "model": "gpt-3.5-turbo",
                            "messages": [
                                {"role": "user", "content": prompt}
                            ]
                        }
                        response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=data)
                        response_json = response.json()
                        return response_json["choices"][0]["message"]["content"].strip()

                    lines = ask_chatgpt(f"Rephrase the following to sound less structured and more human, but do not change the numbers or conclusions. Keep the numbers the same. Don't remove the word deliberation: {lines}")

                if not lines:
                    lines = "API error. Ensure there is a valid API key for ChatGPT."

                if data:
                    df = pd.DataFrame(data, columns=["Group", f"Selected {stance_positive} Before Deliberation", f"Selected {stance_positive} After Deliberation"])
                    if Demographic_Category != "Overall":
                        df.set_index("Group", inplace=True)
                    else:
                        df = df.transpose()
                        df.columns = df.iloc[0]
                        df = df.iloc[1:]

                    # Create a table in the document
                    table = document.add_table(rows=df.shape[0]+1, cols=df.shape[1])
                    table.autofit = False

                    # Set table style
                    table.style = "Table Grid"

                    # Set column widths
                    widths = [1, 2.5, 2.5]
                    for col, width in zip(table.columns, widths):
                        col.width = Pt(width * 72)

                    # Add table headers
                    headers = df.columns.tolist()
                    for i, header in enumerate(headers):
                        cell = table.cell(0, i)
                        cell.text = header
                        cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                    # Add table data
                    for row, row_data in enumerate(dataframe_to_rows(df, index=True, header=False)):
                        for col, cell_value in enumerate(row_data):
                            cell = table.cell(row+1, col)
                            cell.text = str(cell_value)
                            cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                    # Set cell alignment
                    for row in table.rows:
                        for cell in row.cells:
                            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                question_paragraph = document.add_paragraph(question)
                question_paragraph.style = "Heading 1"

                for line in lines:
                    document.add_paragraph(line)

                document.add_paragraph()
                document.add_paragraph()

            file_name = File_Name.replace("Tables", "Report") + ".docx"
            file_path = Outputs + "/" + file_name
            document.save(file_path)
            print(f"Exported: {file_name}")

# Example usage:
crosstabs =  # Provide the Crosstabs data as a list or a pandas DataFrame
outputs =  # Provide the path to the output directory
file_name =  # Provide the desired file name
name_group =  # Provide the Name_Group parameter value
template =  # Provide the path to the template file
document_title =  # Provide the Document_Title parameter value
type =  # Provide the Type parameter value
demographic_category =  # Provide the Demographic_Category parameter value
api_key =  # Provide the API Key for ChatGPT
group1 =  # Provide the Group1 parameter value

export_report(crosstabs, outputs, file_name, name_group, template, document_title, type, demographic_category, api_key, group1)
import pandas as pd
import numpy as np


def ordinal(Dataset, Template, Outputs, Group1, Group2, Alpha, Only_Significant, Only_Means, API_Key):


# Gets the codebook, datasets, and names in a list.
List_Datasets_Names = Dataset_Import(Dataset, Group1, Group2)

# Extracts the list elements.
Codebook = List_Datasets_Names[0]
Dataset_Group1 = List_Datasets_Names[1]
Dataset_Group2 = List_Datasets_Names[2]
Name_Group1 = List_Datasets_Names[3]
Name_Group2 = List_Datasets_Names[4]
Name_Group = List_Datasets_Names[5]
Group1 = List_Datasets_Names[6]
Group2 = List_Datasets_Names[7]
Time1 = List_Datasets_Names[8]
Time2 = List_Datasets_Names[9]
Weight1 = List_Datasets_Names[10]
Weight2 = List_Datasets_Names[11]

# Order datasets based on "Identification Number" column
Dataset_Group1 = Dataset_Group1.sort_values(by='Identification Number')
Dataset_Group2 = Dataset_Group2.sort_values(by='Identification Number')

# Adds overall column to codebook
Overall = pd.DataFrame(columns=["Overall"])
Overall.loc[0] = "Nominal"
Codebook = pd.concat([Codebook, Overall], axis=1)

# Organizes the codebook by column value type (opinions, demographics, etc.)
Codebook = Codebook.reindex(columns=Codebook.loc[0].argsort())

# Adds overall column to group data
Dataset_Group1["Overall"] = None
Dataset_Group2["Overall"] = None

# Gets the column number of the first and last demographic and questions.
Demographics = Codebook.columns[Codebook.iloc[0] == "Nominal"]
Questions = Codebook.columns[Codebook.iloc[0] == "Ordinal"]
Demographic_Beg = Demographics[0]
Demographic_End = Demographics[-1]
Question_Beg = Questions[0]
Question_End = Questions[-1]

# Values to stop crosstab creation once all questions and demographics have been analyzed
DemographicsCounter = Demographic_Beg - 1
DemographicsCounter_End = Demographic_End

# Values to stop crosstab creation once all questions have been analyzed
QuestionsCounter_End = Question_End + 1

while True:
# Repeats crosstab creation if final demographic category has not been completed.
    if DemographicsCounter == DemographicsCounter_End:
        break

# Creates a dataset to contain crosstabs.
Crosstabs = pd.DataFrame()

# Counts repetitions.
DemographicsCounter = DemographicsCounter + 1

# Resets to the first question.
QuestionsCounter = Question_Beg - 1

# Repeats crosstab creation for each question.
while True:
# Counts repetitions.
    QuestionsCounter = QuestionsCounter + 1

# Ends crosstab creation if all questions have been analyzed.
if QuestionsCounter == QuestionsCounter_End:
# Ends repetitions.
    break

# Gets the column number of Group 1.
ColumnNumber_Responses_Group1 = Dataset_Group1.columns.get_loc(Codebook[QuestionsCounter])

# Gets the column number of the demographics for Group 1.
ColumnNumber_Demographics_Group1 = Dataset_Group1.columns.get_loc(Codebook[DemographicsCounter])

# Gets the response data from Group 1.
Responses_Group1 = Dataset_Group1.iloc[:, ColumnNumber_Responses_Group1]

# Gets the demographic data for Group 1.
Demographics_Group1 = Dataset_Group1.iloc[:, ColumnNumber_Demographics_Group1]

# Gets numerical weights.
if Weight1 == "Unweighted":
    Weight_Group1 = Dataset_Group1["Overall"].apply(lambda x: 1)
else:
# Check to see if the weights exist
    Input_Test(Dataset_Group1, Weight1)

# Gets the weight
    Weight_Group1 = Dataset_Group1[Weight1]

if Weight2 == "Unweighted":
    Weight_Group2 = Dataset_Group2["Overall"].apply(lambda x: 1)
else:
# Check to see if the weights exist
    Input_Test(Dataset_Group2, Weight2)

# Gets the weight
    Weight_Group2 = Dataset_Group2[Weight2]

# Gets the column number of Group 2.
ColumnNumber_Responses_Group2 = Dataset_Group2.columns.get_loc(Codebook[QuestionsCounter])

# Gets the column number of the demographics for Group 2.
ColumnNumber_Demographics_Group2 = Dataset_Group2.columns.get_loc(Codebook[DemographicsCounter])

# Gets the response data from Group 2.
Responses_Group2 = Dataset_Group2.iloc[:, ColumnNumber_Responses_Group2]

# Gets the demographic data for Group 2.
Demographics_Group2 = Dataset_Group2.iloc[:, ColumnNumber_Demographics_Group2]

# Determines if the ID numbers are the same and, if they are, the data is set to paired.
Paired = Demographics_Group1.equals(Demographics_Group2)

if Paired:
    Demographics_Group2 = Demographics_Group1

# Gets the name of the demographic category
Demographic_Category = Demographics_Group1.name

# Gets the levels for the given question.
Levels_Responses = Codebook.loc[4:6, Codebook.columns == Responses_Group1.name].values.flatten()

# Gets the levels for the demographic category.
Levels_Demographics = Codebook.loc[:, Codebook.columns == Demographic_Category].dropna().values.flatten()
Levels_Demographics = list(range(1, len(Levels_Demographics)))

if Levels_Demographics == [1, 0]:
    Levels_Demographics = None

# Gets the question number for the corresponding responses.
QuestionNumber_Group1 = Responses_Group1.name
QuestionNumber_Group2 = Responses_Group2.name

# Converts responses to text.
Responses_Group1 = Responses_to_Text(Responses_Group1, QuestionNumber_Group1, Codebook)
Responses_Group2 = Responses_to_Text(Responses_Group2, QuestionNumber_Group2, Codebook)

# Generate crosstabs
Crosstab_Group1 = pd.crosstab(pd.Categorical(Responses_Group1, categories=Levels_Responses, ordered=True),
                              pd.Categorical(Demographics_Group1, categories=Levels_Demographics, ordered=True),
                              values=Weight_Group1, aggfunc='sum', normalize='index') * 100

Crosstab_Group2 = pd.crosstab(pd.Categorical(Responses_Group2, categories=Levels_Responses, ordered=True),
                              pd.Categorical(Demographics_Group2, categories=Levels_Demographics, ordered=True),
                              values=Weight_Group2, aggfunc='sum', normalize='index') * 100

# If all responses are NA, then the crosstab is blanked
if Crosstab_Group1.iloc[-1].isna().all():
    Crosstab_Group1.iloc[-1] = None
if Crosstab_Group2.iloc[-1].isna().all():
    Crosstab_Group2.iloc[-1] = None

# Gets the crosstab of the differences
Crosstab_Difference = Crosstab_Group2 - Crosstab_Group1

# Formats crosstabs.
Crosstab_Group1 = Crosstab_Group1.round(1).applymap(lambda x: f"{x}%" if not pd.isna(x) else x)
Crosstab_Group2 = Crosstab_Group2.round(1).applymap(lambda x: f"{x}%" if not pd.isna(x) else x)
Crosstab_Difference = Crosstab_Difference.round(1).applymap(lambda x: f"{x}%" if not pd.isna(x) else x)

# Adds total column to crosstabs.
Totals_Group1 = pd.crosstab(pd.Categorical(Responses_Group1, categories=Levels_Responses, ordered=True),
                            pd.Categorical(['T'] * len(Responses_Group1)), values=Weight_Group1, aggfunc='sum',
                            normalize='index') * 100

Totals_Group2 = pd.crosstab(pd.Categorical(Responses_Group2, categories=Levels_Responses, ordered=True),
                            pd.Categorical(['T'] * len(Responses_Group2)), values=Weight_Group2, aggfunc='sum',
                            normalize='index') * 100

# If all responses are NA, then the crosstab is blanked
if pd.isna(Totals_Group1.iloc[-1].values[0]):
    Totals_Group1.iloc[-1] = None
if pd.isna(Totals_Group2.iloc[-1].values[0]):
    Totals_Group2.iloc[-1] = None

# Gets the crosstab of the differences
Totals_Difference = Totals_Group2 - Totals_Group1

# Formats totals.
Totals_Group1 = pd.DataFrame(np.round(Totals_Group1, decimals=1)).applymap("{:.1f}".format)
Totals_Group2 = pd.DataFrame(np.round(Totals_Group2, decimals=1)).applymap("{:.1f}".format)
Totals_Difference = pd.DataFrame(np.round(Totals_Difference, decimals=1)).applymap("{:.1f}".format)

# Adds percentage symbol "%" to total column entries.
Totals_Group1 = Totals_Group1.applymap("{}%".format)
Totals_Group2 = Totals_Group2.applymap("{}%".format)
Totals_Difference = Totals_Difference.applymap("{}%".format)

# Adds total column to crosstabs.
Crosstab_Group1 = pd.concat([Totals_Group1, Crosstab_Group1], axis=1)
Crosstab_Group2 = pd.concat([Totals_Group2, Crosstab_Group2], axis=1)
Crosstab_Difference = pd.concat([Totals_Difference, Crosstab_Difference], axis=1)

# Replaces "." with spaces.
Crosstab_Difference.index = Crosstab_Difference.index.str.replace(".", " ")
Crosstab_Difference.columns = Crosstab_Difference.columns.str.replace(".", " ")

# Ensures the number of rows of the crosstabs is 4.
NumberofRowsis4 = True
if Crosstab_Group1.shape[0] == 3:
    Spacer = pd.DataFrame(np.full((1, Crosstab_Group1.shape[1]), 9999))
    Spacer.index = [9999]
    Spacer.columns = Crosstab_Group1.columns
    Crosstab_Group1 = pd.concat([Crosstab_Group1, Spacer])
    Crosstab_Group2 = pd.concat([Crosstab_Group2, Spacer])
    Crosstab_Difference = pd.concat([Crosstab_Difference, Spacer])
    NumberofRowsis4 = False

# Adds a spacer between the crosstabs for both groups.
Spacer = pd.Series([9999, 9999, 9999, 9999])
Crosstab_Proportions = pd.concat([Crosstab_Group1, Spacer, Crosstab_Group2, Spacer, Crosstab_Difference], axis=1)

# Creates spacers to fill in with means and no opinion data.
Spacer = pd.DataFrame([[9999]], columns=[None])
Means_Group1 = Means_Group2 = Means_Difference = NoOpinions_Group1 = NoOpinions_Group2 = NoOpinions_Difference = pd.DataFrame(
    np.full((1, Crosstab_Group1.shape[1]), 9999),
    columns=Crosstab_Group1.columns
)

# Gets the numeric response data.
Responses_Group1 = Dataset_Group1[QuestionNumber_Group1]
Responses_Group2 = Dataset_Group2[QuestionNumber_Group2]

# Gets the weighted NA raw percentage.
Responses_Group_NA1 = Responses_Group1.copy()
Responses_Group_NA2 = Responses_Group2.copy()

# TRUE if responses are all NAs, indicating no responses for the given prompt.
AllNAs_1 = Responses_Group_NA1.isna().all().all()
AllNAs_2 = Responses_Group_NA2.isna().all().all()

# Gets the percentage of responses that were NA.
if AllNAs_1:
    NoOpinion_Group1 = "NaN"
elif not Responses_Group_NA1.isna().any().any():
    NoOpinion_Group1 = 0
else:
    Responses_Group_NA1[Responses_Group_NA1.isna()] = -99
    if Responses_Group_NA1.size == 1:
        NoOpinion_Group1 = 100
    else:
        svytable_data = pd.DataFrame({'Responses': Responses_Group_NA1.unstack(), 'Weights': Weight_Group1.unstack()})
        NoOpinion_Group1 = 100 * \
                           (svytable_data.groupby('Responses')['Weights'].sum() / svytable_data['Weights'].sum()).loc[
                               -99]

if AllNAs_2:
    NoOpinion_Group2 = "NaN"
elif not Responses_Group_NA2.isna().any().any():
    NoOpinion_Group2 = 0
else:
    Responses_Group_NA2[Responses_Group_NA2.isna()] = -99
    if Responses_Group_NA2.size == 1:
        NoOpinion_Group2 = 100
    else:
        svytable_data = pd.DataFrame({'Responses': Responses_Group_NA2.unstack(), 'Weights': Weight_Group2.unstack()})
        NoOpinion_Group2 = 100 * \
                           (svytable_data.groupby('Responses')['Weights'].sum() / svytable_data['Weights'].sum()).loc[
                               -99]

if AllNAs_1 or AllNAs_2:
    NoOpinion_Difference = "NaN"
else:
    NoOpinion_Difference = NoOpinion_Group2 - NoOpinion_Group1
    NoOpinion_Difference = "{:.1f}".format(round(NoOpinion_Difference, 1))

if not AllNAs_1:
    NoOpinion_Group1 = "{:.1f}".format(round(NoOpinion_Group1, 1))

if not AllNAs_2:
    NoOpinion_Group2 = "{:.1f}".format(round(NoOpinion_Group2, 1))

# If data is paired, removes NAs.
if Paired and not AllNAs_1 and not AllNAs_2:
# Removes entries for which there are missing responses in either groups.
    Responses_Group1_2 = pd.concat([Responses_Group1, Responses_Group2, Weight_Group1, Weight_Group2], axis=1)
    Responses_Group1_2.dropna(inplace=True)

# Inserts formatted data.
Responses_Group1 = Responses_Group1_2.iloc[:, 0]
Responses_Group2 = Responses_Group1_2.iloc[:, 1]
Weight_Group1 = Responses_Group1_2.iloc[:, 2]
Weight_Group2 = Responses_Group1_2.iloc[:, 3]

# Unlists response data.
Responses_Group1 = Responses_Group1.values.flatten()
Responses_Group2 = Responses_Group2.values.flatten()

# Gets the means of the response data.
Mean_Group1 = round(np.average(Responses_Group1, weights=Weight_Group1, nan_policy='omit'), 3)
Mean_Group2 = round(np.average(Responses_Group2, weights=Weight_Group2, nan_policy='omit'), 3)
Mean_Difference = round(Mean_Group2 - Mean_Group1, 3)

# Performs a t-test if possible and adds data to crosstabs.
# (Assuming Test_T is a custom function for performing t-test)
Mean_Difference = Test_T(Mean_Difference, Responses_Group1, Responses_Group2, Weight_Group1, Weight_Group2, Paired)

# Adds the no opinion data to the dataframe.
NoOpinions_Group1[0] = str(NoOpinion_Group1)
NoOpinions_Group2[0] = str(NoOpinion_Group2)
NoOpinions_Difference[0] = str(NoOpinion_Difference)

# Adds percentage symbol "%" to no opinion entries.
NoOpinions_Group1[0] = NoOpinions_Group1[0] + "%"
NoOpinions_Group2[0] = NoOpinions_Group2[0] + "%"
NoOpinions_Difference[0] = NoOpinions_Difference[0] + "%"

# Adds the mean data to the dataframe.
Means_Group1[0] = str(Mean_Group1)
Means_Group2[0] = str(Mean_Group2)
Means_Difference[0] = str(Mean_Difference)

# Repeats creation of means and no opinion data for each demographic category.
CategoryCounter_End = len(Crosstab_Group1.columns)
CategoryCounter = 0
while CategoryCounter < CategoryCounter_End:
    CategoryCounter += 1
    if CategoryCounter > CategoryCounter_End:
        break
    count1001 = CategoryCounter

# Extracts demographic category.
    Category = Crosstab_Group1.columns[count1001]
    if Category == "T":
        Category = Crosstab_Group1.columns[count1001 + 1]

# Gets responses within each demographic category.
Selections_Group1 = Dataset_Group1[Dataset_Group1[Demographics_Group1].eq(Category).any(axis=1)]
Selections_Group2 = Dataset_Group2[Dataset_Group2[Demographics_Group2].eq(Category).any(axis=1)]
Responses_Group1 = Selections_Group1[QuestionNumber_Group1].values
Responses_Group2 = Selections_Group2[QuestionNumber_Group2].values

# Gets numerical weights.
if len(Responses_Group1) > 0:
    if Weight1 == "Unweighted":
        Weight_Group1 = Selections_Group1["Overall"]
        Weight_Group1 = 1
    elif Weight1 != "Unweighted":
        Weight_Group1 = Selections_Group1[Weight1].values
    else:
        Weight_Group1 = Responses_Group1

if len(Responses_Group2) > 0:
    if Weight2 == "Unweighted":
        Weight_Group2 = Selections_Group2["Overall"]
        Weight_Group2 = 1
    elif Weight2 != "Unweighted":
        Weight_Group2 = Selections_Group2[Weight2].values
    else:
        Weight_Group2 = Responses_Group2

# TRUE if responses are all NAs (indicating respondents weren't surveyed).
AllNAs_1 = np.all(pd.isna(Responses_Group1))
AllNAs_2 = np.all(pd.isna(Responses_Group2))

# Gets the weighted NA raw percentage.
Responses_Group_NA1 = Responses_Group1.copy()
Responses_Group_NA2 = Responses_Group2.copy()

# TRUE if responses are all NAs, indicating no responses for the given prompt.
AllNAs_1 = np.all(pd.isna(Responses_Group_NA1))
AllNAs_2 = np.all(pd.isna(Responses_Group_NA2))

# Gets the percentage of responses that were NA.
# Gets the percentage of responses that were NA for group 1.
if AllNAs_1:
    NoOpinion_Group1 = "NaN"
elif not np.any(pd.isna(Responses_Group_NA1)):
    NoOpinion_Group1 = 0
else:
    Responses_Group_NA1[pd.isna(Responses_Group_NA1)] = -99
if len(Responses_Group_NA1) == 1:
    NoOpinion_Group1 = 100
else:
    svytable_data = pd.DataFrame({'Responses': Responses_Group_NA1, 'Weights': Weight_Group1})
NoOpinion_Group1 = 100 * (svytable_data.groupby('Responses')['Weights'].sum() / svytable_data['Weights'].sum()).loc[-99]

# Gets the percentage of responses that were NA for group 2.
if AllNAs_2:
    NoOpinion_Group2 = "NaN"
elif not np.any(pd.isna(Responses_Group_NA2)):
    NoOpinion_Group2 = 0
else:
Responses_Group_NA2[pd.isna(Responses_Group_NA2)] = -99
if len(Responses_Group_NA2) == 1:
    NoOpinion_Group2 = 100
else:
    svytable_data = pd.DataFrame({'Responses': Responses_Group_NA2, 'Weights': Weight_Group2})
NoOpinion_Group2 = 100 * (svytable_data.groupby('Responses')['Weights'].sum() / svytable_data['Weights'].sum()).loc[-99]

# Gets the percentage of responses that were NA for the difference between group 1 and 2.
if AllNAs_1 or AllNAs_2:
    NoOpinion_Difference = "NaN"
else:
    NoOpinion_Difference = NoOpinion_Group2 - NoOpinion_Group1
NoOpinion_Difference = "{:.1f}".format(round(NoOpinion_Difference, 1))


# Define a function to format the mean and no opinion values
def format_values(value):
    return "{:.3f}".format(round(value, 3))


# If data is paired, remove NAs.
if Paired and not AllNAs_1 and not AllNAs_2:
# Removes entries for which there are missing responses in either groups.
Responses_Group1_2 = pd.concat([Responses_Group1, Responses_Group2, Weight_Group1, Weight_Group2], axis=1)
Responses_Group1_2 = Responses_Group1_2.dropna()

# Inserts formatted data.
Responses_Group1 = Responses_Group1_2.iloc[:, 0]
Responses_Group2 = Responses_Group1_2.iloc[:, 1]
Weight_Group1 = Responses_Group1_2.iloc[:, 2]
Weight_Group2 = Responses_Group1_2.iloc[:, 3]

# Unlists response data.
Responses_Group1 = Responses_Group1.values.flatten()
Responses_Group2 = Responses_Group2.values.flatten()

# Gets the means of the response data.
Mean_Group1 = format_values(np.average(Responses_Group1, weights=Weight_Group1, axis=0, nan_policy='omit'))
Mean_Group2 = format_values(np.average(Responses_Group2, weights=Weight_Group2, axis=0, nan_policy='omit'))
Mean_Difference = format_values(np.average(Responses_Group2, weights=Weight_Group2, axis=0, nan_policy='omit') -
                                np.average(Responses_Group1, weights=Weight_Group1, axis=0, nan_policy='omit'))

# Performs a t-test if possible and adds data to crosstabs.
Mean_Difference = Test_T(Mean_Difference, Responses_Group1, Responses_Group2, Weight_Group1, Weight_Group2, Paired)

# Adds the no opinion data to dataframe.
NoOpinions_Group1[count1001] = str(NoOpinion_Group1)
NoOpinions_Group2[count1001] = str(NoOpinion_Group2)
NoOpinions_Difference[count1001] = str(NoOpinion_Difference)

# Adds percentage symbol "%" to no opinion entries.
NoOpinionPercentageAdder = pd.DataFrame(np.full((1, 1), ""), columns=["Percentage"])
NoOpinions_Group1[count1001] = NoOpinions_Group1[count1001] + "%"
NoOpinions_Group2[count1001] = NoOpinions_Group2[count1001] + "%"
NoOpinions_Difference[count1001] = NoOpinions_Difference[count1001] + "%"

# Adds the mean data to the dataframe.
Means_Group1[count1001] = str(Mean_Group1)
Means_Group2[count1001] = str(Mean_Group2)
Means_Difference[count1001] = str(Mean_Difference)

# Combine means and no opinion data.
Crosstab_Means = pd.concat([Means_Group1, Spacer, Means_Group2, Spacer, Means_Difference], axis=1)
Crosstab_NoOpinions = pd.concat([NoOpinions_Group1, Spacer, NoOpinions_Group2, Spacer, NoOpinions_Difference], axis=1)

# Combine crosstab with means and no opinion.
Crosstab_Means.columns = Crosstab_NoOpinions.columns = Crosstab_Means.columns
Crosstab = pd.concat([Crosstab_Means, Crosstab_Proportions])

# Add header to crosstab.
Crosstab = pd.concat([RowHeaders, Crosstab], axis=1)

# Add spacer for between different crosstabs.
Spacer = pd.DataFrame(np.full((1, Crosstab.shape[1]), 9999), columns=Crosstab.columns)

# Remove unnecessary row.
if not NumberofRowsis4:
    Crosstab.iloc[3, 1] = Crosstab.iloc[4, 1]
Crosstab = Crosstab.drop(4)

# Remove non-means and add mean statistics if specified.
if Only_Means:
    Crosstab = Crosstab.drop([1, 2], axis=0)
Spacer = pd.DataFrame()

# Combine crosstabs to one dataframe.
Crosstabs = pd.concat([Crosstabs, Spacer, Crosstab])


# Function to calculate sample sizes
def calculate_sample_sizes(responses):
    return len(responses)


def replicate_code_in_python():
    for break_indicator in Break_Indicator:
        if break_indicator:
        break


# Adds crosstab headers to the dataframe
ColumnHeaders = list(Crosstabs.columns)
Crosstabs = pd.concat([pd.DataFrame([ColumnHeaders]), Crosstabs], ignore_index=True)

# Adds sample size statistics
SampleSizes_Group1 = pd.DataFrame([[9999] * (len(Crosstab_Group1) - 1)], columns=Crosstab_Group1[:-1])
SampleSizes_Group2 = pd.DataFrame([[9999] * (len(Crosstab_Group2) - 1)], columns=Crosstab_Group2[:-1])

CategoryCounter_End = len(Crosstab_Group1)
for count1001 in range(1, CategoryCounter_End):
    Category = Crosstab_Group1[count1001]
if Category == "T":
    Category = Crosstab_Group1[count1001 + 1]

Selections_Group1 = Dataset_Group1[Dataset_Group1[Demographics_Group1].str.match("^" + Category + "$")]
Selections_Group2 = Dataset_Group2[Dataset_Group2[Demographics_Group2].str.match("^" + Category + "$")]
Responses_Group1 = Selections_Group1[QuestionNumber_Group1]
Responses_Group2 = Selections_Group2[QuestionNumber_Group2]

SampleSizeNotation = "n = "
SampleSize_Group1 = calculate_sample_sizes(Responses_Group1)
SampleSize_Group2 = calculate_sample_sizes(Responses_Group2)

SampleSizes_Group1[count1001 - 1] = SampleSizeNotation + str(SampleSize_Group1)
SampleSizes_Group2[count1001 - 1] = SampleSizeNotation + str(SampleSize_Group2)

SampleSizeNotation = "n = "
SampleSize_Group1 = calculate_sample_sizes(Dataset_Group1)
SampleSize_Group2 = calculate_sample_sizes(Dataset_Group2)

TotalSize_Group1 = SampleSizeNotation + str(SampleSize_Group1)
TotalSize_Group2 = SampleSizeNotation + str(SampleSize_Group2)

Spacer_1 = pd.DataFrame([[9999, 9999]], columns=["Spacer", "Spacer"])
Spacer_2 = pd.DataFrame([[9999]], columns=["Spacer"])
SampleSizes = pd.concat(
    [Spacer_1, TotalSize_Group1, SampleSizes_Group1, Spacer_2, TotalSize_Group2, SampleSizes_Group2], axis=1)
Spacer_3 = pd.DataFrame([[9999] * (len(Crosstabs.columns) - len(SampleSizes.columns))])
SampleSizes = pd.concat([SampleSizes, Spacer_3], ignore_index=True)
Spacer = pd.DataFrame([[9999] * len(Crosstabs.columns)])
SampleSizes.columns = Crosstabs.columns
Spacer.columns = Crosstabs.columns
Crosstabs = pd.concat([SampleSizes, Spacer, Crosstabs], ignore_index=True)

# Creates a header dataframe
ColumnHeaders = pd.DataFrame([[9999] * len(Crosstabs.columns)], columns=Crosstabs.columns)

ColumnHeaders.at[0, Crosstab_Group1[2]] = Name_Group1
ColumnHeaders.at[0, Crosstab_Group1[3] + 1 * len(Crosstab_Group1)] = Name_Group2
ColumnHeaders.at[0, Crosstab_Group1[4] + 2 * len(Crosstab_Group1)] = "Difference"

Crosstabs = pd.concat([ColumnHeaders, Crosstabs], ignore_index=True)

# Creates the name and title of the files
if Weight1 == Weight2:
    Weights = Weight1
else:
    Weights = Weight1 + " and " + Weight2

File_Name = "Tables_Ordinal_" + Name_Group + "_" + Weights + "_" + Demographic_Category
Document_Title = (Name_Group + ", Weighted by " + Weights + " " + Demographic_Category).replace(" by Overall",
                                                                                                "").replace(
    "Weighted by Unweighted", "Unweighted")

# Removes markers
Crosstabs = Crosstabs.applymap(
    lambda x: str(x).replace("matrix.data.....nrow...1..ncol...1.", "").replace("NaN%", "").replace("NaN", "").replace(
        "NA%", "").replace("9999", "").replace("Spacer", "").replace("In.the.middle", "In the middle"))

# Runs a pre-defined function that creates Word documents from the crosstabs
Export_Word(Crosstabs, Outputs, File_Name, Name_Group, Template, Document_Title, Type="Ordinal", Demographic_Category,
            Codebook)


