This package is for analyzing survey data from Deliberative Polling experiments. Although designed for Deliberative Polling, this package can be used to analyze any experimental survey data.

The package is designed with a single, specialized function called `outputs`. This function accepts as input files exclusively in the IBM SPSS Statistics `.SAV` format. Upon execution, it generates output files in both `.xlsx` and `.docx` formats. These output files contain statistical comparisons of all ordinal and nominal variables across all designated treatment groups, time intervals, and statistical weights.

# Installation

To install SPSS, go to [Software at Stanford](https://software.stanford.edu) if you are a Stanford affiliate. Othwerwise, go to [IBM SPSS Software](https://www.ibm.com/spss).

To install Python, go to [Download Python](https://www.python.org/downloads/).

To install DeliberativePolling, run the following in a terminal:

```{bash}
pip install DeliberativePolling
```

# In SPSS

To import data into SPSS, open SPSS and navigate to `File` and `Import Data`.

<div style="text-align: center;">
  <img src="https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Images/Screenshot%202023-09-26%20at%2010.36.26%20PM.png?raw=true" alt="Alt text" width="400"/>
</div>

## Essential Variables

In order for `outputs` to identify the different subjects, experimental groups, and time intervals in the data, the SPSS file must contain three variables: `ID`, `Time`, and `Group`.

<div style="text-align: center;">
  <img src="https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Images/Screenshot%202023-09-26%20at%2010.47.46%20PM.png?raw=true" alt="Alt text" width="400"/>
</div>

### ID

The `ID` variable helps track individual participants in the study. It's like a name tag that stays the same for each person throughout the experiment. This way, you can see how a person's answers change over time. The ID can be a number, an email address, or any other unique identifier.

### Group

The `Group` variable tells you which part of the experiment a participant is in—either the `Treatment` group that receives the intervention, or the `Control` group that doesn't. This helps you compare the effects of the treatment.

### Time

The `Time` variable shows when a participant gave their answers. Labels like `Pre-Deliberation` or `T1` are usually used for answers given before the treatment, and `Post-Deliberation` or `T2` for answers given after. This helps you see how responses change over the course of the experiment.

## Optional Variables

### Weights

By default, the `outputs` function generates unweighted tables that compare survey data between all experimental groups and time intervals; however, you can introduce weighting by including columns with the word `weight` in the header, like `Weight1` in [Sample.SAV](https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Sample.sav). These weight variables must be numeric with their `Measure` set to `Scale`.

<div style="text-align: center;">
  <img src="https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Images/Screenshot%202023-09-26%20at%2010.49.35%20PM.png?raw=true" alt="Alt text" width="400"/>
</div>

### Ignored

To keep variables in the SPSS file that you don't want included in the `outputs` function's analysis but might use later, set their `Measure` to `Scale`; variables with this setting won't be part of the analysis unless they are designated as weight variables.

## Remove Duplicate Variables

When you first import your dataset into SPSS, you might notice that a single question (e.g., "How well is democracy functioning?") is represented by multiple variables each corresponding to a different time point. For instance, you could have one variable named `Question1` for responses collected at the first time point (e.g., `T1`, `Pre-Deliberation`) and another variable named `T2Question1` for responses collected at the second time point (e.g., `T2`, `Post-Deliberation`). These variables need to be consolidated into a single variable.

To achieve this, you'll need to create additional rows in your dataset for each participant, capturing their responses to the same question at different time intervals.

- **First Row**: This row will contain the participant's response to the first time-specific variable (e.g., `Question1`). In the `Time` column, you should enter the label that corresponds to this time interval (e.g., `T1`, `Pre-Deliberation`).

- **Second Row**: This row will contain the participant's response to the second time-specific variable (e.g., `T2Question1`). In the `Time` column, you should enter the label that corresponds to this time interval (e.g., `T2`, `Post-Deliberation`).

By following this approach, you will stack all the responses under a single variable (e.g., `Question1`). This allows each participant's responses to be represented multiple times in the dataset, each corresponding to a different time interval. While these new rows will, of course, have different values for `Time`, their values for `ID` and `Group` cannot change.

<div style="text-align: center;">
  <img src="https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Images/Screenshot%202023-09-27%20at%205.15.46%20AM.png?raw=true" alt="Alt text" width="700"/>
</div>

## Measures

In the `Measure` column of `Variable View`, variables can be classified as `Nominal`, `Ordinal`, or `Scale`.

<div style="text-align: center;">
  <img src="https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Images/Screenshot%202023-09-26%20at%2010.41.48%20PM.png?raw=true" alt="Alt text" width="400"/>
</div>

### Nominal

Nominal variables are categorical variables that lack a sequential order. For instance, the variable `Employment` in the [Sample.SAV](https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Sample.sav) file includes the categories `Employed`, `Unemployed`, `Student`, and `Other`, which don't follow a specific sequence. While there are exceptions, such as `Education Level`, which do have an order, it's generally advisable (but not mandatory) to categorize variables containing demographic data as `Nominal`.

### Ordinal

Ordinal variables are categorical variables that have a well-defined order. For example, the variable `Question1` in the [Sample.SAV](https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Sample.sav) file. This variable uses a Likert scale that ranges from 0 to 10, representing a progression from `Poorly` to `Well` in response to the question "How well does democracy function?" Typically, it's recommended (but not obligatory) to classify variables with responses that change between time intervals as `Ordinal`.

*Some statisticians indicate non-response to survey questions using high numeric codes like `77`, `98`, or `99`. It's crucial to remove these high numeric codes from ordinal variables before analysis. The `outputs` function calculates the average of survey responses in ordinal variables, assuming a consistent scale like 0-10, 1-5, or 1-3. Including out-of-scale high values like `99` can significantly distort the calculated mean. To avoid this, replace these numeric codes with blank cells; blank cells will be counted as `DK/NA` (Don't Know/Not Applicable) and will not affect mean calculations. The exception to this rule would, of course, be if the scale naturally includes `77`, `98`, or `99`, such as 0-100.*

### Scale

Any variables that don't fit into the `Nominal` or `Ordinal` categories should be classified as `Scale` variables. These can either be continuous or discrete. All variables related to weight should be categorized as `Scale`.

## Labels

In SPSS, labels help clarify the meaning of variable names and values.

### Column Labels

Variable names can't have spaces or punctuation. Descriptive `Column Labels` can be set in `Variable View` under the column `Label` to provide more information about the variables.

<div style="text-align: center;">
  <img src="https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Images/Screenshot%202023-09-26%20at%2010.51.07%20PM.png?raw=true" alt="Alt text" width="400"/>
</div>

**Nominal Variables**: For nominal variables use concise labels. For example, the variable `Education` in [Sample.SAV](https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Sample.sav) has the column label `Education Level`. Keep these labels short because they will appear in file names like `Tables - Ordinal Variables - Treatment at T1 v. T2 (Unweighted) - Education Level`.

**Ordinal Variables**: For ordinal variables you can use fuller more descriptive labels. For example, the variable `Question1` in [Sample.SAV](https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Sample.sav) has the column label `How well does democracy function?`. These ordinal column labels do not appear in file names, only within cells in the outputted files so length is unlikely to cause an issue.

### Value Labels

<div style="text-align: center;">
  <img src="https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Images/Screenshot%202023-09-26%20at%2010.51.54%20PM.png?raw=true" alt="Alt text" width="400"/>
</div>

When working with SPSS, it's essential to set the `Type` of both `Ordinal` and `Nominal` variables to `Numeric` in the `Variable View`. Since the data will be numeric, you'll use value labels to provide meaningful context to these coded numbers.

**Numeric Codes**: Value labels allow you to give meaning to numerically coded data. Use the `Values` column in `Variable View` to associate each number in ordinal or nominal variables with a label. For example, in the ordinal variable `Age` in [Sample.SAV](https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Sample.sav), the value labels indicate that a value of `1` means `18-30` and `2` means `30-50`.

<div style="text-align: center;">
  <img src="https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Images/Screenshot%202023-09-26%20at%2010.52.45%20PM.png?raw=true" alt="Alt text" width="400"/>
</div>

**Shared Labels**: Some variables might have several numeric codes that mean the same thing. For instance, in the ordinal variable `Question1`, the codes `0` through `4` are all labeled as `Poorly`, while `6` through `10` are labeled as `Well`.

<div style="text-align: center;">
  <img src="https://github.com/WilloughbyWinograd/DeliberativePolling/blob/main/Images/Screenshot%202023-09-26%20at%2010.53.10%20PM.png?raw=true" alt="Alt text" width="400"/>
</div>

*Ensure that all values have labels, otherwise the `outputs` function will return an error message indicating which values are unlabeled.*

Once you've included all essential variables and assigned column and value labels to all nominal and ordinal variables, you can run the `outputs` function on the SPSS file. If any metadata is missing, the `outputs` function will return an error and specify what data is lacking.

*Columns like `Width`, `Decimals`, `Missing`, `Columns`, `Align`, and `Role` in `Variable View` can usually be ignored.*

# In Python

To execute the `outputs` function, open a terminal with the directory containing the `.SAV` file. Then, run the following commands:

```{bash}
Python3
from DeliberativePolling import outputs
outputs("your_file.SAV")
```

## Outputs

After running the function, a new folder named `Outputs` will be created in the directory. This folder will contain all the generated tables and reports in `.xlsx` and `.docx` format.