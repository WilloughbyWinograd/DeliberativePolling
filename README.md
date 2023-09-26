This package is for analyzing survey data from Deliberative Polling experiments. Although designed for Deliberative Polling, this package can be used to analyze survey data from any experiment.

The package is designed with a single, specialized function called `outputs`. This function is engineered to accept input files exclusively in the IBM SPSS Statistics `.SAV` format. Upon execution, it generates output files in both `.xlsx` and `.docx` formats. These output files contain comprehensive comparisons of responses across all designated treatment groups, time intervals, and statistical weights.

# Installation

To install SPSS, go to [Software at Stanford](https://software.stanford.edu) if you are a Stanford affiliate. Othwerwise, go to [IBM SPSS Software](https://www.ibm.com/spss).

To install Python, go to [Download Python](https://www.python.org/downloads/).

To install DeliberativePolling, run the following in Terminal.

```{bash}
pip install DeliberativePolling
```

# In SPSS

To import data into SPSS, open SPSS and navigate to `File`, `Import Data` or simply copy and paste the data directly into the tab `Data View`.

Once the data has been imported into SPSS, you need to provide metadata about the variables in the tab `Variable View`.

## Measures

In the `Measure` column of `Variable View`, variables can be classified as `Nominal`, `Ordinal`, or `Scale`.

### Nominal

Nominal variables are categorical variables that lack a natural order. For instance, the variable `Employment` in the `Sample.SAV` file includes the categories `Employed`, `Unemployed`, `Student`, and `Other`, which don't follow a specific sequence. While there are exceptions, such as `Education Level`, which do have an order, it's generally advisable (but not mandatory) to categorize demographic data as `Nominal`.

### Ordinal

Ordinal variables are categorical variables that have a well-defined order. For example, the variable `Question1` in the `Sample.SAV` file. This variable uses a Likert scale that ranges from 0 to 10, representing a progression from `Poorly` to `Well` in response to the question "How well does democracy function?" Typically, it's recommended (but not obligatory) to classify variables with responses that change between time intervals as `Ordinal`.

    Note: Some statisticians indicate non-response to survey questions using high numeric codes like `77`, `98`, or `99`. It's crucial to remove these high numeric codes from ordinal variables before analysis. The `outputs` function calculates the average of ordinal values, assuming a consistent scale like 0-10, 1-5, or 1-3. Including out-of-scale high values like `99` can significantly distort the calculated mean. To avoid this, replace these numeric codes with blank cells; blank cells will be counted as `DK/NA` (Don't Know/Not Applicable) and will not affect mean calculations.

### Scale

Any variables that don't fit into the `Nominal` or `Ordinal` categories should be classified as `Scale` variables. These can either be continuous or discrete. For example, all variables related to weight should be categorized as `Scale`.

## Essential Variables
In order for `outputs` to identify the different subjects, experimental groups, and time intervals in the data, the SPSS file must contain three variables: `ID`, `Time`, and `Group`.

### ID
The `ID` variable helps track individual participants in the study. It's like a name tag that stays the same for each person throughout the experiment. This way, you can see how a person's answers change over time. The ID can be a number, an email address, or any other unique identifier.

### Group
The `Group` variable tells you which part of the experiment a participant is in—either the `Treatment` group that receives the intervention, or the `Control` group that doesn't. This helps you compare the effects of the treatment.

### Time
The `Time` variable shows when a participant gave their answers. Labels like `Pre-Deliberation` or `T1` are usually used for answers given before the treatment, and `Post-Deliberation` or `T2` for answers given after. This helps you see how responses change over the course of the experiment.

## Optional Variables

### Weights
By default, the `outputs` function generates tables that unweightedly compare survey data between all experimental groups and time intervals; however, you can introduce weighting by including columns with the word `weight` in the header, like `Weight1` in `Sample.SAV`. These weight variables must be numeric with their `Measure` set to `Scale`.

### Unanalyzed
To keep variables in the SPSS file that you don't want included in the `outputs` function's analysis but might use later, set their `Measure` to `Scale`; variables with this setting won't be part of the analysis unless they are designated as weight variables.

## Labels

In SPSS, labels help clarify the meaning of variables and their values.

### Column Labels
Variable names can't have spaces or punctuation. Descriptive `Column Labels` can be set in `Variable View` under the column `Label` to provide more information about the variables.

**Nominal Variables**: For nominal variables use concise labels. For example, the variable `Education` in `Sample.SAV` has the column label `Education Level`. Keep these labels short because they will appear in file names like `Tables - Ordinal Variables - Treatment at T1 v. T2 (Unweighted) - Education Level`.

**Ordinal Variables**: For ordinal variables you can use fuller more descriptive labels. For example, the variable `Question1` in `Sample.SAV` has the column label `How well does democracy function?`. These ordinal column labels do not appear in file names, only within cells in the outputted files so length is less of an issue.

### Value Labels

When working with SPSS, it's essential to set the `Type` of both `Ordinal` and `Nominal` variables to `Numeric` in the `Variable View`. Since the data will be numeric, you'll use value labels to provide meaningful context to these coded numbers.

**Numeric Codes**: For ordinal variables like `Age` in `Sample.SAV`, you'll need to specify what each numeric code (`1`, `2`, `3`, and `4`) represents. Use the `Values` column in `Variable View` to associate each number with a label, such as `1` for `18-30` and `2` for `30-50`.

**Shared Labels**: Some variables might have several numeric codes that mean the same thing. For instance, in the ordinal variable `Question1`, the codes `0` through `4` are all labeled as `Poorly`, while `6` through `10` are labeled as `Well`.

    Note: Ensure that all values have labels, otherwise the `outputs` function will return an error message indicating which values are unlabeled.

Once you've included all essential variables and assigned column and value labels to all nominal and ordinal variables, you can run the `outputs` function on the SPSS file. If any metadata is missing, the `outputs` function will return an error and specify what data is lacking.

    Note: Columns like `Width`, `Decimals`, `Missing`, `Columns`, `Align`, and `Role` in `Variable View` can usually be ignored.

# In Python

To execute the `outputs` function, open Terminal or your preferred IDE in the directory where the `.SAV` file is located. Then, run the following commands:

```{bash}
Python3
```

```{bash}
from DeliberativePolling import outputs
outputs("your_file.SAV")
```

To test execution in Python, you can run the following command:
```{bash}
outputs("Sample.SAV")
```

## Outputs

After running the function, a new folder named `Outputs` will be created in the directory. This folder will contain all the generated tables and reports in `.xlsx` format. If these tables and reports are reasonably sized (under 10,000 cells), they will also be exported in `.docx` format.