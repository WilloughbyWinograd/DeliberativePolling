# How To

Purpose: This guide is designed to assist professionals in efficiently leveraging a Python package tailored for the analysis of survey data in Deliberative Polling experiments.

Ensure your data is in an  `.SAV` file format from IBM SPSS Statistics.

Variables can be classified into three categories.

Nominal Variables: These are categorical variables that do not have an intrinsic order. For exmaple, Gender, where categories like male, female, and non-binary do not have a specific sequence.

Ordinal Variables: These are categorical variables with a clear, definable order. For example, data derived from a Likert scale ranging from 0 to 10. The values indicate a progression from least to most favorable (or vice versa).

Scale Variables: Variables not classified as either Nominal or Ordinal are listed under this category. These can be continuous or discrete variables. Essential variables such as Time, Group, and ID fall under this category.

To run the function, install this Python package to your device by running Python in the terminal or an IDE.

In a Python terminal, run:
"pip install DeliberativePolling"

\Then in a directory containing the .SAV file run:
"from DeliberativePolling import outputs"
"outputs("your_file.SAV")"
For example, "outputs("Sample.SAV")"

Conclusion: With the data appropriately classified and organized, you are now poised to employ the Python package for rigorous analysis of your Deliberative Polling experimental data.
