# Parse CloudSploit Command Line scan results into an observations sheet.

This python script will format CloudSploit CSV output taken from the **command line interface** version of the tool into an excel spreadsheet. The spreadsheet will consist of a single observations table consisting of failed test results aggregated by the test titles. The tool will also create two charts (a pie chart of failed tests by observation domains and a bar chart of failed tests by assigned risk level.)<sup>[[1]](#footnote1)</sup>. Worksheet names in the output file will default to the filename of the raw CSV results.

*Note: The tool will overwrite files if the output filename is the same as one that already exists. It is recommended to use the -o flag to make sure the default 'observations.xlsx' file is not overwritten.*

<a name="footnote1">[1]</a> Modification of chart data can be done by unhiding the 'ChartData' worksheet and editting the relevant values. The tool does not match the charts to RSM's normal design themes. It is recommended to edit the colors to match the firm's typical coloring.

The tool supports various command line arguments that will change its behavior.

|Flag|Elaborated Flag|Description|
|:---|:---|:---|
|-o|--output|Determines the output filename. Defaults to 'results.csv'. If the specified output filename or default already exists, it will be appended to.|
|-v|--verbose|Increase verbosity of standard output to the shell.|
|-t|--target|Specifies a single CSV file parse. This mode of operation will also append the raw results to an additional spreadsheet for easy reference without switching windows.*Note: Mutually exclusive with -l and -d flags*|
|-l|--list|Instructs the tool to read from a file containing a list of CSV results to format. The list is expected to be a text file with a single filename on each line. A single spreadsheet is created for each CSV file listed. *Note: Mutually exclusive with -t and -d flags*|
|-d|--directory|Causes the script to recursively search for every CSV file contained within the specified directory. A single spreadsheet is created for each CSV file found. *Note: Mutually exclusive with the -t and -l flags*|
|-z|--zip|Makes the tool create a second compressed version of the resulting CSV. Useful when merging a high volume of files.|
|-h|--help|Print an example of tool usage and exit.|

---

### Some examples of usage can be seen as follows:

Parse the 'client_scan.csv' results into a new 'preliminary_observations.xlsx' workbook. Run with increased verbosity.
```
py format_cloudsploit_cli.py -v -t ./path/to/a/client_scan.csv -o preliminary_observations.xlsx
```
After creating a list of CSV files in 'list.txt', parse all of those scan files into a 'merged.xlsx' workbook.
```
py format_cloudsploit_cli.py -l ./list.txt -o merged.xlsx
```
Merge all CloudSploit results within the 'ClientScans' folder into a single workbook (Defaults to observations.xlsx). Zip the resulting output (Makes a second observations.xlsx.zip file).
```
py format_cloudsploit_cli.py -z -d ./ClientScans
```
