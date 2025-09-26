# dynamic-file-generator
Populate placeholders with values on template files

## How to use

To use the generator, two files are necessary:
 - The model file (.pptx)
 - The data file (.csv or .txt)

While running the generator, select the files indicated above and the destination folder, where the dynamic generated files should be saved.

The program also allows to automatically export the files to PDF, selecting the option "Gerar arquivos PDF".

### Model
The model file is a normal PowerPoint file (.pptx) with placeholders for the variable data. The only requirement for the model file is that each placeholder has a name in the markup tag format, a text between '<' and '>' symbols.

#### Example:

```
<full-name>
```
In the model file, use the same presentation of text, for the placeholders, you want the variable data to have in the generated files.

### Data file
The variable data is taken from a Comma-Separated Values file (.csv). The first line must contain the names of the placeholders separated by ',' (comma) or ';' (semicolon). The remaining lines must contain the data to be inserted on the placeholders, in the same order which their respective variable names appear in the first line.

The generated files are named after the values of the first variable of each line.

#### Example:

```csv
full-name; age
John Smith; 32
Jane Dee; 25
```
CSV files can also be generated from Excel.

## Author

- @yannicktmessias

