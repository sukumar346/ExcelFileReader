# Cli
- Prints MetaData of XLS files
- Prints Worksheets in clean table format in Terminal
- Prints summary about the given Worksheet
- Prints output for given Query (for perticular Column)

INSTALLATION:

git clone https://github.com/sukumar346/RunExcelReader 

mvn clean install

<dependency>
  <groupId>com.innovative</groupId>
  <artifactId>RunExcelReader</artifactId>
  <version>2.0</version>
</dependency>

 METHODS:

1. Cli.getWorksheet(String filePath, int sheetIndex);

 /*
        *This method displays the worksheet in table format and the summary.
        * Input: file path, worksheet number
        * Output: prints Table, Row Counts, Column Counts, Data Types of each column
 */

2. Cli.MetadataDetails(String filePath)

/*
        * This method prints Metadata of the Excel File
        * Input:
        * @param filePath : it takes file path of the Excel File
        * Output: Prints Total No of worksheets, List of worksheets, summary of all worksheets
 */
 
 3.Cli.getQueryPrint(String filePath, int sheetIndex, int columnNum, char operator, String Operand)
 
/*
        * This method filters the given worksheet by the column and condition and prints it.
        * input:
        * @param filePath : Path of Excel file
        * @param sheetIndex: Sheet number
        * @param columnNum: column number in the worksheet
        * @operator: '=' or '<' or '>' (for String DataType, only '=')
        * @operand: any number or string
        * output: prints filtered table
 */
