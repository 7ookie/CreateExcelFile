# CreateExcelFile

The solution include two projects
  1* Factorial calculator (CalculateFactorial)
  2* Excel file creator which include also generator for simple excel file (ExcelFileGenerator)
  
  1* Factorial calculator (CalculateFactorial) - include two functions "CalculateFactorial()" and "CalculateFactorialOptimized()"
"CalculateFactorial()" is calculated like n! = n × (n−1)!
"CalculateFactorialOptimized()" here in use is different approach https://sites.google.com/site/examath/research/factorials

  2* Excel file creator (ExcelFileGenerator) - include two functions "CreateFile()" and "CreateTable()"
 Microsoft.Office.Interop.Excel is used in addition
 In "CreateTable()" the excel table is created with columns and fulfilled with sample data, then the table is returned.
 In "CreateFile()" the excel file is created, get the table from "CreateTable()" and write it to the file.
 
