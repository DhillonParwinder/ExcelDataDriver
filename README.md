# Data-driven testing

In a real-world application, Excel spreadsheet could be provided by the client or the end-user with the business logic encoded within the spreadsheet. (The POI library handles numerical calculations). In this scenario, the Excel spreadsheet becomes part of your acceptance tests, and helps to define your requirements, allows effective test-driven development of the code itself, and also acts as part of your acceptance tests.

This application program prefrom two tasks:
  * Derive data from Excel Spreadsheet and dispay on your console
  * Derive data from Excel Spreadhseet and store to Array, so pass to a testcase.
  
  The DeriveSpreadsheetData class uses the Apache POI to load data from Spreadsheet and display it onto Console.
  
  DataDrivenTestCase uses data for testcases which is deriven by DataDrivenInArray class.
  
  ## MS Excel Test Data ##
  
  The Excel data file is located in a spreadsheet in the 'ExcelDocs' folder.
  
  ## Technology Stack ##
   JDK : 14.0.1
   Gradle : 6.5.1
   JUnit : 4.12
   Apache POI : 4.1.2
   APache POI-OOXML :4.1.2
  
   
   
  
