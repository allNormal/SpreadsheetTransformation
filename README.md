# SpreadsheetTransformation

SpreadsheetTransformation is a java library to transform excel data into a RDF triples.

## Installation

incoming

## Methods

### ExcelBasedReader
 ```java
 ExcelBasedReader excelBasedReader = new ExcelBasedReader(new File("path"));
 excelBasedReader.readExcelConverter();
 ```
 
 ExcelBasedReader take an excel file path as parameter, and readExcelConverter() is a function in ExcelBasedReader to read all the data, ExcelBasedReader is preparing the data for
 ExcelTransformer.
 
 ### TableBasedReader
 ```java
 TableBasedReader tableBasedReader = new TableBasedReader(new File("path"));
 tableBasedReader.readExcelConverter();
 ```
 same as ExcelBasedReader, the different is TableBasedReader prepare the data for TableTransformer.
 
 ### ExcelTransformer
 ```java
 ExcelTransformer excelTransformer = new ExcelTransformer(excelBasedReader);
 ```
 ExcelTransformer take an ExcelBasedReader as parameter. ExcelTransfromer use Excel Ontology model.
 
 #### create
 ```java
 excelTransformer.create();
 ```
 create function in ExcelTransformer is to convert the data in ExcelBasedReader into an Jena OntModel.
 
 #### saveModel
 ```java
 excelTransformer.saveModel("path","fileName");
 ```
 saveModel function take a path and a fileName as parameter, saveModel is to export the OntModel data into a .ttl file.

 
 ### TableTransformer
  ```java
 TableTransformer tableTransformer = new TableTransformer(excelBasedReader);
 ```
 TableTransformer take an TableBasedReader as parameter. TableTransfromer use Table Ontology model.
 
 #### create
 ```java
 excelTransformer.create();
 ```
 create function in TableTransformer is to convert the data in TableBasedReader into an Jena OntModel.
 
 #### saveModel
 ```java
 excelTransformer.saveModel("path","fileName");
 ```
 saveModel function take a path and a fileName as parameter, saveModel is to export the OntModel data into a .ttl file.

 #### IncorrectTypeException
 if an IncorrectTypeException occurs in TableTransformer its because there are a column in which 1 or more datatype found inside it
