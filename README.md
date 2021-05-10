# SpreadsheetTransformation

SpreadsheetTransformation is a java library to transform excel data into a RDF triples.

## Installation


create a jar file using this command.

```console
mvn clean package
```

And to install it in your project locally use this command

```console
mvn install:install-file -Dfile=path -DgroupId=io.github.allNormal -DartifactId=SpreadsheetTransformation -Dversion=0.0.1 -Dpackaging=jar -DgeneratePom=true
```

Change path into SpreadsheetTransformation jar file location.

Lastly add it in your maven dependency

```maven
        <dependency>
            <groupId>io.github.allNormal</groupId>
            <artifactId>SpreadsheetTransformation</artifactId>
            <version>0.0.1</version>
        </dependency>
```

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
 Same as ExcelBasedReader, the difference is TableBasedReader prepare the data for TableTransformer.
 
 ### ExcelTransformer
 
 ![excel_design](/src/main/resources/excel_ontology.png)
 
 ExcelTransformer use Ontology design as shown above.
 
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

 #### Model
 ```java
 excelTransformer.getModel()
 ```
 You can get the OntModel by using the getModel function.
 
 #### getCellDependencies
 
 ```java
 List<String> testList = excelTransformer.getCellDependencies(cellId, worksheetName);
 ```
 cellDependencies function takes 2 parameter, a cellId (case sensitive) and a worksheet name. 
 cellDependencies return a list of string.
 Example if an excel cell has a formula value = Sum(B3:C3).
 cellDependencies will return = [B3, C3].
 
 #### getReverseDependencies
 ```java
 List<String> testList = excelTransformer.getReverseDependencies(cellId, worksheetName);
 ```
 Like getDependencies, getReverseDependencies also takes 2 parameter and return a list of string.
 getReverseDependencies checks which formula function use the cellId.
 Example if an excel cell D3 has a formula value = Sum(B3:C3).
 getReverseDependencies("B3","worksheetName") will return = [D3].

 
 ### TableTransformer
 
 ![table_design](/src/main/resources/Excel-Table-Based.png)
 
 TableTransformer use Ontology design as shown above.
 
  ```java
 TableTransformer tableTransformer = new TableTransformer(tableBasedReader);
 ```
 TableTransformer take an TableBasedReader as parameter. TableTransfromer use Table Ontology model. For TableTransformer to run flawlessly, please make sure that each cell in
 1 column has the same datatype, and if the column is a formula type that all the cell have the same formula pattern.
 
 #### create
 ```java
 tableTransformer.create();
 ```
 create function in TableTransformer is to convert the data in TableBasedReader into an Jena OntModel.
 
 #### saveModel
 ```java
 tableTransformer.saveModel("path","fileName");
 ```
 saveModel function take a path and a fileName as parameter, saveModel is to export the OntModel data into a .ttl file.

 #### Model
 ```java
 tableTransformer.getModel()
 ```
 You can get the OntModel by using the getModel function.
 
 #### getCellDependencies
 
 ```java
 List<String> testList = tableTransformer.getCellDependencies(cellId, worksheetName);
 ```
 cellDependencies function takes 2 parameter, a cellId (case sensitive) and a worksheet name. 
 cellDependencies return a list of string.
 Different than cellDependencies in ExcelTransformer, in cellDependencies TableTransformer it will return the ColumnHeader name.
 Example if an excel cell has a formula value = Sum(B3:C3) and the ColumnHeader from B3 = Price and C3 = Quantity.
 cellDependencies will return = [Price, Quantity].
 
 #### getReverseDependencies
 ```java
 List<String> testList = tableTransformer.getReverseDependencies(cellId, worksheetName);
 ```
 Like getDependencies, getReverseDependencies also takes 2 parameter and return a list of string.
 getReverseDependencies checks which formula function use the column header from the cellId.
 Example if an excel cell D3 has a formula value = Sum(B3:C3) and ColumnHeader from D3 = Total.
 getReverseDependencies("B3","worksheetName") will return = [Total].


 #### IncorrectTypeException
 If an IncorrectTypeException occurs in TableTransformer its because there are a column in which 1 or more datatype found inside it
