package Transformer;

import org.apache.jena.ontology.DatatypeProperty;
import org.apache.jena.ontology.ObjectProperty;
import org.apache.jena.ontology.OntClass;
import org.apache.jena.ontology.OntModel;
import org.apache.jena.rdf.model.ModelFactory;

import static org.apache.jena.ontology.OntModelSpec.OWL_MEM;

public class ExcelModel {
    private OntModel model;

    private final String NS = "http://www.semanticweb.org/GregorKaefer/ontologies/2020/8/excelOntology#";
    private String MODEL_NAME;

    private OntClass workbook;
    private OntClass worksheet;
    private OntClass chart;
    private OntClass table;
    private OntClass sheetElement;
    private OntClass basicElement;
    private OntClass cell;
    private OntClass row;
    private OntClass column;
    private OntClass macro;
    private OntClass Value;
    private OntClass booleanValue;
    private OntClass formulaValue;
    private OntClass stringValue;
    private OntClass errorValue;
    private OntClass numericValue;
    private OntClass comment;
    private OntClass illustration;
    private OntClass text;

    private DatatypeProperty cellID;
    private DatatypeProperty cellStatus;
    private DatatypeProperty cellValue;
    private DatatypeProperty chartLegend;
    private DatatypeProperty chartTitle;
    private DatatypeProperty columnID;
    private DatatypeProperty commentsValue;
    private DatatypeProperty description;
    private DatatypeProperty elementName;
    private DatatypeProperty errorName;
    private DatatypeProperty extension;
    private DatatypeProperty fileName;
    private DatatypeProperty filePath;
    private DatatypeProperty functionName;
    private DatatypeProperty macroTitle;
    private DatatypeProperty macroInput;
    private DatatypeProperty rowID;
    private DatatypeProperty sheetName;
    private DatatypeProperty colStart;
    private DatatypeProperty colEnd;
    private DatatypeProperty rowStart;
    private DatatypeProperty rowEnd;

    private ObjectProperty hasCell;
    private ObjectProperty hasColumn;
    private ObjectProperty hasComment;
    private ObjectProperty hasInput;
    private ObjectProperty hasMacro;
    private ObjectProperty hasRow;
    private ObjectProperty hasWorksheet;
    private ObjectProperty hasSheetElement;
    private ObjectProperty hasValue;
    private ObjectProperty isUsedIn;
    private ObjectProperty isPartOfWorksheet;
    private ObjectProperty hasConstraint;
    private ObjectProperty isInWorkbook;

    public ExcelModel(String modelName) {
        this.MODEL_NAME = modelName;
        initializeTemplateModel();

    }


    /**
     * initialize model + link model if there is no such model in dataset.
     */
    private void initializeTemplateModel(){
        loadModels();
        initializeModels();
    }


    /**
     * Link every Class, Object property, Data Property, with the one in dataset.
     */
    private void initializeModels() {
        this.hasMacro = this.model.getObjectProperty(NS + "hasMacro");
        this.macro = this.model.getOntClass(NS + "ExcelMacro");
        this.macroInput = this.model.getDatatypeProperty(NS + "Macro");
        this.cell = this.model.getOntClass(NS + "cell");
        this.row = this.model.getOntClass(NS + "row");
        this.column = this.model.getOntClass(NS + "column");
        this.workbook = this.model.getOntClass(NS + "Workbook");
        this.worksheet = this.model.getOntClass(NS + "Worksheet");
        this.sheetElement = this.model.getOntClass(NS + "SheetElement");
        this.basicElement = this.model.getOntClass(NS + "BasicElement");
        this.chart = this.model.getOntClass(NS + "Chart");
        this.comment = this.model.getOntClass(NS + "Comments");
        this.table = this.model.getOntClass(NS + "Table");
        this.illustration = this.model.getOntClass(NS + "Illustration");
        this.text = this.model.getOntClass(NS + "Text");
        this.Value = this.model.getOntClass(NS + "Value");
        this.errorValue = this.model.getOntClass(NS + "ErrorValue");
        this.formulaValue = this.model.getOntClass(NS + "FormulaValue");
        this.stringValue = this.model.getOntClass(NS + "StringValue");
        this.booleanValue = this.model.getOntClass(NS + "BooleanValue");
        this.numericValue = this.model.getOntClass(NS + "NumericValue");
        this.hasCell = this.model.getObjectProperty(NS + "hasCell");
        this.hasColumn = this.model.getObjectProperty(NS + "hasColumn");
        this.hasComment = this.model.getObjectProperty(NS + "hasComment");
        this.hasInput = this.model.getObjectProperty(NS + "hasInput");
        this.hasRow = this.model.getObjectProperty(NS + "hasRow");
        this.hasSheetElement = this.model.getObjectProperty(NS + "hasSheetElement");
        this.hasValue = this.model.getObjectProperty(NS + "hasValue");
        this.hasWorksheet = this.model.getObjectProperty(NS + "hasWorkSheet");
        this.isPartOfWorksheet = this.model.getObjectProperty(NS + "isPartOfWorksheet");
        this.isUsedIn = this.model.getObjectProperty(NS + "isUsedIn");
        this.isInWorkbook = this.model.getObjectProperty(NS + "isInWorkbook");
        this.description = this.model.getDatatypeProperty(NS + "Description");
        this.cellID = this.model.getDatatypeProperty(NS + "CellID");
        this.cellStatus = this.model.getDatatypeProperty(NS + "CellStatus");
        this.cellValue = this.model.getDatatypeProperty(NS + "CellValue");
        this.chartLegend = this.model.getDatatypeProperty(NS + "ChartLegend");
        this.chartTitle = this.model.getDatatypeProperty(NS + "ChartTitle");
        this.columnID = this.model.getDatatypeProperty(NS + "ColumnID");
        this.commentsValue = this.model.getDatatypeProperty(NS + "CommentsValue");
        this.elementName = this.model.getDatatypeProperty(NS + "ElementName");
        this.errorName = this.model.getDatatypeProperty(NS + "ErrorName");
        this.extension = this.model.getDatatypeProperty(NS + "Extension");
        this.fileName = this.model.getDatatypeProperty(NS + "FileName");
        this.filePath = this.model.getDatatypeProperty(NS + "FilePath");
        this.functionName = this.model.getDatatypeProperty(NS + "FunctionName");
        this.macroTitle = this.model.getDatatypeProperty(NS + "MacroTitle");
        this.rowID = this.model.getDatatypeProperty(NS + "RowID");
        this.sheetName = this.model.getDatatypeProperty(NS + "SheetName");
        this.colStart = this.model.getDatatypeProperty(NS + "X-Start");
        this.colEnd = this.model.getDatatypeProperty(NS + "X-End");
        this.rowStart = this.model.getDatatypeProperty(NS + "Y-Start");
        this.rowEnd = this.model.getDatatypeProperty(NS + "Y-End");
    }

    /**
     * get model from dataset and load it.
     */
    private void loadModels(){
        this.model = ModelFactory.createOntologyModel(OWL_MEM);
        this.model.read("src/main/resources/excel_owl/excel_ontology.ttl");

    }


    public DatatypeProperty getMacroInput() {
        return macroInput;
    }

    public OntModel getModel() {
        return model;
    }


    public OntClass getWorkbook() {
        return workbook;
    }

    public OntClass getWorksheet() {
        return worksheet;
    }

    public OntClass getChart() {
        return chart;
    }

    public OntClass getTable() {
        return table;
    }

    public OntClass getSheetElement() {
        return sheetElement;
    }

    public OntClass getBasicElement() {
        return basicElement;
    }

    public OntClass getCell() {
        return cell;
    }

    public OntClass getRow() {
        return row;
    }

    public OntClass getColumn() {
        return column;
    }

    public OntClass getMacro() {
        return macro;
    }

    public OntClass getValue() {
        return Value;
    }

    public OntClass getBooleanValue() {
        return booleanValue;
    }

    public OntClass getFormulaValue() {
        return formulaValue;
    }

    public OntClass getStringValue() {
        return stringValue;
    }

    public OntClass getErrorValue() {
        return errorValue;
    }

    public OntClass getNumericValue() {
        return numericValue;
    }

    public OntClass getComment() {
        return comment;
    }

    public OntClass getIllustration() {
        return illustration;
    }

    public OntClass getText() {
        return text;
    }

    public DatatypeProperty getCellID() {
        return cellID;
    }

    public DatatypeProperty getCellStatus() {
        return cellStatus;
    }

    public DatatypeProperty getCellValue() {
        return cellValue;
    }

    public DatatypeProperty getChartLegend() {
        return chartLegend;
    }

    public DatatypeProperty getChartTitle() {
        return chartTitle;
    }

    public DatatypeProperty getColumnID() {
        return columnID;
    }

    public DatatypeProperty getCommentsValue() {
        return commentsValue;
    }

    public DatatypeProperty getDescription() {
        return description;
    }

    public DatatypeProperty getElementName() {
        return elementName;
    }

    public DatatypeProperty getErrorName() {
        return errorName;
    }

    public DatatypeProperty getExtension() {
        return extension;
    }

    public DatatypeProperty getFileName() {
        return fileName;
    }

    public DatatypeProperty getFilePath() {
        return filePath;
    }

    public DatatypeProperty getFunctionName() {
        return functionName;
    }

    public DatatypeProperty getMacroTitle() {
        return macroTitle;
    }

    public DatatypeProperty getRowID() {
        return rowID;
    }

    public DatatypeProperty getSheetName() {
        return sheetName;
    }

    public DatatypeProperty getColStart() {
        return colStart;
    }

    public DatatypeProperty getColEnd() {
        return colEnd;
    }

    public DatatypeProperty getRowStart() {
        return rowStart;
    }

    public DatatypeProperty getRowEnd() {
        return rowEnd;
    }

    public ObjectProperty getHasCell() {
        return hasCell;
    }

    public ObjectProperty getHasColumn() {
        return hasColumn;
    }

    public ObjectProperty getHasComment() {
        return hasComment;
    }

    public ObjectProperty getHasInput() {
        return hasInput;
    }

    public ObjectProperty getHasMacro() {
        return hasMacro;
    }

    public ObjectProperty getHasRow() {
        return hasRow;
    }

    public ObjectProperty getHasWorksheet() {
        return hasWorksheet;
    }

    public ObjectProperty getHasSheetElement() {
        return hasSheetElement;
    }

    public ObjectProperty getHasValue() {
        return hasValue;
    }

    public ObjectProperty getIsUsedIn() {
        return isUsedIn;
    }

    public ObjectProperty getIsPartOfWorksheet() {
        return isPartOfWorksheet;
    }

    public ObjectProperty getHasConstraint() {
        return hasConstraint;
    }

    public ObjectProperty getIsInWorkbook() {
        return isInWorkbook;
    }

}
