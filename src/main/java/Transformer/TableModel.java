package Transformer;

import org.apache.jena.ontology.DatatypeProperty;
import org.apache.jena.ontology.ObjectProperty;
import org.apache.jena.ontology.OntClass;
import org.apache.jena.ontology.OntModel;
import org.apache.jena.rdf.model.ModelFactory;

import java.nio.file.Path;
import java.nio.file.Paths;

import static org.apache.jena.ontology.OntModelSpec.OWL_MEM;

public class TableModel {

    private final String NS = "http://www.semanticweb.org/GregorKaefer/ontologies/2020/11/excelOntologyVersion2#";
    private OntModel model;
    private String MODEL_NAME;

    private OntClass workbook;
    private OntClass worksheet;
    private OntClass valueType;
    private OntClass formula;
    private OntClass string;
    private OntClass number;
    private OntClass bool;
    private OntClass error;
    private OntClass table;
    private OntClass rowHeader;
    private OntClass columnHeader;
    private OntClass value;

    private ObjectProperty hasColumnHeader;
    private ObjectProperty hasRowHeader;
    private ObjectProperty hasDependency;
    private ObjectProperty hasTable;
    private ObjectProperty hasValue;
    private ObjectProperty hasValueType;
    private ObjectProperty hasWorksheet;
    private ObjectProperty isUsedIn;
    private ObjectProperty valueFor;

    private DatatypeProperty actualValue;
    private DatatypeProperty columnHeaderID;
    private DatatypeProperty rowHeaderID;
    private DatatypeProperty columnHeaderTitle;
    private DatatypeProperty rowHeaderTitle;
    private DatatypeProperty formulaFunction;
    private DatatypeProperty rowEnd;
    private DatatypeProperty rowStart;
    private DatatypeProperty sheetName;
    private DatatypeProperty tableName;
    private DatatypeProperty workbookName;

    TableModel(String name){
        this.MODEL_NAME = name;
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
        this.workbook = this.model.getOntClass(NS + "Workbook");
        this.worksheet = this.model.getOntClass(NS + "Worksheet");
        this.table = this.model.getOntClass(NS + "Table");
        this.value = this.model.getOntClass(NS + "Value");
        this.valueType = this.model.getOntClass(NS + "ValueType");
        this.columnHeader = this.model.getOntClass(NS + "ColumnHeader");
        this.rowHeader = this.model.getOntClass(NS + "RowHeader");
        this.valueType = this.model.getOntClass(NS + "ValueType");
        this.formula = this.model.getOntClass(NS + "Formula");
        this.string = this.model.getOntClass(NS + "String");
        this.bool = this.model.getOntClass(NS + "Boolean");
        this.error = this.model.getOntClass(NS + "Error");
        this.number = this.model.getOntClass(NS + "Number");

        this.hasColumnHeader = this.model.getObjectProperty(NS + "hasColumnHeader");
        this.hasDependency = this.model.getObjectProperty(NS + "hasDependency");
        this.hasTable = this.model.getObjectProperty(NS + "hasTable");
        this.hasValue = this.model.getObjectProperty(NS + "hasValue");
        this.hasWorksheet = this.model.getObjectProperty(NS + "hasWorksheet");
        this.hasRowHeader = this.model.getObjectProperty(NS + "hasRowHeader");
        this.hasValueType = this.model.getObjectProperty(NS + "hasValueType");
        this.isUsedIn = this.model.getObjectProperty(NS + "isUsedIn");
        this.valueFor = this.model.getObjectProperty(NS + "ValueFor");

        this.actualValue = this.model.getDatatypeProperty(NS + "ActualValue");
        this.columnHeaderID = this.model.getDatatypeProperty(NS + "ColumnHeaderID");
        this.columnHeaderTitle = this.model.getDatatypeProperty(NS + "ColumnHeaderTitle");
        this.formulaFunction = this.model.getDatatypeProperty(NS + "FormulaFunction");
        this.rowEnd = this.model.getDatatypeProperty(NS + "RowEnd");
        this.rowStart = this.model.getDatatypeProperty(NS + "RowStart");
        this.rowHeaderID = this.model.getDatatypeProperty(NS + "RowHeaderID");
        this.rowHeaderTitle = this.model.getDatatypeProperty(NS + "RowHeaderTitle");
        this.sheetName = this.model.getDatatypeProperty(NS + "SheetName");
        this.tableName = this.model.getDatatypeProperty(NS + "TableName");
        this.workbookName = this.model.getDatatypeProperty(NS + "WorkbookName");
    }

    /**
     * get model from dataset and load it.
     */
    private void loadModels(){
        this.model = ModelFactory.createOntologyModel(OWL_MEM);
        Path path = Paths.get("excel_owl/excel_ontology_version2.ttl");
        this.model.read(path.toString());

    }

    public String getNS() {
        return NS;
    }

    public OntModel getModel() {
        return model;
    }

    public String getMODEL_NAME() {
        return MODEL_NAME;
    }

    public OntClass getWorkbook() {
        return workbook;
    }

    public OntClass getWorksheet() {
        return worksheet;
    }

    public OntClass getValueType() {
        return valueType;
    }

    public OntClass getFormula() {
        return formula;
    }

    public OntClass getString() {
        return string;
    }

    public OntClass getNumber() {
        return number;
    }

    public OntClass getBool() {
        return bool;
    }

    public OntClass getError() {
        return error;
    }

    public OntClass getTable() {
        return table;
    }

    public OntClass getRowHeader() {
        return rowHeader;
    }

    public OntClass getColumnHeader() {
        return columnHeader;
    }

    public OntClass getValue() {
        return value;
    }

    public ObjectProperty getHasColumnHeader() {
        return hasColumnHeader;
    }

    public ObjectProperty getHasRowHeader() {
        return hasRowHeader;
    }

    public ObjectProperty getHasDependency() {
        return hasDependency;
    }

    public ObjectProperty getHasTable() {
        return hasTable;
    }

    public ObjectProperty getHasValue() {
        return hasValue;
    }

    public ObjectProperty getHasValueType() {
        return hasValueType;
    }

    public ObjectProperty getHasWorksheet() {
        return hasWorksheet;
    }

    public ObjectProperty getIsUsedIn() {
        return isUsedIn;
    }

    public ObjectProperty getValueFor() {
        return valueFor;
    }

    public DatatypeProperty getActualValue() {
        return actualValue;
    }

    public DatatypeProperty getColumnHeaderID() {
        return columnHeaderID;
    }

    public DatatypeProperty getRowHeaderID() {
        return rowHeaderID;
    }

    public DatatypeProperty getColumnHeaderTitle() {
        return columnHeaderTitle;
    }

    public DatatypeProperty getRowHeaderTitle() {
        return rowHeaderTitle;
    }

    public DatatypeProperty getFormulaFunction() {
        return formulaFunction;
    }

    public DatatypeProperty getRowEnd() {
        return rowEnd;
    }

    public DatatypeProperty getRowStart() {
        return rowStart;
    }

    public DatatypeProperty getSheetName() {
        return sheetName;
    }

    public DatatypeProperty getTableName() {
        return tableName;
    }

    public DatatypeProperty getWorkbookName() {
        return workbookName;
    }

}
