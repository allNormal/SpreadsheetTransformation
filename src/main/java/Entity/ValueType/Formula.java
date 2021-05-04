package Entity.ValueType;

import Entity.SheetElement.BasicElement.Cell;
import Entity.SheetElement.BasicElement.Column;

import java.util.ArrayList;
import java.util.List;

public abstract class Formula {
    private String formulaFunction;
    private List<Cell> cellDependency =  new ArrayList<>();
    private List<Column> columnDependency = new ArrayList<>();
    private Value valueType;
    private FunctionType functionType;
    private String stringValue;
    private Boolean booleanValue;
    private Byte errorValue;
    private float numericValue;


    public Formula(String formulaFunction) {
        this.formulaFunction = formulaFunction;
    }

    public void addDependencies(Cell cell) {
        cellDependency.add(cell);
    }

    public void addColumnDependencies(Column column) { columnDependency.add(column); }

    public List<Cell> getCellDependency() {
        return cellDependency;
    }

    public Value getValueType() {
        return valueType;
    }

    public void setValueType(Value valueType) {
        this.valueType = valueType;
    }

    public FunctionType getFunctionType() {
        return functionType;
    }

    public void setFunctionType(FunctionType functionType) {
        this.functionType = functionType;
    }

    public String getFormulaFunction() {
        return formulaFunction;
    }

    public void setFormulaFunction(String formulaFunction) {
        this.formulaFunction = formulaFunction;
    }

    public String getStringValue() {
        return stringValue;
    }

    public abstract void add(Formula formula);

    public void setStringValue(String stringValue) {
        this.stringValue = stringValue;
    }

    public Boolean getBooleanValue() {
        return booleanValue;
    }

    public void setBooleanValue(Boolean booleanValue) {
        this.booleanValue = booleanValue;
    }

    public Byte getErrorValue() {
        return errorValue;
    }

    public void setErrorValue(Byte errorValue) {
        this.errorValue = errorValue;
    }

    public float getNumericValue() {
        return numericValue;
    }

    public void setNumericValue(float numericValue) {
        this.numericValue = numericValue;
    }

    public List<Column> getColumnDependency() {
        return columnDependency;
    }

}
