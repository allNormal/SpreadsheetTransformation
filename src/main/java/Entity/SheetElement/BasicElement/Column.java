package Entity.SheetElement.BasicElement;

import Entity.ValueType.Formula;
import Entity.ValueType.Value;
import Entity.Worksheet.Worksheet;


import java.util.ArrayList;
import java.util.List;

public class Column extends BasicElement{

    private String columnID;
    private String columnTitle;
    private Formula formulaValue;
    private List<Cell> cell = new ArrayList<>();
    private Value value;

    public Column(Worksheet worksheet, String columnID) {
        super(worksheet, columnID);
        this.columnID = columnID;
    }

    public Column(Worksheet worksheet, String columnID, String columnTitle) {
        super(worksheet, columnID);
        this.columnID = columnID;
        this.columnTitle = columnTitle;
    }

    public void addCell(Cell cell){
        this.cell.add(cell);
    }

    public String getColumnID() {
        return columnID;
    }

    public List<Cell> getCell() {
        return cell;
    }

    public String getColumnTitle() {
        return columnTitle;
    }

    public void setColumnTitle(String columnTitle) {
        this.columnTitle = columnTitle;
    }

    public void setValue(Value value) {
        this.value = value;
    }

    public Value getValue() {
        return value;
    }

    public Formula getFormulaValue() {
        return formulaValue;
    }

    public void setFormulaValue(Formula formulaValue) {
        this.formulaValue = formulaValue;
    }

}
