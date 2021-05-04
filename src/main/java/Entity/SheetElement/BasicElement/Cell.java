package Entity.SheetElement.BasicElement;

import Entity.ValueType.ErrorValue;
import Entity.ValueType.Formula;
import Entity.ValueType.Value;
import Entity.Worksheet.Worksheet;
import org.apache.poi.xssf.usermodel.XSSFComment;

public class Cell extends BasicElement{

    private String cellId;
    private int row;
    private String rowID;
    private String column;
    private String columnTitle;
    private String stringValue;
    private float numericValue;
    private boolean booleanValue;
    private Formula formulaValue;
    private ErrorValue error;
    private Value value;
    private XSSFComment comment;
    private String tableName;
    private boolean sameValueAsColumnHeader = true;

    public Cell(Worksheet worksheet, String column, int row){
        super(worksheet, column+(row+1));
        this.column = column;
        this.row = row + 1;
        this.cellId = column + (row+1);
    }

    public void setStringValue(String value){
        try{
            if(this.value == Value.STRING){
                stringValue = value;
            }
            else{
                throw new IllegalArgumentException("Cell Values is not a String but a " + this.value);
            }
        } catch (IllegalArgumentException err){
            System.out.println(err);
        }
    }

    public void setNumericValue(float value){
        try{
            if(this.value == Value.NUMERIC){
                numericValue = value;
            }
            else{
                throw new IllegalArgumentException("Cell Values is not a Numeric but a " + this.value);
            }
        } catch (IllegalArgumentException err){
            System.out.println(err);
        }
    }



    public void setBooleanValue(boolean value){
        try{
            if(this.value == Value.BOOLEAN){
                booleanValue = value;
            }
            else{
                throw new IllegalArgumentException("Cell Values is not a Boolean but a " + this.value);
            }
        } catch (IllegalArgumentException err){
            System.out.println(err);
        }
    }

    public void setFormulaValue(Formula basicFormula){
        this.formulaValue = basicFormula;
    }

    public void SetErrorValue(String value){
        this.error = new ErrorValue(value);
    }

    public String getCellId() {
        return cellId;
    }

    public int getRow() {
        return row;
    }

    public String getColumn() {
        return column;
    }

    public void setValue(Value value) {
        this.value = value;
    }

    public void setComment(XSSFComment comment) {
        this.comment = comment;
    }

    public String getStringValue() {
        return stringValue;
    }

    public float getNumericValue() {
        return numericValue;
    }

    public boolean isBooleanValue() {
        return booleanValue;
    }

    public Formula getFormulaValue() {
        return formulaValue;
    }

    public ErrorValue getError() {
        return error;
    }

    public XSSFComment getComment() {
        return comment;
    }

    public Value getValue() {
        return value;
    }

    public String getRowID() {
        return rowID;
    }

    public void setRowID(String rowID) {
        this.rowID = rowID;
    }

    public String getColumnTitle() {
        return columnTitle;
    }

    public void setColumnTitle(String columnTitle) {
        this.columnTitle = columnTitle;
    }

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName;
    }

    public boolean isSameValueAsColumnHeader() {
        return sameValueAsColumnHeader;
    }

    public void setSameValueAsColumnHeader(boolean sameValueAsColumnHeader) {
        this.sameValueAsColumnHeader = sameValueAsColumnHeader;
    }
}
