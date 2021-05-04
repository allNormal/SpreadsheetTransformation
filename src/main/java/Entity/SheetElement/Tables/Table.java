package Entity.SheetElement.Tables;

import Entity.SheetElement.BasicElement.Cell;
import Entity.SheetElement.BasicElement.Column;
import Entity.SheetElement.BasicElement.Row;
import Entity.SheetElement.SheetElement;
import Entity.Worksheet.Worksheet;

import java.util.ArrayList;
import java.util.List;

public class Table extends SheetElement {
    private String rowStart;
    private String rowEnd;
    private String columnStart;
    private String columnEnd;
    private List<Cell> cell = new ArrayList<>();
    private List<Column> columns = new ArrayList<>();
    private List<Row> rows = new ArrayList<>();
    private String elementName;



    public Table(Worksheet worksheet, String columnEnd, String columnStart, String rowStart, String rowEnd
            , String elementName) {
        super(worksheet);
        this.columnEnd = columnEnd;
        this.columnStart = columnStart;
        this.rowEnd = rowEnd;
        this.rowStart = rowStart;
        this.elementName = elementName;
    }

    public Table(Worksheet worksheet, String elementName) {
        super(worksheet);
        this.elementName = elementName;
    }

    public void addCell(Cell cell){
        this.cell.add(cell);
    }

    public void addRowHeader(Row row) {
        this.rows.add(row);
    }

    public void addColumnHeader(Column column) {
        this.columns.add(column);
    }

    public String getRowStart() {
        return rowStart;
    }

    public void setRowStart(String rowStart) {
        this.rowStart = rowStart;
    }

    public String getRowEnd() {
        return rowEnd;
    }

    public void setRowEnd(String rowEnd) {
        this.rowEnd = rowEnd;
    }

    public String getColumnStart() {
        return columnStart;
    }

    public void setColumnStart(String columnStart) {
        this.columnStart = columnStart;
    }

    public String getColumnEnd() {
        return columnEnd;
    }

    public void setColumnEnd(String columnEnd) {
        this.columnEnd = columnEnd;
    }

    public List<Cell> getCell() {
        return cell;
    }

    public String id() {
        return elementName;
    }

    public void setElementName(String elementName) {
        this.elementName = elementName;
    }

    public void setCell(List<Cell> cell) {
        this.cell = cell;
    }

    public void setColumns(List<Column> columns) {
        this.columns = columns;
    }

    public void setRows(List<Row> rows) {
        this.rows = rows;
    }

    public List<Column> getColumns() {
        return columns;
    }

    public List<Row> getRows() {
        return rows;
    }

}
