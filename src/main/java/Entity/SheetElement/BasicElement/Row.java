package Entity.SheetElement.BasicElement;

import Entity.Worksheet.Worksheet;

import java.util.ArrayList;
import java.util.List;

public class Row extends BasicElement{

    private String rowId;
    private String rowTitle;
    private String columnTitle;
    private List<Cell> cell = new ArrayList<>();

    public Row(Worksheet worksheet, String rowId) {
        super(worksheet, rowId);
        this.rowId = ""+(Integer.parseInt(rowId) + 1);
    }

    public void addCell(Cell cell){
        this.cell.add(cell);
    }

    public String getRowId() {
        return rowId;
    }

    public void setRowId(String rowId) {
        this.rowId = rowId;
    }

    public List<Cell> getCell() {
        return cell;
    }

    public String getRowTitle() {
        return rowTitle;
    }


    public void setRowTitle(String rowTitle) {
        this.rowTitle = rowTitle;
    }

    public String getColumnTitle() {
        return columnTitle;
    }

    public void setColumnTitle(String columnTitle) {
        this.columnTitle = columnTitle;
    }

}
