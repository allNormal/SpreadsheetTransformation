package Entity;

public class WorksheetEditor {

    private String worksheetName;
    private int rowHeaderIndex;
    private int columnHeaderIndex;
    private String rowsFromAnotherWs;

    public WorksheetEditor(String worksheetName, int rowHeaderIndex, int columnHeaderIndex) {
        this.worksheetName = worksheetName;
        this.rowHeaderIndex = rowHeaderIndex;
        this.columnHeaderIndex = columnHeaderIndex;
    }

    public WorksheetEditor(String worksheetName, String rowsFromAnotherWs) {
        this.worksheetName = worksheetName;
        this.rowsFromAnotherWs = rowsFromAnotherWs;
    }

    public String getRowsFromAnotherWs() {
        return rowsFromAnotherWs;
    }

    public void setRowsFromAnotherWs(String rowsFromAnotherWs) {
        this.rowsFromAnotherWs = rowsFromAnotherWs;
    }

    public String getWorksheetName() {
        return worksheetName;
    }

    public void setWorksheetName(String worksheetName) {
        this.worksheetName = worksheetName;
    }

    public int getRowHeaderIndex() {
        return rowHeaderIndex;
    }

    public void setRowHeaderIndex(int rowHeaderIndex) {
        this.rowHeaderIndex = rowHeaderIndex;
    }

    public int getColumnHeaderIndex() {
        return columnHeaderIndex;
    }

    public void setColumnHeaderIndex(int columnHeaderIndex) {
        this.columnHeaderIndex = columnHeaderIndex;
    }
}
