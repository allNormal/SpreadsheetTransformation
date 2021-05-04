package Entity.Worksheet;

import Entity.SheetElement.ElementType;
import Entity.SheetElement.SheetElement;
import Entity.Workbook.Workbook;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Worksheet {
    private Map<ElementType, List<SheetElement>> sheets = new HashMap<>();
    private String sheetName;
    private Workbook workbook;
    private int columnHeaderIndex = -1;
    private int rowHeaderIndex = -1;
    private String rowHeaderFrom;

    public Worksheet(String sheetName, Workbook workbook) {
        this.sheetName = sheetName;
        this.workbook = workbook;
    }

    public Worksheet(String sheetName, Workbook workbook, int columnHeaderIndex, int rowHeaderIndex) {
        this.sheetName = sheetName;
        this.workbook = workbook;
        this.columnHeaderIndex = columnHeaderIndex;
        this.rowHeaderIndex = rowHeaderIndex;
    }

    public void addElement(ElementType type, List<SheetElement> sheetElement){
        this.sheets.put(type, sheetElement);
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Map<ElementType, List<SheetElement>> getSheets() {
        return sheets;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public int getColumnHeaderIndex() {
        return columnHeaderIndex;
    }

    public void setColumnHeaderIndex(int columnHeaderIndex) {
        this.columnHeaderIndex = columnHeaderIndex;
    }

    public int getRowHeaderIndex() {
        return rowHeaderIndex;
    }

    public void setRowHeaderIndex(int rowHeaderIndex) {
        this.rowHeaderIndex = rowHeaderIndex;
    }

    public String getRowHeaderFrom() {
        return rowHeaderFrom;
    }

    public void setRowHeaderFrom(String rowHeaderFrom) {
        this.rowHeaderFrom = rowHeaderFrom;
    }

}
