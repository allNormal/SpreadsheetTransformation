package ExcelReader;

import Entity.Workbook.Workbook;

import java.io.IOException;

public interface ExcelReader {

    void readExcelConverter() throws IOException;

    Workbook getWorkbook();
}
