package ExcelReader;

import Entity.SheetElement.BasicElement.Column;
import Entity.SheetElement.Charts.Chart;
import Entity.SheetElement.ElementType;
import Entity.SheetElement.Illustrations.Illustration;
import Entity.SheetElement.SheetElement;
import Entity.SheetElement.Tables.Table;
import Entity.SheetElement.Texts.Text;
import Entity.ValueType.Formula;
import Entity.ValueType.FunctionType;
import Entity.ValueType.NestedFormula;
import Entity.ValueType.Value;
import Entity.Workbook.Macro;
import Entity.Workbook.Workbook;
import Entity.Worksheet.Worksheet;
import org.apache.logging.log4j.core.util.FileUtils;
import org.apache.poi.poifs.macros.VBAMacroReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelBasedReader implements ExcelReader{
    private XSSFWorkbook myWorkBook;
    private Workbook workbook;
    private File file;
    private final Logger log = LoggerFactory.getLogger(ExcelBasedReader.class);


    public ExcelBasedReader(File file) throws IOException {
        this.file = file;
        //readExcelConverter();
    }

    /*
    public void initializeWorkbook() throws IOException {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            System.out.println(e.getMessage());
        }

        // Finds the workbook instance for XLSX file
        myWorkBook = new XSSFWorkbook (fis);
        this.workbook = new Workbook((FileUtils.getFileExtension(file)),file.getName().replaceAll(" ", "_"));
        this.workbookExcel.setWorkbookName(this.workbook.getFileName());
        List<String> worksheets = new ArrayList<>();
        //iterate through sheet in workbook
        for(Sheet sheet : myWorkBook) {
            Worksheet worksheet = new Worksheet(sheet.getSheetName().replaceAll(" ", ""), workbook);
            worksheets.add(worksheet.getSheetName());
            workbook.addWorksheet(worksheet);
        }

        this.workbookExcel.setWorksheets(worksheets);
    }


     */
    /**
     * convert excel file into Workbook Object.
     * @throws IOException if file not found.
     */
    public void readExcelConverter() throws IOException {

        FileInputStream fis = null;
        try {
            fis = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            log.error(e.getMessage());
        }

        // Finds the workbook instance for XLSX file
        myWorkBook = new XSSFWorkbook (fis);
        this.workbook = new Workbook((FileUtils.getFileExtension(file)),file.getName().replaceAll(" ", "_"));
        //iterate through sheet in workbook
        log.info("iterating through the worksheet in workbook and convert all the data found into entity");
        for(Sheet sheet : myWorkBook){
            Worksheet worksheet = new Worksheet(sheet.getSheetName().replaceAll(" ", ""), workbook);
            log.info(sheet.getSheetName() + " converting basic element if found into entity");
            readBasicElement(sheet, worksheet);
            log.info(sheet.getSheetName() + " converting table if found into entity");
            readTable(myWorkBook, sheet.getSheetName(), worksheet);
            log.info(sheet.getSheetName() + " converting chart if found into entity");
            readChart(myWorkBook, sheet.getSheetName(), worksheet);
            log.info(sheet.getSheetName() + " converting illustration if found into entity");
            readIllustration(myWorkBook, sheet.getSheetName(), worksheet);
            log.info(sheet.getSheetName() + " converting text if found into entity");
            readText(myWorkBook, sheet.getSheetName(), worksheet);
            this.workbook.addWorksheet(worksheet);
        }

        log.info("converting all the formula value found into a formula dependecy like entity");
        for(int i = 0; i< this.workbook.getWorksheets().size(); i++) {
            Worksheet ws = this.workbook.getWorksheets().get(i);
            List<SheetElement> cell = ws.getSheets().get(ElementType.CELL);
            formulaCellDependencyCheck(cell);
        }

        //check if extension = macro or not
        /*
        if(this.workbook.getExtension().equals("xlsm")) {
            readMacro(file);
        }

         */

    }

    /*
    @Override
    public void readExcelConverterWithRestriction(Restriction restriction) throws IOException {

    }

     */

    /**
     * read Macro and save it in workbook
     * @param file com.java.eto.entity.excel file with macro extension
     */
    private void readMacro(File file) {
        try {
            VBAMacroReader macroReader = new VBAMacroReader(file);
            if(macroReader != null) {
                Macro m = new Macro(macroReader.toString());
                macroReader.close();
                this.workbook.setMacro(m);
            }
        } catch (IOException err) {
            log.error(err.getMessage());
        }

    }

    /**
     * read text sheet element from com.java.eto.entity.excel, convert it into text object and add it into worksheet object
     * @param workbook workbook target(poi)
     * @param sheetName worksheet target (poi)
     * @param worksheet add text element into worksheet object(self-made)
     */
    private void readText(XSSFWorkbook workbook, String sheetName, Worksheet worksheet) {
        List<SheetElement> texts = new ArrayList<>();
        try {
            XSSFSheet sheet = workbook.getSheet(sheetName);
            if(sheet == null) throw new FileNotFoundException("there is not sheet with name " + sheetName);
            XSSFDrawing shape = sheet.getDrawingPatriarch();

            if(shape != null) {
                List<XSSFShape> shapes = shape.getShapes();

                for(int i = 0;i<shapes.size();i++){
                    if(shapes.get(i) instanceof XSSFSimpleShape){
                        XSSFSimpleShape temp = (XSSFSimpleShape)shapes.get(i);
                        Text text = new Text(worksheet, "Text", temp.getText());
                        texts.add(text);
                    }
                }
                worksheet.addElement(ElementType.TEXT, texts);
            }
        } catch (FileNotFoundException err) {
            log.error(err.getMessage());
        }
    }

    /**
     * read illustration sheet element from com.java.eto.entity.excel, convert it into illustration object and add it into worksheet object
     * @param workbook workbook target(poi)
     * @param sheetName worksheet target (poi)
     * @param worksheet add illustration element into worksheet object(self-made)
     */
    private void readIllustration(XSSFWorkbook workbook, String sheetName, Worksheet worksheet){
        List<SheetElement> sheetElements = new ArrayList<>();
        try {
            XSSFSheet sheet = workbook.getSheet(sheetName);
            if(sheet == null) throw new FileNotFoundException("there is not sheet with name " + sheetName);
            XSSFDrawing chart = sheet.getDrawingPatriarch();

            if(chart != null) {
                //add to ontology
                List<XSSFShape> charts = chart.getShapes();
                for(int i = 0;i<charts.size();i++){
                    if(charts.get(i) instanceof Picture){
                        Picture picture = (Picture)charts.get(i);
                        Illustration illustration = new Illustration(worksheet, "Illustration");
                        sheetElements.add(illustration);
                    }
                }
                worksheet.addElement(ElementType.ILLUSTRATION, sheetElements);
            }
        } catch (FileNotFoundException err) {
            log.error(err.getMessage());
        }
    }

    /**
     * read chart sheet element from com.java.eto.entity.excel, convert it into chart object and add it into worksheet object
     * @param workbook workbook target(poi)
     * @param sheetName worksheet target (poi)
     * @param worksheet add chart element into worksheet object(self-made)
     */
    private void readChart(XSSFWorkbook workbook, String sheetName, Worksheet worksheet){
        List<SheetElement> sheetElements = new ArrayList<>();
        try {
            XSSFSheet sheet = workbook.getSheet(sheetName);
            if(sheet == null) throw new FileNotFoundException("there is not sheet with name " + sheetName);
            XSSFDrawing chart = sheet.getDrawingPatriarch();

            if(chart != null) {
                //add to ontology
                List<XSSFChart> charts = chart.getCharts();

                for(int i = 0; i<charts.size(); i++){
                    Chart chart1 = new Chart(worksheet, charts.get(i).getOrAddLegend().getEntries(),
                            charts.get(i).getTitleText());
                    sheetElements.add(chart1);
                }
                worksheet.addElement(ElementType.CHART, sheetElements);
            }
        } catch (FileNotFoundException err) {
            log.error(err.getMessage());
        }
    }

    /**
     * read basic element from worksheet(com.java.eto.entity.excel), convert it into basicElement object, and add it into worksheet object
     * @param sheet worksheet(com.java.eto.entity.excel + poi)
     * @param worksheet worksheet object (self-created)
     */
    private void readBasicElement(Sheet sheet, Worksheet worksheet) {
        List<SheetElement> cellTemp = new ArrayList<>();
        Map<String, Column> columnTemp = new HashMap<>();
        List<SheetElement> rowTemp = new ArrayList<>();
        List<String> colums = new ArrayList<>();
        List<String> rows = new ArrayList<>();
        boolean emptyCell = false;

        for(Row row : sheet){
            Entity.SheetElement.BasicElement.Row row1 = new Entity.SheetElement.BasicElement.Row(worksheet,
                    Integer.toString(row.getRowNum()));
            //log.info("creating row with id " + row1.getRowId());
            rows.add(""+row.getRowNum());
            for(Cell cell : row) {
                Entity.SheetElement.BasicElement.Cell cell1 = new Entity.SheetElement.BasicElement.Cell(worksheet,
                        convertColumn(Integer.toString(cell.getColumnIndex())),cell.getRowIndex());
                cell1.setRowID(Integer.toString(row.getRowNum()));
                //log.info("creating cell with id " + cell1.getCellId());
                switch (cell.getCellType()){
                    case STRING:
                        cell1.setValue(Value.STRING);
                        cell1.setStringValue(cell.getRichStringCellValue().getString());
                        break;
                    case BLANK:
                        emptyCell = true;
                        break;
                    case ERROR:
                        cell1.setValue(Value.ERROR);
                        cell1.SetErrorValue(String.valueOf(cell.getErrorCellValue()));
                        break;
                    case BOOLEAN:
                        cell1.setValue(Value.BOOLEAN);
                        cell1.setBooleanValue(cell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        cell1.setValue(Value.FORMULA);
                        FormulaConverter formulaConverter =  new FormulaConverter(cell.getCellFormula());
                        Formula formula = formulaConverter.getFormula();
                        switch (cell.getCachedFormulaResultType()){
                            case BOOLEAN:
                                formula.setBooleanValue(cell.getBooleanCellValue());
                                formula.setValueType(Value.BOOLEAN);
                                break;
                            case STRING:
                                formula.setStringValue(cell.getStringCellValue());
                                formula.setValueType(Value.STRING);
                                break;
                            case NUMERIC:
                                formula.setNumericValue((float)cell.getNumericCellValue());
                                formula.setValueType(Value.NUMERIC);
                                break;
                            case ERROR:
                                formula.setErrorValue(cell.getErrorCellValue());
                                formula.setValueType(Value.ERROR);
                        }
                        cell1.setFormulaValue(formula);
                        break;
                    case NUMERIC:
                        cell1.setValue(Value.NUMERIC);
                        cell1.setNumericValue((float)cell.getNumericCellValue());
                        break;
                }
                //log.info("cell got datatype " + cell1.getValue());

                XSSFComment comment = (XSSFComment)cell.getCellComment();
                if(comment!= null) cell1.setComment(comment);

                if(!columnTemp.containsKey(convertColumn(Integer.toString(cell.getColumnIndex())))){
                    Column column = new Column(worksheet, convertColumn(Integer.toString(cell.getColumnIndex())));
                    column.addCell(cell1);
                    columnTemp.put(convertColumn(Integer.toString(cell.getColumnIndex())), column);
                    colums.add(column.getColumnID());
                    //log.info("creating column with id " + column.getColumnID());
                }
                else{
                    Column column = columnTemp.get(convertColumn(Integer.toString(cell.getColumnIndex())));
                    column.addCell(cell1);
                }
                if(emptyCell) {
                    emptyCell = false;
                    continue;
                }
                cellTemp.add(cell1);
                row1.addCell(cell1);
            }
            rowTemp.add(row1);
        }

        worksheet.addElement(ElementType.ROW, rowTemp);
        worksheet.addElement(ElementType.CELL, cellTemp);
        //formulaCellDependencyCheck(cellTemp);
        List<SheetElement> colTemp = new ArrayList<>();
        for(Map.Entry<String, Column> temp : columnTemp.entrySet()) {
            colTemp.add(temp.getValue());
        }
        worksheet.addElement(ElementType.COLUMN, colTemp);
    }

    /**
     * add cell dependency to cell that have Formula value.
     * @param cell list of cell.
     */
    private void formulaCellDependencyCheck(List<SheetElement> cell){
        for(int i = 0; i<cell.size(); i++){
            Entity.SheetElement.BasicElement.Cell cell1 = (Entity.SheetElement.BasicElement.Cell)cell.get(i);
            if(cell1.getValue() == Value.FORMULA) {
                //log.info("initializing cell with formula value " + cell1.getFormulaValue());
                if(cell1.getFormulaValue().getFunctionType() == FunctionType.BASIC) {
                    addFormulaCellDependency(cell1.getFormulaValue().getFormulaFunction(), cell, cell1);
                } else if(cell1.getFormulaValue().getFunctionType() == FunctionType.NESTED) {
                    NestedFormula nestedFormula = (NestedFormula) cell1.getFormulaValue();
                    formulaDependencyCheck(nestedFormula.getFormulaList(), cell, cell1);
                }
            }
        }
    }

    /**
     * recursive function for nested formula
     * @param formulas list of formula in a nested formula
     * @param cell list of cell
     * @param cell1 cell which the list of formula from
     */
    private void formulaDependencyCheck(List<Formula> formulas, List<SheetElement> cell, Entity.SheetElement.BasicElement.Cell cell1) {
        for(int i = 0; i<formulas.size(); i++){
            if(formulas.get(i).getFunctionType() == FunctionType.BASIC) {
                addFormulaCellDependency(formulas.get(i).getFormulaFunction(), cell, cell1);
            } else if(formulas.get(i).getFunctionType() == FunctionType.NESTED) {
                NestedFormula nestedFormula = (NestedFormula) formulas.get(i);
                formulaDependencyCheck(nestedFormula.getFormulaList(),cell,cell1);
            }
        }
    }

    /**
     * save all the cellid in form of Cell:Cell in a string of list
     * example = C1 : C4 = C1,C2,C3,C4
     * @param text cell to cell that wanted to be converted
     * @return list of cell id
     */
    private List<String> cellToCell(String text) {
        List<String> result = new ArrayList<>();
        String[] temp = text.split(":");
        String fromColumn = temp[0].replaceAll("\\d","");
        String toColumn = temp[1].replaceAll("\\d", "");
        int fromRow = Integer.parseInt(temp[0].replaceAll("[a-zA-Z]*", ""));
        int tempRow = fromRow;
        int toRow = Integer.parseInt(temp[1].replaceAll("[a-zA-Z]*", ""));
        int Aascii = 65;
        result.add(fromColumn+fromRow);
        while(fromRow < toRow || !fromColumn.equals(toColumn)) {
            if(fromRow == toRow) {
                if(fromColumn.equals(toColumn)) {
                    break;
                }
                else {
                    fromRow = tempRow;
                    if(fromColumn.length() == 1) {
                        if(fromColumn.charAt(0) != 'Z') {
                            fromColumn = ""+(char)(fromColumn.charAt(0)+1);
                        }
                        else {
                            fromColumn = ""+(char)(Aascii) + (char)(Aascii);
                        }
                    }
                    else {
                        if(fromColumn.charAt(1) != 'Z') {
                            fromColumn = ""+ fromColumn.charAt(0)+(char)(fromColumn.charAt(1)+1);
                        }
                        else {
                            if(fromColumn.charAt(0) == 'Z') {
                                throw new IndexOutOfBoundsException("too many columns");
                            }
                            fromColumn = "" + (char)(fromColumn.charAt(0)+1) + (char)(Aascii);
                        }
                    }
                }
            }
            else {
                fromRow++;
            }
            result.add(fromColumn+fromRow);
        }
        return result;
    }

    /**
     * get if formula dependency, so it can be more specific.
     * @param formula formula function
     * @param cellList list of cells
     * @param cell cell to add formula dependency
     */
    private void ifFormulaDependency(String formula, List<SheetElement> cellList, Entity.SheetElement.BasicElement.Cell cell) {
        formula = formula.replaceAll("\\s*","");
        String regex = "\\(|\\)|\\+|,|-|\\*|/|;";
        String patternCell = "[a-zA-Z]+\\d+";
        String patternCellToCell = patternCell + ":" + patternCell;
        String patternCellFromOtherSheet = "'*[a-zA-Z]+\\d*'*![a-zA-Z]+\\d+\\s*";
        String patternCellToCellFromOtherSheet = patternCellFromOtherSheet + ":" + patternCellFromOtherSheet;
        String pattern1 = patternCell + "[>|<|>=|<=|==]{1}" +  patternCell;
        String pattern2 = patternCellFromOtherSheet + "[>|<|>=|<=|==]{1}" +  patternCell;
        String pattern3 = patternCellFromOtherSheet + "[>|<|>=|<=|==]{1}" + patternCellFromOtherSheet;
        String pattern4 = patternCell + "[>|<|>=|<=|==]{1}" + patternCellFromOtherSheet;
        String[] temp = formula.split(regex);
        boolean isTrue = false;
        int switcher = 0;
        for(int i = 0; i<temp.length; i++) {
            if(switcher == 0 && (temp[i].matches(patternCell) || temp[i].matches(patternCellFromOtherSheet) ||
                    temp[i].matches(patternCellToCellFromOtherSheet) || temp[i].matches(patternCellToCell))) {
                addFormulaCellDependency(temp[i], cellList, cell);
            }
            else if(switcher == 1 && (temp[i].matches(patternCell) || temp[i].matches(patternCellFromOtherSheet) ||
                    temp[i].matches(patternCellToCellFromOtherSheet) || temp[i].matches(patternCellToCell))) {
                if(isTrue == true) {
                    addFormulaCellDependency(temp[i], cellList, cell);
                    break;
                } else {
                    isTrue = true;
                }
            }
            if(temp[i].matches(pattern1) || temp[i].matches(pattern2) || temp[i].matches(pattern3) || temp[i].matches(pattern4)){
                String[] splitCell = temp[i].split("[>|<|>=|<=|==]{1}");
                Entity.SheetElement.BasicElement.Cell cell1 = (Entity.SheetElement.BasicElement.Cell)cellList.stream()
                        .filter(x -> splitCell[0].equals(x.id()))
                        .findAny()
                        .orElse(null);
                Entity.SheetElement.BasicElement.Cell cell2 = (Entity.SheetElement.BasicElement.Cell)cellList.stream()
                        .filter(x -> splitCell[1].equals(x.id()))
                        .findAny()
                        .orElse(null);
                if(cell1 != null && cell2 != null) {
                    if(cell1.getValue() != cell2.getValue()) {
                        return;
                    }
                    else {
                        if(temp[i].contains(">=")) {
                            isTrue = compare2Cell(cell1, cell2, ">=");
                        } else if(temp[i].contains(">")) {
                            isTrue = compare2Cell(cell1, cell2, ">");
                        } else if(temp[i].contains("<=")) {
                            isTrue = compare2Cell(cell1, cell2, "<=");
                        } else if(temp[i].contains("<")) {
                            isTrue = compare2Cell(cell1, cell2, "<");
                        } else if(temp[i].contains("==")) {
                            isTrue = compare2Cell(cell1, cell2, "==");
                        } else {
                            isTrue = compare2Cell(cell1, cell2, "<>");
                        }
                        switcher = 1;
                    }

                }
            }
        }
    }

    private boolean compare2Cell(Entity.SheetElement.BasicElement.Cell cell1, Entity.SheetElement.BasicElement.Cell cell2, String operator) {
        switch (cell1.getValue()) {
            case BOOLEAN:
                if(operator.equals("==")) {
                    if(cell1.isBooleanValue() == cell2.isBooleanValue()) return true;
                    else return false;
                }
                else {
                    if(cell1.isBooleanValue() != cell2.isBooleanValue()) return true;
                    else return false;
                }
            case NUMERIC :
                if(operator.equals("==")) {
                    if(cell1.getNumericValue() == cell2.getNumericValue()) return true;
                    else return false;
                } else if(operator.equals("<>")) {
                    if(cell1.getNumericValue() != cell2.getNumericValue()) return true;
                    else return false;
                } else if(operator.equals(">=")) {
                    if(cell1.getNumericValue() >= cell2.getNumericValue()) return true;
                    else return false;
                } else if(operator.equals(">")) {
                    if(cell1.getNumericValue() > cell2.getNumericValue()) return true;
                    else return false;
                } else if(operator.equals("<=")) {
                    if(cell1.getNumericValue() <= cell2.getNumericValue()) return true;
                    else return false;
                } else if(operator.equals("<")) {
                    if(cell1.getNumericValue() < cell2.getNumericValue()) return true;
                    else return false;
                }
            case STRING:
                if(operator.equals("==")) {
                    if(cell1.getStringValue() == cell2.getStringValue()) return true;
                    else return false;
                } else if(operator.equals("<>")) {
                    if(cell1.getStringValue() != cell2.getStringValue()) return true;
                    else return false;
                }
            case FORMULA:
                switch (cell1.getFormulaValue().getValueType()){
                    case BOOLEAN:
                        if(operator.equals("==")) {
                            if(cell1.getFormulaValue().getBooleanValue() == cell2.getFormulaValue().getBooleanValue()) return true;
                            else return false;
                        }
                        else {
                            if(cell1.getFormulaValue().getBooleanValue() != cell2.getFormulaValue().getBooleanValue()) return true;
                            else return false;
                        }
                    case NUMERIC :
                        if(operator.equals("==")) {
                            if(cell1.getFormulaValue().getNumericValue() == cell2.getFormulaValue().getNumericValue()) return true;
                            else return false;
                        } else if(operator.equals("<>")) {
                            if(cell1.getFormulaValue().getNumericValue() != cell2.getFormulaValue().getNumericValue()) return true;
                            else return false;
                        } else if(operator.equals(">=")) {
                            if(cell1.getFormulaValue().getNumericValue() >= cell2.getFormulaValue().getNumericValue()) return true;
                            else return false;
                        } else if(operator.equals(">")) {
                            if(cell1.getFormulaValue().getNumericValue() > cell2.getFormulaValue().getNumericValue()) return true;
                            else return false;
                        } else if(operator.equals("<=")) {
                            if(cell1.getFormulaValue().getNumericValue() <= cell2.getFormulaValue().getNumericValue()) return true;
                            else return false;
                        } else if(operator.equals("<")) {
                            if(cell1.getFormulaValue().getNumericValue() < cell2.getFormulaValue().getNumericValue()) return true;
                            else return false;
                        }
                    case STRING:
                        if(operator.equals("==")) {
                            if(cell1.getFormulaValue().getStringValue() == cell2.getFormulaValue().getStringValue()) return true;
                            else return false;
                        } else if(operator.equals("<>")) {
                            if(cell1.getFormulaValue().getStringValue() != cell2.getFormulaValue().getStringValue()) return true;
                            else return false;
                        }
                }
        }
        return false;
    }

    /**
     * use regex to detect Cell and Worksheet in a formula
     * @param formula formula function of a cell
     * @param cellList list of cell
     * @param cell cell that have a Formula value.
     */
    private void addFormulaCellDependency(String formula, List<SheetElement> cellList, Entity.SheetElement.BasicElement.Cell cell) {
        formula = formula.replaceAll("\\$", "");
        formula = formula.replaceAll("'", "");
        formula = formula.replaceAll(" ","");
        String patternCell = "[a-zA-Z]+\\d+";
        String patternCellToCell = patternCell + ":" + patternCell;
        String patternCellFromOtherSheet = "'*[a-zA-Z]+\\d*'*![a-zA-Z]+\\d+\\s*";
        String patternCellToCellFromOtherSheet = patternCellFromOtherSheet + ":" + patternCellFromOtherSheet;
        String patternCellToCellFromOtherSheet2 = patternCellFromOtherSheet + ":" + patternCell;
        String regex = "\\(|\\)|,|\\+|-|\\*|/|;";
        String[] temp = formula.split(regex);

        /*
        if(formula.contains("IF") || formula.contains("if")) {
            ifFormulaDependency(formula.replaceFirst("IF|if", ""), cellList, cell);
            return;
        }

         */
        for(int i = 0;i<temp.length; i++) {
            if(temp[i].matches(patternCell)){
                String check = temp[i];
                Entity.SheetElement.BasicElement.Cell cell1 = (Entity.SheetElement.BasicElement.Cell)cellList.stream()
                        .filter(cellCheck -> check.equals(cellCheck.id()))
                        .findAny()
                        .orElse(null);
                if(cell1 == null) {
                    log.info(cell.getWorksheet().getSheetName() + " matches cell pattern but cannot find cell with cell id " + check + " for formula " + formula);
                    continue;
                }
                else{
                    Formula basicFormula1 = cell.getFormulaValue();
                    basicFormula1.addDependencies(cell1);
                }
            }
            else if(temp[i].matches(patternCellFromOtherSheet)) {
                String[] worksheetCellSplit = temp[i].split("!");
                List<Worksheet> worksheets = this.workbook.getWorksheets();
                if(worksheetCellSplit[0].toLowerCase().equals("tabelle")) {
                    addFormulaCellDependency(worksheetCellSplit[1], cellList, cell);
                }
                else {
                    Worksheet worksheet = worksheets.stream()
                            .filter(worksheet1 -> worksheetCellSplit[0].equals(worksheet1.getSheetName()))
                            .findAny()
                            .orElse(null);

                    if (worksheet != null) {
                        List<SheetElement> cells = worksheet.getSheets().getOrDefault(ElementType.CELL, null);
                        if (cells != null) {
                            Entity.SheetElement.BasicElement.Cell cell1 = (Entity.SheetElement.BasicElement.Cell) cells.stream()
                                    .filter(cellTemp -> worksheetCellSplit[1].equals(cellTemp.id()))
                                    .findAny()
                                    .orElse(null);
                            if (cell1 != null) {
                                cell.getFormulaValue().addDependencies(cell1);
                            } else {
                                log.info(cell.getWorksheet().getSheetName() + " matches cell from another sheet pattern but " +
                                        "cannot find cell with cell id " + worksheetCellSplit[1] + " for formula " + formula);
                            }
                        }
                    } else {
                        log.info(cell.getWorksheet().getSheetName() + " matches cell from another sheet pattern but " +
                                "cannot find worksheet with name " + worksheetCellSplit[0] + " for formula " + formula);
                    }
                }
            }
            else if(temp[i].matches(patternCellToCell)) {
                List<String> cellToCells = cellToCell(temp[i]);
                for(int l = 0; l<cellToCells.size(); l++) {
                    String check1 = cellToCells.get(l);
                    Entity.SheetElement.BasicElement.Cell cell2 = (Entity.SheetElement.BasicElement.Cell)cellList.stream()
                            .filter(cellCheck -> check1.equals(cellCheck.id()))
                            .findAny()
                            .orElse(null);
                    if(cell2 == null) {
                        log.info(cell.getWorksheet().getSheetName() + " matches cell to cell pattern but cannot find cell with id " + check1 + " for formula " + formula);
                        continue;
                    }
                    else{
                        Formula basicFormula1 = cell.getFormulaValue();
                        basicFormula1.addDependencies(cell2);
                    }
                }
            }
            else if(temp[i].matches(patternCellToCellFromOtherSheet)) {
                String[] worksheetCellSplit = temp[i].split(":");
                String tempSplit = worksheetCellSplit[0].split("!")[1] + ":" + worksheetCellSplit[1].split("!")[1];
                List<Worksheet> worksheets = this.workbook.getWorksheets();
                Worksheet worksheet = worksheets.stream()
                        .filter(worksheet1 -> worksheetCellSplit[0].split("!")[0].equals(worksheet1.getSheetName()))
                        .findAny()
                        .orElse(null);

                if(worksheet!= null){
                    List<String> cell2 = cellToCell(tempSplit);
                    List<SheetElement> cells = worksheet.getSheets().getOrDefault(ElementType.CELL, null);
                    if(cells != null) {
                        for(int l = 0; l<cell2.size(); l++) {
                            String workSheetTemp = cell2.get(l);
                            Entity.SheetElement.BasicElement.Cell cell1 = (Entity.SheetElement.BasicElement.Cell) cells.stream()
                                    .filter(cellTemp -> workSheetTemp.equals(cellTemp.id()))
                                    .findAny()
                                    .orElse(null);
                            if (cell1 != null) {
                                cell.getFormulaValue().addDependencies(cell1);
                            } else {
                                log.info(cell.getWorksheet().getSheetName() + " matches cell to cell from another sheet pattern but " +
                                        "cannot find cell with id " + workSheetTemp + " for formula " + formula);
                            }
                        }
                    }
                } else {
                    log.info(cell.getWorksheet().getSheetName() + " matches cell to cell from another sheet pattern but " +
                            "cannot find worksheet with name " + worksheetCellSplit[0] + " for formula " + formula);
                }
            } else if(temp[i].matches(patternCellToCellFromOtherSheet2)) {
                String[] worksheetCellSplit = temp[i].split("!");
                List<Worksheet> worksheets = this.workbook.getWorksheets();
                Worksheet worksheet = worksheets.stream()
                        .filter(worksheet1 -> worksheetCellSplit[0].equals(worksheet1.getSheetName()))
                        .findAny()
                        .orElse(null);

                if(worksheet!= null){
                    List<String> cell2 = cellToCell(worksheetCellSplit[1]);
                    List<SheetElement> cells = worksheet.getSheets().getOrDefault(ElementType.CELL, null);
                    if(cells != null) {
                        for(int l = 0; l<cell2.size(); l++) {
                            String workSheetTemp = cell2.get(l);
                            Entity.SheetElement.BasicElement.Cell cell1 = (Entity.SheetElement.BasicElement.Cell) cells.stream()
                                    .filter(cellTemp -> workSheetTemp.equals(cellTemp.id()))
                                    .findAny()
                                    .orElse(null);
                            if (cell1 != null) {
                                cell.getFormulaValue().addDependencies(cell1);
                            } else {
                                log.info(cell.getWorksheet().getSheetName() + " matches cell to cell from another sheet second pattern but " +
                                        "cannot find with id " + workSheetTemp + " for formula " + formula);
                            }
                        }
                    }
                } else {
                    log.info(cell.getWorksheet().getSheetName() + " matches cell to cell from another sheet second pattern but " +
                            "cannot find worksheet with name " + worksheetCellSplit[0] + " for formula " + formula);
                }
            }
        }
    }

    /**
     * convert column index from poi into actual index in com.java.eto.entity.excel
     * @param column colum index from poi
     * @return actual com.java.eto.entity.excel index
     */
    private String convertColumn(String column) {
        String result = "";
        int min = 65;
        int temp = Integer.parseInt(column);
        boolean check = false;
        while(temp >= 0) {
            if(temp <= 25){
                check = true;
            }
            int calc = temp % 26;
            temp = temp/26;
            if(temp >=1 && temp <=25) {
                char character = (char) (min + (temp-1));
                result = result + character;
                character = (char) (min + calc);
                result = result + character;
                break;
            } else if(temp > 25) {
                char character = (char)(min + ((temp/26) - 1));
                result = result + character;
            } else {
                char character = (char) (min + calc);
                result = result + character;
            }
            if(check){
                break;
            }
        }
        return result;
    }

    /**
     * read table sheet element from com.java.eto.entity.excel, convert it into table object and add it into worksheet object
     * @param workbook workbook target(poi)
     * @param sheetName worksheet target (poi)
     * @param worksheet add table element into worksheet object(self-created)
     */
    private void readTable(XSSFWorkbook workbook, String sheetName, Worksheet worksheet) {
        List<SheetElement> sheetElements = new ArrayList<>();
        try{
            XSSFSheet sheet = workbook.getSheet(sheetName);
            if(sheet == null) throw new FileNotFoundException("there is not sheet with name " + sheetName);
            List<XSSFTable> tables = sheet.getTables();
            List<XSSFPivotTable> pivotTables = sheet.getPivotTables();

            for(int i = 0; i<tables.size(); i++) {
                XSSFTable t = tables.get(i);
                String colStart = convertColumn(Integer.toString(t.getStartColIndex()));
                String colEnd = convertColumn(Integer.toString(t.getEndColIndex()));
                String rowStart = Integer.toString(t.getStartRowIndex()+1);
                String rowEnd = Integer.toString(t.getEndRowIndex()+1);
                Table table = new Table(worksheet, colEnd, colStart, rowStart, rowEnd, t.getName());
                List<SheetElement> cell = worksheet.getSheets().get(ElementType.CELL);
                List<String> temp= cellToCell(colStart + rowStart + ":" + colEnd + rowEnd);
                for(int l = 0; l<temp.size(); l++) {
                    String str = temp.get(l);
                    Entity.SheetElement.BasicElement.Cell cell1 = (Entity.SheetElement.BasicElement.Cell)cell.stream()
                            .filter(x -> str.equals(x.id()))
                            .findAny()
                            .orElse(null);
                    if(cell1 != null) {
                        table.addCell(cell1);
                    }
                }
                sheetElements.add(table);
            }
            worksheet.addElement(ElementType.TABLE, sheetElements);

        } catch (FileNotFoundException err) {
            log.error(err.getMessage());
        }

    }

    @Override
    public Workbook getWorkbook() {
        return workbook;
    }

}
