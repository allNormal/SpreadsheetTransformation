package Transformer;

import Entity.Operator;
import Entity.SheetElement.BasicElement.Cell;
import Entity.SheetElement.Charts.Chart;
import Entity.SheetElement.ElementType;
import Entity.SheetElement.SheetElement;
import Entity.SheetElement.Tables.Table;
import Entity.SheetElement.Texts.Text;
import Entity.Workbook.Workbook;
import Entity.Worksheet.Worksheet;
import ExcelReader.ExcelReader;
import org.apache.jena.ontology.Individual;
import org.apache.jena.ontology.OntModel;
import org.apache.jena.query.*;
import org.apache.jena.vocabulary.RDFS;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import Exception.*;
import ExcelReader.*;
import java.util.List;
import java.util.Map;

public class ExcelTransformer implements ITransformer {

    private ExcelModel excelModel;
    private Workbook workbook;
    private OntModel model;
    private ExcelBasedReader excelBasedReader;

    public ExcelTransformer(ExcelBasedReader excelBasedReader){
        this.excelBasedReader = excelBasedReader;
    }


    @Override
    public void create(){
        if(excelBasedReader.getWorkbook() == null) {
            throw new NullPointerException("workbook is null");
        }
        this.workbook = excelBasedReader.getWorkbook();
        this.excelModel = new ExcelModel(workbook.getFileName());
        this.model = this.excelModel.getModel();
        convertExcelToOntology();
    }


    @Override
    public List<String> getCellDependencies(String cellID, String worksheetName) {
        String queryString = "PREFIX owl: <http://www.w3.org/2002/07/owl#>\n" +
                "PREFIX xsd: <http://www.w3.org/2001/XMLSchema#>\n" +
                "PREFIX rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#>\n" +
                "PREFIX : <http://www.semanticweb.org/GregorKaefer/ontologies/2020/8/excelOntology#>\n" +
                "PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>\n" +
                "PREFIX ns: <http://xmlns.com/foaf/0.1/>\n" +
                "\n" +
                "SELECT ?cellDependency\n" +
                "WHERE {\n" +
                "    ?worksheet rdf:type :Worksheet.\n" +
                "    ?worksheet :SheetName ?worksheetName.\n" +
                "    ?cell rdf:type :cell.\n" +
                "    ?cell :CellID ?cellId.\n" +
                "    ?worksheet :hasSheetElement ?cell.\n" +
                "    ?cell (:hasValue/:hasCell)+ ?cellDependency.\n" +
                "    Filter(?cellId = '" +cellID + "').\n" +
                "    Filter(?worksheetName = '" + worksheetName + "').\n" +
                "}";

        Query query = QueryFactory.create(queryString);
        List<String> dependencyList = new ArrayList<>();

        QueryExecution qExec = QueryExecutionFactory.create(query, this.model);
        try {
            ResultSet result = qExec.execSelect();
            while (result.hasNext()) {
                QuerySolution solt = result.nextSolution();
                dependencyList.add(solt.getResource("cellDependency").getProperty(this.excelModel.getCellID()).getString());
            }
        } finally {
            qExec.close();
        }

        return dependencyList;
    }

    @Override
    public List<String> addConstraint(ElementType type, String typeID, String worksheetName, Operator operator, String value) {

        List<String> result = new ArrayList<>();
        float valueNumeric = 0;
        boolean isNumeric = true;
        try {
            valueNumeric = Float.parseFloat(value);
        } catch (NumberFormatException e) {
            isNumeric = false;
        }

        String queryString = "PREFIX owl: <http://www.w3.org/2002/07/owl#>\n" +
                "PREFIX xsd: <http://www.w3.org/2001/XMLSchema#>\n" +
                "PREFIX rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#>\n" +
                "PREFIX : <http://www.semanticweb.org/GregorKaefer/ontologies/2020/8/excelOntology#>\n" +
                "PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>\n" +
                "PREFIX ns: <http://xmlns.com/foaf/0.1/>\n" +
                "\n" +
                "SELECT ?cellId\n" +
                "WHERE {\n" +
                "    ?worksheet rdf:type :Worksheet.\n" +
                "    ?worksheet :SheetName ?worksheetName.\n" +
                "    ?cell rdf:type :cell.\n" +
                "    ?cell :CellID ?cellId.\n" +
                "    ?worksheet :hasSheetElement ?cell.\n" +
                "    ?column rdf:type :column.\n" +
                "    ?column :ColumnID ?columnID.\n" +
                "    ?row rdf:type :row.\n" +
                "    ?row :RowID ?rowID.\n" +
                "    ?cell :hasRow ?row.\n" +
                "    ?cell :hasColumn ?column.\n" +
                "    ?cell :hasValue ?valueType.\n" +
                "    ?valueType :CellValue ?value.\n";

        if(!isNumeric) value = "'" + value + "'";

        queryString += "Filter (?value "+ operator.getOperator() + " " + value + "). \n";

        if(type == ElementType.CELL) {
            queryString +=
                    "Filter (?cellId = '" +typeID + "').\n" +
                            "Filter (?worksheetName = '" + worksheetName +"').\n" +
                            "}\n" +
                            "ORDER BY ASC(?cellId)";

        } else if(type == ElementType.COLUMN) {

            queryString +=
                    "Filter (?columnID = '" +typeID + "').\n" +
                            "Filter (?worksheetName = '" + worksheetName +"').\n" +
                            "}\n" +
                            "ORDER BY ASC(?cellId)";


        } else if(type == ElementType.ROW) {

            queryString +=
                    "Filter(?rowID = '" +typeID + "').\n" +
                            "Filter(?worksheetName = '" + worksheetName +"').\n" +
                            "}\n" +
                            "ORDER BY ASC(?cellId)";

        }

        Query query = QueryFactory.create(queryString);

        QueryExecution qExec = QueryExecutionFactory.create(query, this.model);
        try {
            ResultSet resultSet = qExec.execSelect();
            while (resultSet.hasNext()) {
                QuerySolution solt = resultSet.nextSolution();
                result.add(solt.get("cellId").toString());
            }
        } finally {
            qExec.close();
        }

        return result;
    }

    @Override
    public List<String> getReverseDependencies(String cellID, String worksheetName) {

        String queryString = "PREFIX owl: <http://www.w3.org/2002/07/owl#>\n" +
                "PREFIX xsd: <http://www.w3.org/2001/XMLSchema#>\n" +
                "PREFIX rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#>\n" +
                "PREFIX : <http://www.semanticweb.org/GregorKaefer/ontologies/2020/8/excelOntology#>\n" +
                "PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>\n" +
                "PREFIX ns: <http://xmlns.com/foaf/0.1/>\n" +
                "\n" +
                "SELECT DISTINCT ?cellDependencyIsUsedIn\n" +
                "WHERE {\n" +
                "    ?worksheet rdf:type :Worksheet.\n" +
                "    ?worksheet :SheetName ?worksheetName.\n" +
                "    ?cell rdf:type :cell.\n" +
                "    ?cell :CellID ?cellId.\n" +
                "    ?worksheet :hasSheetElement ?cell.\n" +
                "    ?cell :isUsedIn ?valueType.\n" +
                "    ?cellDependencyIsUsedIn rdf:type :cell.\n" +
                "    ?cellDependencyIsUsedIn :hasValue ?value.\n" +
                "    ?valueType :FunctionName ?function1.\n" +
                "    ?value :FunctionName ?function2.\n" +
                "    Filter(?cellId = '" + cellID + "').\n" +
                "    Filter(?worksheetName = '" + worksheetName + "').\n" +
                "    Filter(?function1 = ?function2).\n" +
                "}\n";

        Query query = QueryFactory.create(queryString);
        List<String> dependencyListIsUsedIn = new ArrayList<>();

        QueryExecution qExec = QueryExecutionFactory.create(query, this.model);
        try {
            ResultSet result = qExec.execSelect();
            while (result.hasNext()) {
                QuerySolution solt = result.nextSolution();
                dependencyListIsUsedIn.add(solt.getResource("cellDependencyIsUsedIn").getProperty(this.excelModel.getCellID()).getString());
            }
        } finally {
            qExec.close();
        }

        return dependencyListIsUsedIn;
    }

    @Override
    public OntModel getModel() {
        return model;
    }

    /**
     * add com.java.eto.entity.excel data into the rdf-graph
     */
    private void convertExcelToOntology(){
        //System.out.println("converting 1 workbook");
        Individual workbook = addWorkbook();
        //System.out.println("converting " + this.workbook.getWorksheets().size() + " worksheet");
        for(int i = 0;i<this.workbook.getWorksheets().size();i++){
            //  System.out.println("worksheet " + (i+1));
            Individual worksheet = addWorksheet(this.workbook.getWorksheets().get(i));
            workbook.addProperty(this.excelModel.getHasWorksheet(), worksheet);
            worksheet.addProperty(this.excelModel.getIsInWorkbook(), workbook);

            List<SheetElement> chart = new ArrayList<>();
            List<SheetElement> illustration = new ArrayList<>();
            List<SheetElement> table =  new ArrayList<>();
            List<SheetElement> text = new ArrayList<>();
            List<SheetElement> cell = new ArrayList<>();
            List<SheetElement> row = new ArrayList<>();
            List<SheetElement> column = new ArrayList<>();

            //initialize list with the saved data in map.
            for(Map.Entry<ElementType, List<SheetElement>> m : this.workbook.getWorksheets().get(i).getSheets().entrySet()) {
                switch (m.getKey()){
                    case TABLE:
                        table = m.getValue();
                        break;
                    case COLUMN:
                        column = m.getValue();
                        break;
                    case CELL:
                        cell = m.getValue();
                        break;
                    case ROW:
                        row = m.getValue();
                        break;
                    case CHART:
                        chart = m.getValue();
                        break;
                    case ILLUSTRATION:
                        illustration = m.getValue();
                        break;
                    case TEXT:
                        text = m.getValue();
                        break;
                    default: break;
                }
            }

            //check if worksheet contain basic + sheet element or not
            if(chart.size() != 0 || illustration.size() != 0 || table.size() != 0 || cell.size() != 0 ||
                    row.size() != 0 || column.size() != 0 || text.size() != 0) {

                Individual sheetElement = this.excelModel.getSheetElement().createIndividual();
                worksheet.addProperty(this.excelModel.getHasSheetElement(), sheetElement);
                sheetElement.addProperty(this.excelModel.getIsPartOfWorksheet(), worksheet);
                addBasicElement(cell, worksheet, this.workbook.getWorksheets().get(i));
                addChart(chart, worksheet, this.workbook.getWorksheets().get(i));
                addTable(table, worksheet, this.workbook.getWorksheets().get(i));
                addText(text, worksheet, this.workbook.getWorksheets().get(i));
                addIllustration(illustration, worksheet, this.workbook.getWorksheets().get(i));
            }
        }
    }


    /**
     * save model into file with turtle format
     * @param fileName file name that want to be converted into file
     */
    @Override
    public void saveModel(String path, String fileName) {
        FileOutputStream out = null;
        try {
            Path path1 = Paths.get(path + fileName + ".ttl");
            out = new FileOutputStream(path1.toString());
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        this.excelModel.getModel().write(out, "TTL");
    }

    /**
     * add workbook data into the rdf-graph
     * @return workbook instance.
     */
    private Individual addWorkbook() {
        Individual workbook = this.excelModel.getWorkbook().createIndividual(this.excelModel.getWorkbook().getURI() + "_" +
                this.workbook.getFileName());
        workbook.addLiteral(RDFS.label, this.workbook.getFileName());
        workbook.addLiteral(this.excelModel.getExtension(), this.workbook.getExtension());
        workbook.addLiteral(this.excelModel.getFileName(), this.workbook.getFileName());
        if(this.workbook.getMacro()!=null) {
            Individual macro = this.excelModel.getMacro().createIndividual(this.excelModel.getMacro().getURI() + "_" +
                    "Macro");
            macro.addLiteral(this.excelModel.getMacroInput(), this.workbook.getMacro().getMacro());
            workbook.addProperty(this.excelModel.getHasMacro(), macro);
        }
        return workbook;
    }

    /**
     * add worksheet data into the rdf-graph
     * @param worksheet for naming purposes(not really important can be removed i think)
     * @return worksheet instance.
     */
    private Individual addWorksheet(Worksheet worksheet){
        Individual i = this.excelModel.getWorksheet().createIndividual(this.excelModel.getWorksheet().getURI() + "_" +
                worksheet.getSheetName());
        i.addLiteral(this.excelModel.getSheetName(), worksheet.getSheetName());
        i.addLiteral(RDFS.label, worksheet.getSheetName());
        return i;
    }

    /**
     * Iterate through cell list and call addCell O(cell.size)
     * @param cell list of cell in the worksheet that already been converted from com.java.eto.entity.excel
     * @param worksheet for isPartOfWorksheet relation so data can be tracked coming from which worksheet
     * @param ws for naming purposes(not really important can be removed i think)
     */
    private void addBasicElement(List<SheetElement> cell, Individual worksheet, Worksheet ws) {
        //System.out.println(cell.size());
        for(int i = 0;i<cell.size();i++) {
            if(cell.get(i) instanceof Cell) {
                addCell((Cell)cell.get(i), worksheet, ws);
            }
        }

        //System.out.println("converting " + cell.size() + " cells");
    }

    /**
     * add basicelement(Row, Column, Cell) into the rdf-graph
     * @param cell cell data that already been converted from com.java.eto.entity.excel and
     *            since row, column also saved in cell, we only need cell parameter.
     * @param worksheet for the relation isPartOfWorksheet (so data can be tracked coming from which worksheet)
     * @param ws for naming purposes(not really important can be removed i think)
     */
    private void addCell(Cell cell, Individual worksheet, Worksheet ws) {
        Individual i = this.excelModel.getCell().createIndividual(this.excelModel.getCell().getURI() +
                "_Worksheet" + ws.getSheetName() +
                "_" + cell.getCellId());
        i.addLiteral(RDFS.label, "Cell_"+cell.getCellId());
        i.addLiteral(this.excelModel.getCellID(), cell.getCellId());
        i.addProperty(this.excelModel.getIsPartOfWorksheet(), worksheet);
        Individual val;
        switch (cell.getValue()){
            case STRING:
                val = this.excelModel.getStringValue().createIndividual(i.getURI() + "_value" );
                val.addLiteral(RDFS.label, "StringValue");
                val.addLiteral(this.excelModel.getCellValue(), cell.getStringValue());
                i.addProperty(this.excelModel.getHasValue(),val);
                break;
            case NUMERIC:
                val = this.excelModel.getNumericValue().createIndividual(i.getURI() + "_value");
                val.addLiteral(RDFS.label, "NumericValue");
                val.addLiteral(this.excelModel.getCellValue(), cell.getNumericValue());
                i.addProperty(this.excelModel.getHasValue(),val);
                break;
            case FORMULA:
                val = this.excelModel.getFormulaValue().createIndividual(i.getURI() + "_value");
                val.addLiteral(RDFS.label, "FormulaValue");
                val.addLiteral(this.excelModel.getFunctionName(), cell.getFormulaValue().getFormulaFunction());
                /**
                 * add value into formula
                 */
                switch(cell.getFormulaValue().getValueType()) {
                    case BOOLEAN:
                        val.addLiteral(this.excelModel.getCellValue(), cell.getFormulaValue().getBooleanValue());
                        break;
                    case STRING:
                        val.addLiteral(this.excelModel.getCellValue(), cell.getFormulaValue().getStringValue());
                        break;
                    case NUMERIC:
                        val.addLiteral(this.excelModel.getCellValue(), cell.getFormulaValue().getNumericValue());
                        break;
                    case ERROR:
                        val.addLiteral(this.excelModel.getCellValue(), "ERROR");
                        break;
                }
                for(int k = 0; k<cell.getFormulaValue().getCellDependency().size(); k++) {
                    Individual temp = this.excelModel.getCell().createIndividual(this.excelModel.getCell().getURI() +
                            "_Worksheet" + cell.getFormulaValue().getCellDependency().get(k).getWorksheet().getSheetName() +
                            "_" + cell.getFormulaValue().getCellDependency().get(k).getCellId());
                    val.addProperty(this.excelModel.getHasCell(), temp);
                    temp.addProperty(this.excelModel.getIsUsedIn(), val);
                }
                i.addProperty(this.excelModel.getHasValue(),val);
                break;
            case BOOLEAN:
                val = this.excelModel.getBooleanValue().createIndividual(i.getURI() + "_value");
                val.addLiteral(RDFS.label, "BooleanValue");
                val.addLiteral(this.excelModel.getCellValue(), cell.isBooleanValue());
                i.addProperty(this.excelModel.getHasValue(),val);
                break;
            case ERROR:
                val = this.excelModel.getErrorValue().createIndividual(i.getURI() + "_value");
                val.addLiteral(RDFS.label, "ErrorValue");
                val.addLiteral(this.excelModel.getErrorName(), cell.getError().getErrorName());
                i.addProperty(this.excelModel.getHasValue(),val);
                break;
        }

        if(cell.getComment() != null) {
            Individual comment = this.excelModel.getComment().createIndividual(i.getURI() + "comment");
            comment.addLiteral(this.excelModel.getCommentsValue(), cell.getComment().getString().toString());
            comment.addLiteral(RDFS.label, cell.getCellId() + "-Comment");
            i.addProperty(this.excelModel.getHasComment(), comment);
        }

        Individual row = this.excelModel.getRow().createIndividual( this.excelModel.getRow().getURI() +
                "_Worksheet" + ws.getSheetName() +
                "_" + cell.getRow());
        row.addLiteral(RDFS.label, "Row_"+cell.getRow());
        row.addLiteral(this.excelModel.getRowID(), cell.getRow());
        row.addProperty(this.excelModel.getHasCell(), i);
        row.addProperty(this.excelModel.getIsPartOfWorksheet(), worksheet);
        i.addProperty(this.excelModel.getHasRow(),row);

        Individual column = this.excelModel.getColumn().createIndividual(this.excelModel.getColumn().getURI() +
                "_Worksheet" + ws.getSheetName() +
                "_" + cell.getColumn());
        column.addLiteral(RDFS.label, "Column_"+cell.getColumn());
        column.addLiteral(this.excelModel.getColumnID(),cell.getColumn());
        column.addProperty(this.excelModel.getHasCell(),i);
        column.addProperty(this.excelModel.getHasRow(), row);
        column.addProperty(this.excelModel.getIsPartOfWorksheet(), worksheet);
        i.addProperty(this.excelModel.getHasColumn(),column);
        row.addProperty(this.excelModel.getHasColumn(), column);
        worksheet.addProperty(this.excelModel.getHasSheetElement(), i);
        worksheet.addProperty(this.excelModel.getHasSheetElement(), column);
        worksheet.addProperty(this.excelModel.getHasSheetElement(), row);
    }

    /**
     * add chart data into the rdf-graph O(chart.size)
     * @param chart
     * @param worksheet
     * @param ws for naming purposes(not really important can be removed i think)
     */
    private void addChart(List<SheetElement> chart, Individual worksheet, Worksheet ws) {
        for(int i = 0;i<chart.size(); i++) {
            if(chart.get(i) instanceof Chart) {
                if(((Chart) chart.get(i)).getTitle() == null) {
                    Individual j = this.excelModel.getChart().createIndividual(this.excelModel.getChart().getURI() +
                            "_Worksheet" + ws.getSheetName() +
                            "_" + "Chart" + (i+1));
                    j.addProperty(this.excelModel.getIsPartOfWorksheet(), worksheet);
                    j.addLiteral(RDFS.label, "Chart" + i+1);
                }
                else if(((Chart) chart.get(i)).getTitle().toString().isEmpty()) {
                    Individual j = this.excelModel.getChart().createIndividual(this.excelModel.getChart().getURI() +
                            "_Worksheet" + ws.getSheetName() +
                            "_" + "Chart" + (i+1));
                    j.addProperty(this.excelModel.getIsPartOfWorksheet(), worksheet);
                    j.addLiteral(RDFS.label, "Chart" + i+1);
                }
                else {
                    Individual j = this.excelModel.getChart().createIndividual(this.excelModel.getChart().getURI() +
                            "_Worksheet" + ws.getSheetName() +
                            "_" + ((Chart) chart.get(i)).getTitle().toString());
                    j.addLiteral(this.excelModel.getChartTitle(), ((Chart) chart.get(i)).getTitle().toString());
                    j.addProperty(this.excelModel.getIsPartOfWorksheet(), worksheet);
                    j.addLiteral(RDFS.label, ((Chart) chart.get(i)).getTitle().toString());
                }
            }
        }
        //System.out.println("converting " + chart.size() + " chart");
    }

    /**
     * add table data into the rdf-graph O(table.size)
     * @param table table data that already been converted from com.java.eto.entity.excel
     * @param worksheet for relation isPartOfWorksheet (so data can be tracked from which worksheet)
     * @param ws for naming purposes(not really important can be removed i think)
     */
    private void addTable(List<SheetElement> table, Individual worksheet, Worksheet ws) {
        for(int j = 0;j<table.size();j++){
            if(table.get(j) instanceof Table) {
                if(table.get(j).id() != null) {
                    addTableRelation((Table)table.get(j), ws, table.get(j).id(), worksheet);
                }
                else{
                    addTableRelation((Table)table.get(j), ws, "Table" + (j+1), worksheet);
                }
            }
        }

        //System.out.println("converting " + table.size() + " table");

    }

    private void addTableRelation (Table table, Worksheet ws, String title, Individual worksheet){
        Individual Table = this.excelModel.getTable().createIndividual(this.excelModel.getTable().getURI() +
                "_Worksheet" + ws.getSheetName() +
                "_" + title);
        Table.addLiteral(RDFS.label, title);
        Table.addLiteral(this.excelModel.getElementName(), title);
        Table.addLiteral(this.excelModel.getColStart(),  table.getColumnStart());
        Table.addLiteral(this.excelModel.getColEnd(),  table.getColumnEnd());
        Table.addLiteral(this.excelModel.getRowEnd(),  table.getRowEnd());
        Table.addLiteral(this.excelModel.getRowStart(), table.getRowStart());
        Table.addProperty(this.excelModel.getIsPartOfWorksheet(), worksheet);
        for(int i = 0;i<table.getCell().size();i++) {
            Individual cell = this.excelModel.getModel().getIndividual(this.excelModel.getCell().getURI() +
                    "_Worksheet" + ws.getSheetName() +
                    "_" + table.getCell().get(i).getCellId());
            Table.addProperty(this.excelModel.getHasCell(), cell);
            worksheet.removeProperty(this.excelModel.getHasSheetElement(),cell);
            cell.removeProperty(this.excelModel.getIsPartOfWorksheet(), worksheet);
        }
    }

    /**
     * add Text - Sheet element instances into rdf-graph O(text.size)
     * @param text text data that already been converted from com.java.eto.entity.excel
     * @param worksheet worksheet to give the text class relation isPartOfWorksheet (so data can be tracked from which worksheet)
     * @param ws for naming purposes(not really important can be removed i think)
     */
    private void addText(List<SheetElement> text, Individual worksheet, Worksheet ws) {
        for(int j = 0;j<text.size();j++) {
            if(text.get(j) instanceof Text) {
                Individual i = this.excelModel.getText().createIndividual(this.excelModel.getText().getURI() +
                        "_Worksheet" + ws.getSheetName() +
                        "_" + text.get(j).id() + (j+1));
                i.addLiteral(RDFS.label, text.get(j).id() + (j+1));
                i.addLiteral(this.excelModel.getHasValue(), ((Text) text.get(j)).getValue());
                i.addProperty(this.excelModel.getIsPartOfWorksheet(), worksheet);
            }
        }
        //System.out.println("converting " + text.size() + " text");
    }

    /**
     * add Illustrations(Pictures,Icons,etc) element into rdf-graph O(illustrations.size)
     * @param illustrations illustrations data that already been converted from com.java.eto.entity.excel.
     * @param worksheet to link isPartOfWorksheet with illustrations (so we can track which data is from which worksheet)
     * @param ws for naming purposes(not really important can be removed i think)
     */
    private void addIllustration(List<SheetElement> illustrations, Individual worksheet, Worksheet ws) {
        for(int j = 0;j<illustrations.size();j++) {
            Individual i = this.excelModel.getIllustration().createIndividual(this.excelModel.getIllustration() +
                    "_Worksheet" + ws.getSheetName() +
                    "_" + illustrations.get(j).id() + (j+1));
            i.addLiteral(RDFS.label, illustrations.get(j).id() + (j+1));
            i.addProperty(this.excelModel.getIsPartOfWorksheet(), worksheet);
        }
        //System.out.println("converting " + illustrations.size() +  " illustrations");
    }


}
