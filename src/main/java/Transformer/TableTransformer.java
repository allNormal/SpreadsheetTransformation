package Transformer;

import Entity.Operator;
import Entity.SheetElement.BasicElement.Cell;
import Entity.SheetElement.BasicElement.Column;
import Entity.SheetElement.BasicElement.Row;
import Entity.SheetElement.ElementType;
import Entity.SheetElement.SheetElement;
import Entity.SheetElement.Tables.Table;
import Entity.Workbook.Workbook;
import Entity.Worksheet.Worksheet;
import org.apache.jena.ontology.Individual;
import org.apache.jena.ontology.OntModel;
import org.apache.jena.query.*;
import org.apache.jena.vocabulary.RDFS;

import java.io.FileOutputStream;
import java.io.IOException;
import Exception.IncorrectTypeException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

public class TableTransformer implements ITransformer{
    private TableModel tableModel;
    private OntModel model;
    private Workbook workbook;

    public TableTransformer() {
    }

    @Override
    public void create(Workbook workbook) throws IncorrectTypeException {
        this.workbook = workbook;
        this.tableModel = new TableModel(this.workbook.getFileName());
        this.model = this.tableModel.getModel();
        convertExcelToOntology();
    }


    @Override
    public Collection<String> getCellDependencies(String cellID, String worksheetName) {
        String queryString = "PREFIX owl: <http://www.w3.org/2002/07/owl#>\n" +
                "PREFIX xsd: <http://www.w3.org/2001/XMLSchema#>\n" +
                "PREFIX rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#>\n" +
                "PREFIX : <http://www.semanticweb.org/GregorKaefer/ontologies/2020/11/excelOntologyVersion2#>\n" +
                "PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>\n" +
                "PREFIX ns: <http://xmlns.com/foaf/0.1/>\n" +
                "\n" +
                "SELECT ?columnDependency\n" +
                "WHERE {\n" +
                "    ?worksheet rdf:type :Worksheet.\n" +
                "    ?worksheet :SheetName ?worksheetName.\n" +
                "    ?table rdf:type :Table.\n" +
                "    ?table :TableName ?tableName.\n" +
                "    ?worksheet :hasTable ?table.\n" +
                "    ?column rdf:type :ColumnHeader.\n" +
                "    ?column :ColumnHeaderID ?columnId.\n" +
                "    ?table :hasColumnHeader ?column.\n" +
                "    ?column (:hasValueType/:hasDependency)+ ?columnDependency.\n" +
                "    Filter(?columnId = '" + cellID.replaceAll("\\d","") + "').\n" +
                "    Filter(?worksheetName = '" + worksheetName + "').\n" +
                "}\n";

        Query query = QueryFactory.create(queryString);
        List<String> dependencyList = new ArrayList<>();

        QueryExecution qExec = QueryExecutionFactory.create(query, this.model);
        try {
            ResultSet result = qExec.execSelect();
            while (result.hasNext()) {
                QuerySolution solt = result.nextSolution();
                dependencyList.add(solt.getResource("columnDependency").getLocalName());
            }
        } finally {
            qExec.close();
        }

        return dependencyList;
    }

    @Override
    public Collection<String> addConstraint(ElementType type, String typeID, String worksheetName, Operator operator, String value) {
        List<String> result = new ArrayList<>();
        float valueNumeric = 0;
        boolean isNumeric = true;
        try {
            valueNumeric = Float.parseFloat(value);
        } catch (NumberFormatException e) {
            isNumeric = false;
        }

        String search = "";

        String queryString = "PREFIX owl: <http://www.w3.org/2002/07/owl#>\n" +
                "PREFIX xsd: <http://www.w3.org/2001/XMLSchema#>\n" +
                "PREFIX rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#>\n" +
                "PREFIX : <http://www.semanticweb.org/GregorKaefer/ontologies/2020/11/excelOntologyVersion2#>\n" +
                "PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>\n" +
                "PREFIX ns: <http://xmlns.com/foaf/0.1/>\n" +
                "\n" +
                "SELECT ?rowTitle\n" +
                "WHERE {\n" +
                "    ?worksheet rdf:type :Worksheet.\n" +
                "    ?worksheet :SheetName ?worksheetName.\n" +
                "    ?table rdf:type :Table.\n" +
                "    ?worksheet :hasTable ?table.\n" +
                "\t?column rdf:type :ColumnHeader.\n" +
                "    ?column :ColumnHeaderID ?columnId.\n" +
                "    ?table :hasColumnHeader ?column.\n" +
                "    ?row rdf:type :RowHeader.\n" +
                "    ?row :RowHeaderTitle ?rowTitle.\n" +
                "    ?table :hasRowHeader ?row.\n" +
                "    ?value rdf:type :Value.\n" +
                "    ?value :ValueFor ?column.\n" +
                "    ?value :ActualValue ?valueConstraint.\n" +
                "    ?row :hasValue ?value.\n";

        if(!isNumeric) value = "'" + value + "'";

        queryString += "Filter (?value "+ operator.getOperator() + " " + value + "). \n";

        if(type == ElementType.COLUMN) {
            queryString +=
                    "Filter (?columnId = '" +typeID.replaceAll("\\d","") + "').\n" +
                            "Filter (?worksheetName = '" + worksheetName +"').\n" +
                            "}\n";
            search = "rowTitle";

        } else if(type == ElementType.ROW) {

            queryString = "PREFIX owl: <http://www.w3.org/2002/07/owl#>\n" +
                    "PREFIX xsd: <http://www.w3.org/2001/XMLSchema#>\n" +
                    "PREFIX rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#>\n" +
                    "PREFIX : <http://www.semanticweb.org/GregorKaefer/ontologies/2020/11/excelOntologyVersion2#>\n" +
                    "PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>\n" +
                    "PREFIX ns: <http://xmlns.com/foaf/0.1/>\n" +
                    "\n" +
                    "SELECT DISTINCT ?columnTitle\n" +
                    "WHERE {\n" +
                    "    ?worksheet rdf:type :Worksheet.\n" +
                    "    ?worksheet :SheetName ?worksheetName.\n" +
                    "    ?table rdf:type :Table.\n" +
                    "    ?worksheet :hasTable ?table.\n" +
                    "\t?column rdf:type :ColumnHeader.\n" +
                    "    ?column :ColumnHeaderID ?columnId.\n" +
                    "    ?column :ColumnHeaderTitle ?columnTitle.\n" +
                    "    ?table :hasColumnHeader ?column.\n" +
                    "    ?row rdf:type :RowHeader.\n" +
                    "    ?row :RowHeaderTitle ?rowTitle.\n" +
                    "    ?row :RowHeaderID ?rowId.\n" +
                    "    ?table :hasRowHeader ?row.\n" +
                    "    ?value rdf:type :Value.\n" +
                    "    ?value :ValueFor ?column.\n" +
                    "    ?value :ActualValue ?valueConstraint.\n" +
                    "    ?row :hasValue ?value.\n";


            if(!isNumeric) value = "'" + value + "'";

            queryString += "Filter (?value "+ operator.getOperator() + " " + value + "). \n";


            queryString +=
                    "Filter(?rowID = '" + typeID.replaceAll("[a-zA-Z]","") + "').\n" +
                            "Filter(?worksheetName = '" + worksheetName +"').\n" +
                            "}\n";
            search = "columnTitle";

        }


        Query query = QueryFactory.create(queryString);

        QueryExecution qExec = QueryExecutionFactory.create(query, this.model);
        try {
            ResultSet resultSet = qExec.execSelect();
            while (resultSet.hasNext()) {
                QuerySolution solt = resultSet.nextSolution();
                result.add(solt.get(search).toString());
            }
        } finally {
            qExec.close();
        }

        return result;

    }

    @Override
    public Collection<String> getReverseDependencies(String cellID, String worksheetName) {
        String queryString = "PREFIX owl: <http://www.w3.org/2002/07/owl#>\n" +
                "PREFIX xsd: <http://www.w3.org/2001/XMLSchema#>\n" +
                "PREFIX rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#>\n" +
                "PREFIX : <http://www.semanticweb.org/GregorKaefer/ontologies/2020/11/excelOntologyVersion2#>\n" +
                "PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>\n" +
                "PREFIX ns: <http://xmlns.com/foaf/0.1/>\n" +
                "\n" +
                "SELECT DISTINCT ?columnDependencyIsUsedIn\n" +
                "WHERE {\n" +
                "    ?worksheet rdf:type :Worksheet.\n" +
                "    ?worksheet :SheetName ?worksheetName.\n" +
                "    ?table rdf:type :Table.\n" +
                "    ?worksheet :hasTable ?table.\n" +
                "    ?column rdf:type :ColumnHeader.\n" +
                "    ?column :ColumnHeaderID ?columnId.\n" +
                "    ?column (:isUsedIn)+ ?valueType.\n" +
                "    ?table :hasColumnHeader ?column.\n" +
                "    ?columnDependencyIsUsedIn rdf:type :ColumnHeader.\n" +
                "    ?columnDependencyIsUsedIn :hasValueType ?value.\n" +
                "    ?valueType :FormulaFunction ?function1.\n" +
                "    ?value :FormulaFunction ?function2.\n" +
                "    Filter(?columnId = '" + cellID.replaceAll("\\d","") + "').\n" +
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
                dependencyListIsUsedIn.add(solt.getResource("columnDependencyIsUsedIn").getLocalName());
            }
        } finally {
            qExec.close();
        }

        return dependencyListIsUsedIn;

    }

    /**
     * add com.java.eto.entity.excel data into the rdf-graph
     */
    private void convertExcelToOntology() throws IncorrectTypeException {
        //System.out.println("converting 1 workbook");
        Individual workbook = addWorkbook();
        //System.out.println("converting " + this.workbook.getWorksheets().size() + " worksheet");
        for(int i = 0;i<this.workbook.getWorksheets().size();i++){
            //  System.out.println("worksheet " + (i+1));
            Individual worksheet = addWorksheet(this.workbook.getWorksheets().get(i));
            workbook.addProperty(this.tableModel.getHasWorksheet(), worksheet);
            List<SheetElement> table =  new ArrayList<>();

            //initialize list with the saved data in map.
            for(Map.Entry<ElementType, List<SheetElement>> m : this.workbook.getWorksheets().get(i).getSheets().entrySet()) {
                switch (m.getKey()){
                    case TABLE:
                        table = m.getValue();
                        break;
                    default: break;
                }
            }

            //check if worksheet contain basic + sheet element or not
            if(table.size() != 0) {
                for(int j = 0; j<table.size(); j++) {
                    addTable((Table) table.get(j), worksheet, this.workbook.getWorksheets().get(i));
                }
            }
        }
    }

    private void addTable(Table table, Individual worksheet, Worksheet ws) throws IncorrectTypeException {
        Individual table1 = this.tableModel.getTable().createIndividual(this.tableModel.getWorkbook().getURI() + "_" +
                ws.getSheetName() + "_" + table.id());
        table1.addLiteral(RDFS.label, table.id());
        //table1.addLiteral(this.converter.getRowEnd(), table.getRowEnd());
        //table1.addLiteral(this.converter.getRowStart(), table.getRowStart());
        table1.addLiteral(this.tableModel.getTableName(), table.id());
        for(int i = 0; i< table.getColumns().size(); i++) {
            Column column = table.getColumns().get(i);
            Individual columnHeader = this.tableModel.getColumnHeader().createIndividual(this.tableModel.getWorkbook().getURI() + "_" +
                    ws.getSheetName() + "_" + column.getColumnTitle().replaceAll(" ","_"));
            columnHeader.addLiteral(this.tableModel.getColumnHeaderID(), column.getColumnID());
            columnHeader.addLiteral(this.tableModel.getColumnHeaderTitle(), column.getColumnTitle());
            columnHeader.addLiteral(RDFS.label, column.getColumnTitle().replaceAll(" ", "_"));
            Individual val;
            if(column.getValue() != null) {
                switch (column.getValue()) {
                    case STRING:
                        val = this.tableModel.getString().createIndividual(table1.getURI() + "_" +
                                column.getColumnTitle().replaceAll(" ","_") +"_StringValueType");
                        val.addLiteral(RDFS.label, "String");
                        columnHeader.addProperty(this.tableModel.getHasValueType(), val);
                        break;
                    case BOOLEAN:
                        val = this.tableModel.getBool().createIndividual(table1.getURI()+ "_" +
                                column.getColumnTitle().replaceAll(" ","_") +"_BooleanValueType");
                        val.addLiteral(RDFS.label, "Boolean");
                        columnHeader.addProperty(this.tableModel.getHasValueType(), val);
                        break;
                    case ERROR:
                        val = this.tableModel.getError().createIndividual(table1.getURI()+ "_" +
                                column.getColumnTitle().replaceAll(" ","_") +"_ErrorValueType");
                        val.addLiteral(RDFS.label, "Error");
                        columnHeader.addProperty(this.tableModel.getHasValueType(), val);
                        break;
                    case NUMERIC:
                        val = this.tableModel.getNumber().createIndividual(table1.getURI()+ "_" +
                                column.getColumnTitle().replaceAll(" ","_") +"_NumericValueType");
                        val.addLiteral(RDFS.label, "Numeric");
                        columnHeader.addProperty(this.tableModel.getHasValueType(), val);
                        break;
                    case FORMULA:
                        val = this.tableModel.getFormula().createIndividual(table1.getURI()+ "_" +
                                column.getColumnTitle().replaceAll(" ","_") +"_FormulaValueType");
                        val.addLiteral(RDFS.label, "Formula");
                        val.addLiteral(this.tableModel.getFormulaFunction(), column.getFormulaValue().getFormulaFunction());
                        for (int k = 0; k < column.getFormulaValue().getCellDependency().size(); k++) {
                            Cell cell = column.getFormulaValue().getCellDependency().get(k);
                            Individual cell1 = this.tableModel.getValue().createIndividual(this.tableModel.getWorkbook().getURI() + "_" +
                                    cell.getWorksheet().getSheetName() + "_" + cell.getTableName() + "_" +
                                    cell.getColumnTitle().replaceAll(" ", "_") + "_Value_For_Row" + cell.getRowID());
                            cell1.addProperty(this.tableModel.getIsUsedIn(), val);
                            switch (cell.getValue()) {
                                case STRING:
                                    cell1.addLiteral(this.tableModel.getActualValue(), cell.getStringValue());
                                    break;
                                case NUMERIC:
                                    cell1.addLiteral(this.tableModel.getActualValue(), cell.getNumericValue());
                                    break;
                                case ERROR:
                                    cell1.addLiteral(this.tableModel.getActualValue(), cell.getError().getErrorName());
                                    break;
                                case BOOLEAN:
                                    cell1.addLiteral(this.tableModel.getActualValue(), cell.isBooleanValue());
                                    break;
                                case FORMULA:
                                    switch (cell.getFormulaValue().getValueType()) {
                                        case BOOLEAN:
                                            cell1.addLiteral(this.tableModel.getActualValue(), cell.getFormulaValue().getBooleanValue());
                                            break;
                                        case STRING:
                                            cell1.addLiteral(this.tableModel.getActualValue(), cell.getFormulaValue().getStringValue());
                                            break;
                                        case NUMERIC:
                                            cell1.addLiteral(this.tableModel.getActualValue(), cell.getFormulaValue().getNumericValue());
                                            break;
                                        case ERROR:
                                            cell1.addLiteral(this.tableModel.getActualValue(), "ERROR");
                                            break;
                                    }
                                    break;
                            }
                            val.addProperty(this.tableModel.getHasDependency(), cell1);
                        }
                        for (int k = 0; k < column.getFormulaValue().getColumnDependency().size(); k++) {
                            Column column1 = column.getFormulaValue().getColumnDependency().get(k);
                            Individual col = this.tableModel.getColumnHeader().createIndividual(this.tableModel.getWorkbook().getURI() + "_" +
                                    column1.getWorksheet().getSheetName()  + "_" + column1.getColumnTitle().replaceAll(" ", "_"));
                            col.addProperty(this.tableModel.getIsUsedIn(), val);
                            val.addProperty(this.tableModel.getHasDependency(), col);
                        }
                        columnHeader.addProperty(this.tableModel.getHasValueType(), val);
                        break;
                }

                table1.addProperty(this.tableModel.getHasColumnHeader(), columnHeader);
            }
        }
        for(int i = 0; i<table.getCell().size();i++) {
            Cell cell = table.getCell().get(i);
            if(!cell.isSameValueAsColumnHeader()) {
                throw new IncorrectTypeException(cell.getCellId() + " doesnt have the same datatype as " + cell.getColumn());
            }
            Individual cell1 = this.tableModel.getValue().createIndividual(table1.getURI()+"_"+
                    cell.getColumnTitle().replaceAll(" ", "_")+"_Value_For_Row"+
                    cell.getRowID().replaceAll(" ", "_"));
            cell1.addLiteral(RDFS.label, cell.getColumnTitle()+"_Value_For_Row"+cell.getRowID());
            switch (cell.getValue()) {
                case STRING:
                    cell1.addLiteral(this.tableModel.getActualValue(), cell.getStringValue());
                    break;
                case NUMERIC:
                    cell1.addLiteral(this.tableModel.getActualValue(), cell.getNumericValue());
                    break;
                case ERROR:
                    cell1.addLiteral(this.tableModel.getActualValue(), cell.getError().getErrorName());
                    break;
                case BOOLEAN:
                    cell1.addLiteral(this.tableModel.getActualValue(), cell.isBooleanValue());
                    break;
                case FORMULA:
                    switch(cell.getFormulaValue().getValueType()) {
                        case BOOLEAN:
                            cell1.addLiteral(this.tableModel.getActualValue(), cell.getFormulaValue().getBooleanValue());
                            break;
                        case STRING:
                            cell1.addLiteral(this.tableModel.getActualValue(), cell.getFormulaValue().getStringValue());
                            break;
                        case NUMERIC:
                            cell1.addLiteral(this.tableModel.getActualValue(), cell.getFormulaValue().getNumericValue());
                            break;
                        case ERROR:
                            cell1.addLiteral(this.tableModel.getActualValue(), "ERROR");
                            break;
                    }
                    cell1.addLiteral(this.tableModel.getFormulaFunction(), cell.getFormulaValue().getFormulaFunction());
                    break;
            }
            Individual col = this.tableModel.getColumnHeader().createIndividual(this.tableModel.getWorkbook().getURI() + "_" +
                    ws.getSheetName() + "_" +  cell.getColumnTitle().replaceAll(" ", "_"));
            cell1.addProperty(this.tableModel.getValueFor(), col);
            table1.addProperty(this.tableModel.getHasValue(), cell1);
        }

        for(int i = 0; i<table.getRows().size();i++) {
            Row row = table.getRows().get(i);
            Individual rowHeader = this.tableModel.getRowHeader().createIndividual(this.tableModel.getWorkbook().getURI() + "_" +
                    ws.getSheetName() + "_"+
                    row.getRowTitle().replaceAll(" ", "_"));
            if(row.getRowTitle() != null) {
                rowHeader.addLiteral(RDFS.label, row.getRowTitle());
            }
            Individual colHeader = this.tableModel.getColumnHeader().createIndividual(this.tableModel.getWorkbook().getURI() + "_" +
                    ws.getSheetName()  + "_" +  row.getColumnTitle().replaceAll(" ", ""));
            rowHeader.addProperty(this.tableModel.getHasColumnHeader(), colHeader);
            for(int j = 0; j<row.getCell().size();j++) {
                Cell cell = row.getCell().get(j);
                Individual cell1 = this.tableModel.getValue().createIndividual(table1.getURI()+"_"+
                        cell.getColumnTitle().replaceAll(" ", "_")+"_Value_For_Row"+
                        cell.getRowID().replaceAll(" ","_"));
                rowHeader.addProperty(this.tableModel.getHasValue(), cell1);
            }
            rowHeader.addLiteral(this.tableModel.getRowHeaderID(), row.getRowId());
            if(row.getRowTitle() != null) {
                rowHeader.addLiteral(this.tableModel.getRowHeaderTitle(), row.getRowTitle());
            }
            table1.addProperty(this.tableModel.getHasRowHeader(), rowHeader);
        }
        worksheet.addProperty(this.tableModel.getHasTable(), table1);
    }

    /**
     * save model into file with turtle format
     * @param fileName file model that want to be converted into file
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
        this.tableModel.getModel().write(out, "TTL");
    }

    @Override
    public OntModel getModel() {
        return model;
    }

    /**
     * add workbook data into the rdf-graph
     * @return workbook instance.
     */
    private Individual addWorkbook() {
        Individual workbook = this.tableModel.getWorkbook().createIndividual(this.tableModel.getWorkbook().getURI() + "_" +
                this.workbook.getFileName());
        workbook.addLiteral(RDFS.label, this.workbook.getFileName());
        return workbook;
    }

    /**
     * add worksheet data into the rdf-graph
     * @param worksheet for naming purposes(not really important can be removed i think)
     * @return worksheet instance.
     */
    private Individual addWorksheet(Worksheet worksheet){
        Individual i = this.tableModel.getWorksheet().createIndividual(this.tableModel.getWorksheet().getURI() + "_" +
                worksheet.getSheetName());
        i.addLiteral(this.tableModel.getSheetName(), worksheet.getSheetName());
        i.addLiteral(RDFS.label, worksheet.getSheetName());
        return i;
    }

}
