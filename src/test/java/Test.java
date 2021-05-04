import ExcelReader.*;
import Transformer.ExcelTransformer;
import Transformer.TableTransformer;
import Exception.IncorrectTypeException;
import java.io.File;
import java.io.IOException;

public class Test {
    public static void main(String[] args) {
        try {
            ExcelBasedReader excelBasedReader = new ExcelBasedReader(new File("src/main/resources/simple_test.xlsm"));
            excelBasedReader.readExcelConverter();
            TableBasedReader tableBasedReader = new TableBasedReader(new File("src/main/resources/simple_test.xlsm"));
            tableBasedReader.readExcelConverter();
            ExcelTransformer excelTransformer = new ExcelTransformer();
            TableTransformer tableTransformer = new TableTransformer();
            tableTransformer.create(tableBasedReader.getWorkbook());
            excelTransformer.create(excelBasedReader.getWorkbook());
            excelTransformer.saveModel("src/main/resources/", "test");
            tableTransformer.saveModel("src/main/resources/", "test2");
        } catch (IOException e) {
            e.getMessage();
        } catch (IncorrectTypeException e) {
            System.out.println(e.getMessage());
        }
    }
}
