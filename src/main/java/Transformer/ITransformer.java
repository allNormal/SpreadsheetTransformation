package Transformer;

import Entity.Operator;
import Entity.SheetElement.ElementType;
import Entity.Workbook.Workbook;
import ExcelReader.ExcelReader;
import org.apache.jena.ontology.OntModel;
import org.apache.jena.ontology.Restriction;
import Exception.*;

import java.util.Collection;
import java.util.List;

public interface ITransformer {

    void create() throws IncorrectTypeException;


    List<String> getCellDependencies(String cellID, String worksheetName);

    List<String> addConstraint(ElementType type, String typeID, String worksheetName,
                                     Operator operator, String value);

    List<String> getReverseDependencies(String cellID, String worksheetName);

    OntModel getModel();

    void saveModel(String path, String fileName);

}
