package Transformer;

import Entity.Operator;
import Entity.SheetElement.ElementType;
import Entity.Workbook.Workbook;
import org.apache.jena.ontology.OntModel;
import org.apache.jena.ontology.Restriction;
import Exception.IncorrectTypeException;

import java.util.Collection;

public interface ITransformer {

    void create(Workbook workbook) throws IncorrectTypeException;


    Collection<String> getCellDependencies(String cellID, String worksheetName);

    Collection<String> addConstraint(ElementType type, String typeID, String worksheetName,
                                     Operator operator, String value);

    Collection<String> getReverseDependencies(String cellID, String worksheetName);

    OntModel getModel();

    void saveModel(String path, String fileName);

}
