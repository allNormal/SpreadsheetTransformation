package Entity.SheetElement.BasicElement;

import Entity.SheetElement.SheetElement;
import Entity.Worksheet.Worksheet;

public class BasicElement extends SheetElement {

    private String id;

    public BasicElement(Worksheet worksheet, String id){
        super(worksheet);
        this.id = id;
    }

    public String id() {
        return this.id;
    }

}
