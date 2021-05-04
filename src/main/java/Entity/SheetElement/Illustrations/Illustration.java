package Entity.SheetElement.Illustrations;

import Entity.SheetElement.SheetElement;
import Entity.Worksheet.Worksheet;

public class Illustration extends SheetElement {
    private String title;

    public Illustration(Worksheet worksheet, String title){
        super(worksheet);
        this.title = title;
    }

    public String id() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

}
