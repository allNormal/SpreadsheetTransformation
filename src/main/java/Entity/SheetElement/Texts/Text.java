package Entity.SheetElement.Texts;

import Entity.SheetElement.SheetElement;
import Entity.Worksheet.Worksheet;

public class Text extends SheetElement {
    private String name;
    private String value;

    public Text(Worksheet worksheet, String name, String value) {
        super(worksheet);
        this.name = name;
        this.value = value;
    }

    public String id() {
        return name;
    }

    public String getValue() {
        return value;
    }

}
