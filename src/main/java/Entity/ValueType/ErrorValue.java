package Entity.ValueType;

public class ErrorValue {
    private String description;
    private String errorName;

    public ErrorValue(String errorName){
        this.errorName = errorName;
    }

    public String getDescription() {
        return description;
    }

    public String getErrorName() {
        return errorName;
    }

}
