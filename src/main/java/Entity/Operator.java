package Entity;

public enum Operator {
    EQ("="), LT("<"), GT(">"), LE("<="), GE(">="), NE("!=");

    private String operator;
    private Operator antonym;

    static {
        EQ.antonym = NE;
        LT.antonym = GT;
        GT.antonym = LT;
        LE.antonym = GE;
        GE.antonym = LE;
        NE.antonym = EQ;
    }

    Operator(String operator) {
        this.operator = operator;
    }


    public String getOperator() {
        return operator;
    }

    public Operator getAntonym() {
        return antonym;
    }

}
