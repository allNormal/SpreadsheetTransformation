package Entity.ValueType;

import java.util.ArrayList;
import java.util.List;

public class NestedFormula extends Formula {
    private List<Formula> FormulaList = new ArrayList<>();

    public NestedFormula(String formulaFunction) {
        super(formulaFunction);
    }

    public void add(Formula formula) {
        this.FormulaList.add(formula);
    }

    public List<Formula> getFormulaList() {
        return FormulaList;
    }

}
