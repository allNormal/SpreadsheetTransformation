package ExcelReader;

import Entity.ValueType.BasicFormula;
import Entity.ValueType.Formula;
import Entity.ValueType.NestedFormula;
import java.util.ArrayList;
import java.util.List;

import static Entity.ValueType.FunctionType.*;

class FormulaConverter {
    private List<String> functionList;
    private String functions;
    private Formula formula;
    private List<String> getFunctionList;

    public FormulaConverter(String functions) {
        this.functions = functions;
        this.getFunctionList = getFunctions(functions);
        if(checkNested(functions) == "Basic") {
            this.formula = new BasicFormula(functions);
            this.formula.setFunctionType(BASIC);
        } else {
            this.formula = new NestedFormula(functions);
            this.formula.setFunctionType(NESTED);
        }
        for (int i = 0; i < this.getFunctionList.size(); i++) {
            String[] result = this.getFunctionList.get(i).split("\\(", 2);
            String result2 = "";
            if (result.length > 1) {
                if (result[1].charAt(result[1].length() - 1) == ')') {
                    result2 = result[1].substring(0, result[1].length() - 1);
                } else {
                    result2 = result[1].substring(0, result[1].length());
                }
            } else if (result.length == 1) {
                result2 = result[0];
            }
            FormulaSplitter splitter = new FormulaSplitter(result2);
            this.functionList = splitter.getResult();
            initializeFormula(this.getFunctionList.get(i).replaceAll(" ",""));
        }
    }

    private List<String> getFunctions(String functions) {
        int countOpenParentheses = 0;
        int countClosedParentheses = 0;
        String resultTemp = "";
        List<String> result = new ArrayList<>();
        for(int i = 0;i< functions.length(); i++){
            char temp = functions.charAt(i);
            if(temp == '(') {
                countOpenParentheses++;
            }
            if(temp == ')') {
                countClosedParentheses++;
            }
            if(temp == '+' || temp == '-' || temp == '*' || temp == '/' || temp == '%'){
                if(countClosedParentheses == countOpenParentheses) {
                    result.add(resultTemp);
                    countClosedParentheses = 0;
                    countOpenParentheses = 0;
                    resultTemp = "";
                    continue;
                }
            }
            resultTemp = resultTemp + temp;
            if(i == functions.length()-1) {
                result.add(resultTemp);
            }
        }
        return result;
    }

    private String checkNested(String formula) {
        if(formula.split("\\(").length > 2) {
            return "Nested";
        }
        return "Basic";
    }

    private void initializeFormula(String functions) {
        if(checkNested(functions).equals("Basic")) {
            BasicFormula formula = new BasicFormula(functions);
            formula.setFunctionType(BASIC);
            this.formula.add(formula);
        } else {
            NestedFormula formula = new NestedFormula(functions);
            formula.setFunctionType(NESTED);
            for(int i = 0; i<this.functionList.size(); i++) {
                if(checkNested(this.functionList.get(i)) == "Basic") {
                    BasicFormula formula1 = new BasicFormula(this.functionList.get(i));
                    formula1.setFunctionType(BASIC);
                    formula.add(formula1);
                } else {
                    NestedFormula formulaTemp = new NestedFormula(this.functionList.get(i));
                    formulaTemp.setFunctionType(NESTED);
                    FormulaConverter formulaConverter = new FormulaConverter(this.functionList.get(i));
                    formulaTemp.add(formulaConverter.getFormula());
                    formula.add(formulaTemp);
                }
            }
            this.formula.add(formula);
        }
    }

    public Formula getFormula() {
        return formula;
    }

}
