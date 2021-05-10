package ExcelReader;

import java.util.ArrayList;
import java.util.List;

class FormulaSplitter {
    private List<String> result = new ArrayList<>();
    private String functions;

    public FormulaSplitter(String function) {
        this.functions = function;
        getFunctions();
    }

    public void getFunctions() {
        functions = functions.replaceAll(" ", "");
        int countOpenParentheses = 0;
        int countClosedParentheses = 0;
        String resultTemp = "";
        for(int i = 0; i<functions.length(); i++) {
            char temp = functions.charAt(i);
            if(temp == ';' || temp == ',') {
                if(countClosedParentheses == countOpenParentheses) {
                    result.add(resultTemp);
                    resultTemp = "";
                    countClosedParentheses = 0;
                    countOpenParentheses = 0;
                    continue;
                }
            }
            if(temp == ')' && countOpenParentheses == 0) {
                continue;
            }
            if((temp == '>' || temp == '=' || temp == '<') && countOpenParentheses == countClosedParentheses) {
                result.add(resultTemp);
                resultTemp = "";
                continue;
            }
            resultTemp = resultTemp + temp;
            if(temp == '(') {
                countOpenParentheses++;
            } else if(temp == ')') {
                countClosedParentheses++;
            }
            if(i == functions.length()-1) {
                result.add(resultTemp);
            }
        }
    }

    public List<String> getResult() {
        return result;
    }

}
