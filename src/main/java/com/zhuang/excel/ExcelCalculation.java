package com.zhuang.excel;

import java.math.BigDecimal;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * excel公式计算 包含 加减乘除 负号 单个百分号 IF AND OR MIN MAX SUM RANK
 */
public class ExcelCalculation {
    /**
     * 函数公式或字母加数字（F2）或者 数字百分号 或 比较符号
     */
    private static final String REGEX = "^[A-Z]+\\(|^[A-Z]\\d{1,2}|^-?\\d+%|^[<>]=?|^=";
    // 数字(包括含百分号)或单元格名称（F2）
    private static final String REGEX3 = "^-?([0-9]{1,}[.][0-9]*)$|^-?([0-9]{1,})$|^\\d+%|^[A-Z]\\d{1,2}";
    // 数字
    private static final String NUMBER_REGEX = "^-?([0-9]{1,}[.][0-9]*)$|^-?([0-9]{1,})$";
    // 单元格名
    private static final String CEIl_REGEX = "^[A-Z]\\d{1,2}";
    // 函数公式或比较符号
    private static final String REGEX2 = "^[A-Z]+\\(|^[<>]=?|^=";
    // 函数公式名称
    private static final String FORMULA_REGEX = "^[A-Z]+\\(";
    /**
     * 含百分比的数字
     */
    private static final String REGEX4 = "^\\d+%";
    // 除法，保留小数位数
    private static final int DIGITS = 10;
    // 除法，保留小数，舍入方式
    private static final int ROUNDING_MODE = BigDecimal.ROUND_HALF_UP;
    /**
     * 隔的行数
     */
    private static int separateRow = 0;
    private static int separateColumn = 0;
//    private static int separateRow = 2;//3;
//    private static int separateColumn = 5;

    // 循环次数，避免死循环
    private static int cycleNum = 0;

    /**
     * 错误种类
     */
    private enum ErrorMessageEnum {
        NAME_ERROR("#NAME?"), ENDLESS_LOOP("#ENDLESS_LOOP?"), FORMULA_ERROR("#FORMULA_ERROR?"),
        ERROR("#ERROR?");

        private final String text;
        ErrorMessageEnum(final String text){
            this.text=text;
        }
        @Override
        public String toString(){
            return text;
        }
    }

    //    static String[][] table = {
//            {"=A3+A4-A5", "3", "3"},
//            {"=A1", "=B2"},
//            {"3"},
//            {"4"},
//            {"=A5"},
//            {"6"},
//            {"=A6*A3"}
//    };
    static String[][] table = {
            {null, "=A0+2", "3"},
            {"=A1", null, "=B2"},
            {"3", "IF(AND(A3<A4, 1%<2, 3 < 5 * 103%), MIN(A3, A4), 0%)"},
            {"=A6+A5"},
            {"5"},
            {"6"},
            {"=A6*A3"}
    };
//    =RANK(E54,(E54,I54,K54,L54),0)
//    =RANK(F24,$F$24:$N$24,0)     最后一个参数可以省略， =RANK(F24,$F$24:$N$24)
//    =SUM(E2,F2,G2)
//    =SUM(E2:G2)
//    static String[][] table = {
//        {"191.087684"},
//        {"248.880452"},
//        {"16"},
//        {"16.5"},
//        {"=F3/F5"},
//        {"=F4/F6"},
//        {"=F8/F7-1"},
//        {"=IF(F9<-15%,90%,IF(AND(F9>=-15%,F9<0%),95%,IF(AND(F9>=0%,F9<3%),100%,IF(AND(F9>=3%,F9<5%),104%,IF(AND(F9>=5%,F9<8%),108%,110%)))))"},
//        {"=50*F10"},
//        {"3719"},
//        {"3747"},
//        {"=ABS((F12-F13)/F13)"},
//        {"=IF(F13<100,10%,IF(AND(F13>=100,F13<1000,ABS(F12-F13)<=20),5%,IF(AND(F13>=1000,ABS(F12-F13)<=30),2%,0)))"},
//        {"=IF(F14>F15,0%,IF(AND(F14<=F15,F14>0.8*F15),100%,MIN(100%+(ABS(F15-F14)/(F15*20%))*2%,110%)))"},
//        {"=50*F16"},
//        {null},
//        {null},
//        {null},
//        {null},
//        {null},
//        {"=F11+F17+F22"},
//        {"=F23/100", "1.05","1.08","1.07","1.01","1.02","1.07","1.02"},//,"1.082"
//        {"=RANK(F24,$F$24:$N$24,0)"},
//        {"=IF(OR(F25=1,F25=2),1,IF(OR(F25=3,F25=4),0.5,IF(F25=5,0,IF(OR(F25=6,F25=7),-0.5,-1))))"},
//        {"=(F24-1)*1*10"},
//        {"=F26*0.6+F27*0.4"},
//        {null}
//    };

    private static List<String> parseToSuffixExpression(List<String> expressionList) {
        //创建一个栈用于保存操作符
        Stack<String> opStack = new Stack<>();
        //创建一个list用于保存后缀表达式
        List<String> suffixList = new ArrayList<>();
        for(String item : expressionList){
            //得到数或操作符
            if(isOperator(item)){
                //是操作符 判断操作符栈是否为空
                if(opStack.isEmpty() || "(".equals(opStack.peek()) || priority(item) > priority(opStack.peek())){
                    //为空或者栈顶元素为左括号或者当前操作符大于栈顶操作符直接压栈
                    opStack.push(item);
                }else {
                    //否则将栈中元素出栈如队，直到遇到大于当前操作符或者遇到左括号时或者其他函数公式
                    while (!opStack.isEmpty() && !"(".equals(opStack.peek()) ){
                        if(priority(item) <= priority(opStack.peek())){
                            suffixList.add(opStack.pop());
                        } else {
                            break;
                        }
                    }
                    //当前操作符压栈
                    opStack.push(item);
                }
            }else if(item.matches(REGEX3)){
                //是数字则直接入队
                suffixList.add(item);
            }else if("(".equals(item) || ",".equals(item) || isFunOperator(item) || item.matches(REGEX2)){
                //是左括号，压栈
                opStack.push(item);
            }else if(")".equals(item)){
                //是右括号 ，将栈中元素弹出入队，直到遇到左括号或者函数公式，左括号出栈，但不入队,函数公式出栈且入队。
                while (!opStack.isEmpty()){
                    if("(".equals(opStack.peek())){
                        opStack.pop();
                        break;
                    } else if (opStack.peek().matches(FORMULA_REGEX)) {
                        suffixList.add(opStack.pop());
                        break;
                    } else {
                        suffixList.add(opStack.pop());
                    }
                }
            }else if(".".equals(item)){
                suffixList.add(item);
            }else if(":".equals(item)){
                suffixList.add(item);
            }else {
                throw new RuntimeException("有非法字符："+ item);
            }
        }
        //循环完毕，如果操作符栈中元素不为空，将栈中元素出栈入队
        while (!opStack.isEmpty()){
            suffixList.add(opStack.pop());
        }
        return suffixList;
    }

    /**
     * 判断字符串是否为基础操作符，包含逗号
     * @param op
     * @return
     */
    public static boolean isOperator(String op){
        return op.equals("+") || op.equals("-") || op.equals("*") || op.equals("/") || op.equals(",");
    }

    /**
     * 判断字符串是否为特殊操作符
     * @param op
     * @return
     */
    public static boolean isFunOperator(String op){
        return op.matches("^-\\(") || op.matches("^[A-Z]+\\(");
    }

    /**
     * 判断是否为数字
     * @param num
     * @return
     */
    public static boolean isNumber(String num){
        return num.matches("^-?([0-9]{1,}[.][0-9]*)$") || num.matches("^-?([0-9]{1,})$");
    }

    /**
     * 获取操作符的优先级
     * @param op
     * @return
     */
    public static int priority(String op){
        if(op.equals("*") || op.equals("/")){
            return 1;
        }else if(op.equals("+") || op.equals("-")){
            return 0;
        } else if (op.matches("^[<>]=?|=")) {
            return -1;
        } else if (op.matches(",")) {
            return -2;
        } else if (op.matches("\\(") || op.matches("\\)") || isFunOperator(op)) {
            return -3;
        }
        return -4;
    }

    /**
     * 将表达式转为list
     * @param expression
     * @return
     */
    private static List<String> expressionToList(String expression) throws Exception {
        expression = expression.replaceAll(" ", "");
        int index = 0;
        List<String> list = new ArrayList<>();
        do{
            char ch = expression.charAt(index);

            // 是操作符 “+-*/ ,:” , -(不是负号),  冒号（58）；
            // Character.isDigit(expression.charAt(index-1) 表示-号前面为数字的时候
            // expression.charAt(index-1)==')') 表示-号前面为)的时候
            if((ch!=45 && ch!=46 && ch <= 47) || ch == 58 || (index!=0 && ch==45 && (Character.isDigit(expression.charAt(index-1)) || expression.charAt(index-1)==')'))){
                //是操作符，直接添加至list中
                index ++ ;
                list.add(ch+"");
            }else{
                // 特殊公式
                String a = expression.substring(index);
                Pattern pattern = Pattern.compile(REGEX);
                Matcher matcher = pattern.matcher(a);
                if (matcher.find()) {
                    String str = matcher.group(0);
                    // 含百分比的数值
                    if (str.endsWith("%")) {
                        str = new BigDecimal(str.substring(0,str.length()-1)).divide(BigDecimal.valueOf(100)) + "";
                    }
                    list.add(str);
                    index = index + matcher.group(0).length();
                    continue;
                }

                // 是数字和负数,判断多位数的情况 注：45 负号， 46 点
                String str = "";
                while (index < expression.length() && (expression.charAt(index) >47 && expression.charAt(index) < 58 || expression.charAt(index)==46
                        || (expression.charAt(index)==45 && !(index!=0 && (Character.isDigit(expression.charAt(index-1)) || expression.charAt(index-1)==')'))))){
                    str += expression.charAt(index);
                    index ++;
                }
                if ("-".equals(str)) {
                    // 负号
                    list.add("-(");
                } else {
                    // 匹配不到，则抛异常
                    if ("".equals(str)) {
                        System.out.println("无法解析公式含义");
                        throw new Exception(ErrorMessageEnum.ERROR.toString());
                    }
                    list.add(str);
                }
            }
        }while (index < expression.length());
        return list;
    }

    /**
     * 根据后缀表达式list计算结果
     * @param list
     * @return
     */
    private static BigDecimal calculate(List<String> list) throws Exception {
        Stack<String> stack = new Stack<>();
        // 逗号个数+1
        int count;
        for(int i=0; i<list.size(); i++){
            // 重置
            count = 1;
            String item = list.get(i);
            if(item.matches(REGEX3)){
                // 是数字
                stack.push(item);
            } else if (item.equals("-(") || item.equals("ABS(")){
                // 是操作符，取出栈顶一个元素
                BigDecimal num = new BigDecimal(explain(stack.pop()));
                BigDecimal res;
                if (item.equals("-(")) {
                    res = BigDecimal.ZERO.subtract(num);
                } else {
                    res = num.abs();
                }
                stack.push(res + "");
            } else if (item.equals(",")) {
                stack.push(",");
            }  else if (item.equals(":")) {
                stack.push(":");
            } else if (item.matches("[<>]=?|=")) {
                Boolean res = false;
                BigDecimal num = new BigDecimal(explain(stack.pop()));
                BigDecimal num2 = new BigDecimal(explain(stack.pop()));

                if (item.contains("=")) {
                    if (num.compareTo(num2) == 0) {
                        res = true;
                    }
                }
                if (item.contains("<")) {
                    if (num2.compareTo(num) == -1) {
                        res = true;
                    }
                } else if (item.contains(">")) {
                    if (num2.compareTo(num) == 1) {
                        res = true;
                    }
                }
                stack.push(res + "");
            } else if (item.equals("MIN(")) {
                List<String> lists = new LinkedList<>();
                BigDecimal min = BigDecimal.ZERO;

                for (int j=0; j<count; ) {
                    if(",".equals(stack.peek())) {
                        count ++;
                        stack.pop();
                    } else {
                        lists.add(stack.pop());
                        j++;
                    }
                }
                for (int j=lists.size()-1; j>=0; j--) {
                    BigDecimal temp = new BigDecimal(explain(lists.get(j)));
                    if (j == lists.size()-1) {
                        min = temp;
                    } else if (temp.compareTo(min) == -1) {
                        min = temp;
                    }
                }
                stack.push(min + "");
            } else if (item.equals("MAX(")) {
                List<String> lists = new LinkedList<>();
                BigDecimal max = BigDecimal.ZERO;

                for (int j=0; j<count; ) {
                    if(",".equals(stack.peek())) {
                        count ++;
                        stack.pop();
                    } else {
                        lists.add(stack.pop());
                        j++;
                    }
                }

                for (int j=lists.size()-1; j>=0; j--) {
                    BigDecimal temp = new BigDecimal(explain(lists.get(j)));
                    if (j == lists.size()-1) {
                        max = temp;
                    } else if (temp.compareTo(max) == 1) {
                        max = temp;
                    }
                }
                stack.push(max + "");
            } else if (item.equals("IF(")) {
                BigDecimal res;
                removeComma(stack);
                BigDecimal num = new BigDecimal(explain(stack.pop()));
                removeComma(stack);
                BigDecimal num2 = new BigDecimal(explain(stack.pop()));
                removeComma(stack);
                boolean num3 = parseBoolean(stack.pop());
                if (num3) {
                    res = num2;
                } else {
                    res = num;
                }
                stack.push(res + "");
            } else if (item.equals("AND(")) {
                boolean flag = true;
                List<String> lists = new LinkedList<>();

                for (int j=0; j<count; ) {
                    if(",".equals(stack.peek())) {
                        count ++;
                        stack.pop();
                    } else {
                        lists.add(stack.pop());
                        j++;
                    }
                }
                for (int j=lists.size()-1; j>=0; j--) {
                    if (!parseBoolean(lists.get(j))) {
                        flag = false;
                        break;
                    }
                }
                stack.push(flag + "");
            } else if (item.equals("OR(")) {
                boolean flag = false;
                List<String> lists = new LinkedList<>();
                for (int j=0; j<count; ) {
                    if(",".equals(stack.peek())) {
                        count ++;
                        stack.pop();
                    } else {
                        lists.add(stack.pop());
                        j++;
                    }
                }
                for (int j=lists.size()-1; j>=0; j--) {
                    if (parseBoolean(lists.get(j))) {
                        flag = true;
                        break;
                    }
                }
                stack.push(flag + "");
            } else if (item.equals("SUM(")) {
                BigDecimal sum = BigDecimal.ZERO;
                List<String> lists = new LinkedList<>();
                if(",".equals(stack.peek())) {
                    for (int j=0; j<count; ) {
                        if(",".equals(stack.peek())) {
                            count ++;
                            stack.pop();
                        } else {
                            lists.add(stack.pop());
                            j++;
                        }
                    }
                    for (int j=lists.size()-1; j>=0; j--) {
                        sum = sum.add(new BigDecimal(explain(lists.get(j))));
                    }
                    // SUM(A1:A2) 形式
                } else  {
                    String end = stack.pop();
                    stack.pop();
                    String start = stack.pop();
                    if (!start.matches(CEIl_REGEX) || !end.matches(CEIl_REGEX)) {
                        System.out.println("SUM公式错误");
                        throw new Exception(ErrorMessageEnum.NAME_ERROR.toString());
                    }
                    int row = Integer.parseInt(start.substring(1))-1-separateRow;
                    int column = start.charAt(0)-65-separateColumn;
                    int row2 = Integer.parseInt(end.substring(1))-1-separateRow;
                    int column2 = end.charAt(0)-65-separateColumn;
                    sum = getSum(row, column, row2, column2);
                }

                stack.push(sum + "");
            }  else if (item.equals("RANK(")) {
                int rank = 0;
                List<String> lists = new LinkedList<>();
                if(stack.search(":") == -1) {
                    for (int j=0; j<count; ) {
                        if(",".equals(stack.peek())) {
                            count ++;
                            stack.pop();
                        } else {
                            lists.add(stack.pop());
                            j++;
                        }
                    }
                    String firstStr = lists.get(lists.size() - 1);
                    lists.remove(lists.size() - 1);
                    String endStr = lists.get(0);
                    String rankWay = "";
                    if (!endStr.matches(CEIl_REGEX)) {
                        rankWay = lists.get(0);
                        lists.remove(0);
                    }

                    rank = getRank(firstStr, lists, rankWay);
                    // RANK(A1,A1:A2) 形式
                } else  {
                    String rankWay = "";
                    // RANK(A1,A1:A2,0) 形式
                    if (stack.size() == 7) {
                        stack.pop();
                        rankWay = stack.pop();
                    }
                    removeComma(stack);
                    String end = stack.pop();
                    stack.pop();
                    String start = stack.pop();
                    if (!start.matches(CEIl_REGEX) || !end.matches(CEIl_REGEX)) {
                        System.out.println("RANK公式错误");
                        throw new Exception(ErrorMessageEnum.NAME_ERROR.toString());
                    }
                    int row = Integer.parseInt(start.substring(1))-1-separateRow;
                    int column = start.charAt(0)-65-separateColumn;
                    int row2 = Integer.parseInt(end.substring(1))-1-separateRow;
                    int column2 = end.charAt(0)-65-separateColumn;
                    for(int x=row; x<=row2; x++) {
                        for(int y=column; y<=column2; y++) {
                            try {
                                lists.add(explain(table[x][y]));
                                // 由于数组没有赋初值，导致数据越界，则为该数据赋初值为零
                            } catch (IndexOutOfBoundsException e) {
                                System.out.println("数组越界");
                                lists.add("0");
                            }
                        }
                    }
                    String firstStr = stack.pop();
                    rank = getRank(firstStr, lists, rankWay);
                }

                stack.push(rank + "");
            } else {
                //是操作符，取出栈顶两个元素
                removeComma(stack);
                BigDecimal num2 = new BigDecimal(explain(stack.pop()));
                //System.out.print(num2);
                removeComma(stack);
                BigDecimal num1 = new BigDecimal(explain(stack.pop()));
                BigDecimal res;
                switch (item) {
                    case "+":
                        res = num1.add(num2);
                        break;
                    case "-":
                        res = num1.subtract(num2);
                        break;
                    case "*":
                        res = num1.multiply(num2);
                        break;
                    case "/":
                        res = num1.divide(num2, DIGITS, ROUNDING_MODE);
                        break;
                    default:
                        throw new RuntimeException("运算符错误：" + item);
                }

                stack.push(res + "");
            }

        }
        return new BigDecimal(explain(stack.pop()));
    }

    /**
     * 过滤栈顶的逗号
     * @param stack
     * @return 非逗号的栈顶值
     */
    private static void removeComma(Stack stack) throws Exception {
        try {
            if (",".equals(stack.peek())) {
                stack.pop();
            }
        } catch (EmptyStackException e) {
            throw new Exception(ErrorMessageEnum.FORMULA_ERROR.toString());
        }
    }

    /**
     * 判断   公式，计算
     *        数值，返回
     * @param s
     * @return
     */
    private static String explain(String s) throws Exception {
        // 单元格没有数据，赋值为0
        if (s == null ) {
            return "0";
        }
        // 是否为 错误信息，如果是错误的传递
        for (ErrorMessageEnum e : ErrorMessageEnum.values()) {
            if (e.toString().equals(s)) {
                return e.toString();
            }
        }

        cycleNum ++;
        if (cycleNum == 250) {
            // 程序死循环
            throw new Exception(ErrorMessageEnum.ENDLESS_LOOP.toString());
        }
        String a;
        int row = 0;
        int column = 0;
        // 公式 等号开头，去除等号； 去除$ 符号
        if (s.charAt(0) == 61) {
            s = s.substring(1);
            // 去除$ 符号
            s = s.replaceAll("\\$", "");
        }
        // 某单元格的值
        if (s.matches(CEIl_REGEX)) {
            try {
                row = Integer.parseInt(s.substring(1))-1-separateRow;
                column = s.charAt(0)-65-separateColumn;
                s = table[row][column];
                // 递归
                a = explain(s);
                table[row][column] = table[row][column] == null ? null : a;
            } catch (NumberFormatException | ArrayIndexOutOfBoundsException e) {
                throw new Exception(ErrorMessageEnum.NAME_ERROR.toString());
            }
            // 数值
        } else if (s.matches(NUMBER_REGEX)) {
            a = s;
            // 公式
        } else {
            a = formula(s);
        }

        return a;
    }

    /**
     * 转换成boolean，是数值且不为零 或者 为‘ture’ 为真
     * @param s
     * @return
     */
    public static boolean parseBoolean(String s) {
        return (s != null) && (s.equalsIgnoreCase("true") || (s.matches("\\d+") && !s.equals("0")));
    }

    /**
     * 公式计算
     * @param expression
     * @return
     */
    private static String formula(String expression) throws Exception {
        String result;
        List<String> expressionList = expressionToList(expression);
        //System.out.println("中缀表达式转为list结构="+expressionList);
        //将中缀表达式转换为后缀表达式
        List<String> suffixList = parseToSuffixExpression(expressionList);
        //System.out.println("对应的后缀表达式列表结构="+suffixList);

        //根据后缀表达式计算结果
        BigDecimal calculateResult = calculate(suffixList);
        result = calculateResult.stripTrailingZeros().toPlainString();

        return result;
    }

    public static BigDecimal getSum(int x1,int y1,int x2,int y2) throws Exception {
        BigDecimal sum = BigDecimal.ZERO;
        for(int i=x1;i<=x2;i++) {
            for(int j=y1;j<=y2;j++) {
                sum = sum.add(new BigDecimal(explain(table[i][j])));
            }
        }
        return sum;
    }

    /**
     * 获取某个单元格（或数值）在某个范围内的排名，该单元格（或数值）必须在该范围里
     * @param str 某个单元格（或数值）
     * @param list 引用（某些单元格）
     * @param rankWay 排位方式，为空或为0，降序， 其他，升序。
     * @return
     */
    private static Integer getRank(String str, List<String> list, String rankWay) throws Exception {
        str = explain(str);
        if (!str.matches(NUMBER_REGEX)) {
            System.out.println("RANK所选单元格错误");
            throw new Exception(str);
        }
        BigDecimal tmp = new BigDecimal(str);
        // 比较大小，根据排位方式，选择大于或小于来进行排名比较, -1 降序，1 升序。
        int compareNum = -1;
        if (rankWay != null && !"".equals(rankWay) && !"0".equals(rankWay)) {
            compareNum = 1;
        }
        // 该单元格排是否在引用中
        boolean isExist = false;
        // 排名
        int rank = 1;
        for (int i = 0; i < list.size(); i++) {
            BigDecimal currentDecimal = new BigDecimal(explain(list.get(i)));
            if (tmp.equals(currentDecimal)) {
                isExist = true;
            }
            if (tmp.compareTo(currentDecimal) == compareNum) {
                rank += 1;
            }
        }
        if (!isExist) {
            System.out.println("排名：数值不在该范围内");
            throw new Exception(ErrorMessageEnum.ERROR.toString());
        }
        return rank;
    }

    public static void main(String []args){
//        String expression = "IF(F4>F5,0%,IF(AND(F4<=F5,F4>0.8*F5),100%,MIN(100%+(ABS(F5-F4)/(F5*20%))*2%,110%)))";
//        String expression = "IF(AND(1<2,3<2,1<3),2+3*5-6,3/4+3)";
//        String expression = "IF(AND(3<5,4<5,1),2+4,5)";
//        String expression = "8+IF(OR(5<3,6<%5,0),2+4,5)";
//        String expression = "-(-(-2))";
//        String expression = "IF(1<0,90,IF( AND(1 >=0,1 <5, 1),MIN( 100%+33.33434,-(343+2)),IF(AND(1>=3,1<5),104%,IF(AND(1>=5%,1<10%),108%,110))))";
//        String expression = "IF(1<0,90,IF(AND(1>=0,0),100,1))";
//        String expression = "IF(A4>A5,0%,IF(AND(A4<=A5,A4>0.8*A5),100%,MIN(100%+(ABS(A5-A4)/(A5*20%))*2%,110%)))";
//        System.out.println( "table = {");
//        for (int i=0; i<table.length; i++) {
//            for (int j=0; j<table[i].length; j++) {
//                System.out.print(table[i][j] + ",\t");
//            }
//            System.out.println();
//        }
//        System.out.println("}");

        for (int i=0; i<table.length; i++) {
            for (int j=0; j<table[i].length; j++) {
                try {
                    if (table[i][j].contains("RANK(")) {
                        System.out.println(table[i][j]);
                    }
                    table[i][j] = table[i][j] == null ? null : explain(table[i][j]);
                } catch (Exception e) {
                    System.out.println(table[i][j] + "公式：" + e.getMessage());
                    table[i][j] = e.getMessage();
                }
                cycleNum = 0;
            }
        }

        System.out.println( "table = {");
        for (int i=0; i<table.length; i++) {
            for (int j=0; j<table[i].length; j++) {
                System.out.print(table[i][j] + ",\t");
            }
            System.out.println();
        }
        System.out.println("}");

    }

}

