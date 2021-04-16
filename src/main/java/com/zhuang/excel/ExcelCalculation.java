package com.zhuang.excel;

import java.math.BigDecimal;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * excel公式计算 包含 加减乘除 比较判断 负号 单个百分号 IF AND OR MIN MAX SUM RANK
 * 未 AVG（在excel为 AVERAGE）、COUNT、=SUM(B1:B2,B1:B2,3)
 *
 * @author zhuang
 */
public class ExcelCalculation {
    /**
     * 函数公式或字母加数字（F2）或者 数字百分号 或 比较符号
     */
    private static final String REGEX = "^[A-Z]+\\(|^[A-Z]\\d{1,2}|^-?\\d+%|^[<>]=?|^=";
    /**
     * 数字(包括含百分号)或单元格名称（F2）
     */
    private static final String REGEX3 = "^-?([0-9]{1,}[.][0-9]*)$|^-?([0-9]{1,})$|^\\d+%|^[A-Z]\\d{1,2}";
    /**
     * 数字
     */
    private static final String NUMBER_REGEX = "^-?([0-9]{1,}[.][0-9]*)$|^-?([0-9]{1,})$";
    /**
     * 单元格名
     */
    private static final String CEIL_REGEX = "^[A-Z]\\d{1,2}";
    /**
     * 函数公式或比较符号
     */
    private static final String REGEX2 = "^[A-Z]+\\(|^[<>]=?|^=";
    /**
     * 函数公式名称
     */
    private static final String FORMULA_REGEX = "^[A-Z]+\\(";
    /**
     * 除法，保留小数位数
     */
    private static final int DIGITS = 10;
    /**
     * 除法，保留小数，舍入方式
     */
    private static final int ROUNDING_MODE = BigDecimal.ROUND_HALF_UP;
    /**
     * 隔的行数
     */
    private static int separateRow = 0;
    private static int separateColumn = 0;

    /**
     * 循环次数，避免死循环
     */
    private static Set cycleSet = new HashSet();

    private static String[][] table;

    /**
     * 左括号
     */
    private static final String LEFT_BRACKET = "(";
    /**
     * 右括号
     */
    private static final String RIGHT_BRACKET = ")";
    /**
     * 冒号
     */
    private static final String COLON = ":";
    /**
     * 逗号
     */
    private static final String COMMA = ",";
    /**
     * RANK公式的排名方式 降序
     */
    private static final String RANK_WAY_DESC = "0";
    /**
     * 转成后缀表达式
     *
     * @param expressionList 前缀表达式
     * @return 后缀表达式
     */
    private static List<String> parseToSuffixExpression(List<String> expressionList) {
        //创建一个栈用于保存操作符
        Stack<String> opStack = new Stack<>();
        //创建一个list用于保存后缀表达式
        List<String> suffixList = new ArrayList<>();
        for (String item : expressionList) {
            //得到数或操作符
            if (isOperator(item)) {
                //是操作符 判断操作符栈是否为空
                if (opStack.isEmpty() || LEFT_BRACKET.equals(opStack.peek()) || priority(item) > priority(opStack.peek())) {
                    //为空或者栈顶元素为左括号或者当前操作符大于栈顶操作符直接压栈
                    opStack.push(item);
                } else {
                    //否则将栈中元素出栈如队，直到遇到大于当前操作符或者遇到左括号时或者其他函数公式
                    while (!opStack.isEmpty() && !LEFT_BRACKET.equals(opStack.peek())) {
                        if (priority(item) <= priority(opStack.peek())) {
                            suffixList.add(opStack.pop());
                        } else {
                            break;
                        }
                    }
                    //当前操作符压栈
                    opStack.push(item);
                }
            } else if (item.matches(REGEX3)) {
                //是数字则直接入队
                suffixList.add(item);
            } else if (LEFT_BRACKET.equals(item) || COLON.equals(item) || COMMA.equals(item) || isFunOperator(item) || item.matches(REGEX2)) {
                //是左括号，压栈
                opStack.push(item);
            } else if (RIGHT_BRACKET.equals(item)) {
                //是右括号 ，将栈中元素弹出入队，直到遇到左括号或者函数公式，左括号出栈，但不入队,函数公式出栈且入队。
                while (!opStack.isEmpty()) {
                    if (LEFT_BRACKET.equals(opStack.peek())) {
                        opStack.pop();
                        break;
                    } else if ("-(".equals(opStack.peek()) || opStack.peek().matches(FORMULA_REGEX)) {
                        suffixList.add(opStack.pop());
                        break;
                    } else {
                        suffixList.add(opStack.pop());
                    }
                }
            } else if (".".equals(item)) {
                suffixList.add(item);
            } else {
                throw new RuntimeException("有非法字符：" + item);
            }
        }
        //循环完毕，如果操作符栈中元素不为空，将栈中元素出栈入队
        while (!opStack.isEmpty()) {
            suffixList.add(opStack.pop());
        }
        return suffixList;
    }

    /**
     * 判断字符串是否为基础操作符，包含逗号
     *
     * @param op 字符
     * @return 是否
     */
    private static boolean isOperator(String op) {
        return "+".equals(op) || "-".equals(op) || "*".equals(op) || "/".equals(op) || COMMA.equals(op);
    }

    /**
     * 判断字符串是否为特殊操作符
     *
     * @param op 字符
     * @return 是否
     */
    private static boolean isFunOperator(String op) {
        return op.matches("^-\\(") || op.matches("^[A-Z]+\\(");
    }

    /**
     * 判断是否为数字
     *
     * @param num 字符
     * @return 是否
     */
    private static boolean isNumber(String num) {
        return num.matches("^-?([0-9]{1,}[.][0-9]*)$") || num.matches("^-?([0-9]{1,})$");
    }

    /**
     * 获取操作符的优先级
     */
    @SuppressWarnings("all")
    private static int priority(String op) {
        if (op.equals("*") || op.equals("/")) {
            return 1;
        } else if (op.equals("+") || op.equals("-")) {
            return 0;
        } else if (op.matches("^[<>]=?|=")) {
            return -1;
        } else if (op.matches(COLON)) {
            return -2;
        } else if (op.matches(COMMA)) {
            return -3;
        } else if (op.matches("\\(") || op.matches("\\)") || isFunOperator(op)) {
            return -4;
        }
        return -5;
    }

    /**
     * 将表达式转为list
     */
    @SuppressWarnings("uncheck")
    private static List<String> expressionToList(String expression) throws Exception {
        expression = expression.replaceAll(" ", "");
        int index = 0;
        List<String> list = new ArrayList<>();
        do {
            char ch = expression.charAt(index);

            // 是操作符 “+-*/ ,:” , -(不是负号),  冒号（58）；
            // Character.isDigit(expression.charAt(index-1) 表示-号前面为数字的时候
            // expression.charAt(index-1)==')') 表示-号前面为)的时候
            if ((ch != 45 && ch != 46 && ch <= 47) || ch == 58 || (index != 0 && ch == 45 && (Character.isDigit(expression.charAt(index - 1)) || expression.charAt(index - 1) == ')'))) {
                //是操作符，直接添加至list中
                index++;
                list.add(ch + "");
            } else {
                // 特殊公式
                String a = expression.substring(index);
                Pattern pattern = Pattern.compile(REGEX);
                Matcher matcher = pattern.matcher(a);
                if (matcher.find()) {
                    String str = matcher.group(0);
                    // 含百分比的数值
                    if (str.endsWith("%")) {
                        str = new BigDecimal(str.substring(0, str.length() - 1)).divide(BigDecimal.valueOf(100)) + "";
                    }
                    list.add(str);
                    index = index + matcher.group(0).length();
                    continue;
                }

                // 是数字和负数,判断多位数的情况 注：45 负号， 46 点
                String str = "";
                while (index < expression.length() && (expression.charAt(index) > 47 && expression.charAt(index) < 58 || expression.charAt(index) == 46
                        || (expression.charAt(index) == 45 && !(index != 0 && (Character.isDigit(expression.charAt(index - 1)) || expression.charAt(index - 1) == ')'))))) {
                    str += expression.charAt(index);
                    index++;
                }
                if ("-".equals(str)) {
                    // 负号
                    list.add("-(");
                    index++;
                } else {
                    // 匹配不到，则抛异常
                    if ("".equals(str)) {
                        System.out.println("无法解析公式含义");
                        throw new Exception(ErrorMessageEnum.ERROR.toString());
                    }
                    list.add(str);
                }
            }
        } while (index < expression.length());
        return list;
    }

    /**
     * 根据后缀表达式list计算结果
     */
    private static String calculate(List<String> list) throws Exception {
        Stack<String> stack = new Stack<>();
        for (int i = 0; i < list.size(); i++) {
            String item = list.get(i);
            if (item.matches(REGEX3)) {
                // 是数字
                stack.push(item);
            } else if (item.matches("[<>]=?|=")) {
                getCompare(stack, item);
            } else {
                switch (item) {
                    case "-(":
                    case "ABS(":
                        getNotOrAbs(stack, item);
                        break;
                    case COMMA:
                    case COLON:
                        stack.push(item);
                        break;
                    case "MIN(":
                    case "MAX(":
                        getMostValue(stack, item);
                        break;
                    case "IF(":
                        getIf(stack);
                        break;
                    case "AND(":
                    case "OR(":
                        getAndOr(stack, item);
                        break;
                    case "SUM(":
                        getSum(stack);
                        break;
                    case "RANK(":
                        getRank(stack);
                        break;
                    default:
                        //是操作符，取出栈顶两个元素
                        BigDecimal res;
                        removeComma(stack);
                        BigDecimal num2 = new BigDecimal(explain(stack.pop()));
                        //System.out.print(num2);
                        removeComma(stack);
                        BigDecimal num1 = new BigDecimal(explain(stack.pop()));
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
                                try {
                                    res = num1.divide(num2, DIGITS, ROUNDING_MODE);
                                } catch (ArithmeticException e) {
                                    throw new Exception(ErrorMessageEnum.BY_ZERO_ERROR.toString());
                                }
                                break;
                            default:
                                throw new RuntimeException("运算符错误：" + item);
                        }
                        stack.push(res + "");
                }
            }

        }
        return explain(stack.pop());
    }

    /**
     * 过滤栈顶的逗号 返回非逗号的栈顶值
     *
     * @param stack 栈
     */
    private static void removeComma(Stack stack) throws Exception {
        try {
            if (COMMA.equals(stack.peek())) {
                stack.pop();
            }
        } catch (EmptyStackException e) {
            throw new Exception(ErrorMessageEnum.FORMULA_ERROR.toString());
        }
    }

    /**
     * 判断   公式，计算
     * 数值，返回
     *
     * @param s
     * @return
     * @throws Exception
     */
    private static String explain(String s) throws Exception {
        return explain(s, false);
    }

    /**
     * 判断   公式，计算
     * 数值，返回
     *
     * @param s
     * @param isReturnNull 当 s 等于 null， 是否返回 null，否，返回 零
     * @return
     */
    private static String explain(String s, Boolean isReturnNull) throws Exception {
        // 单元格没有数据，赋值为0
        if (s == null) {
            if (isReturnNull) {
                return null;
            }
            return "0";
        } else if ("true".equals(s) || "false".equals(s)) {
            return s;
        }
        // 是否为 错误信息，如果是错误的传递
        for (ErrorMessageEnum e : ErrorMessageEnum.values()) {
            if (e.toString().equals(s)) {
                return e.toString();
            }
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
        // 判断是什么类型
        if (s.matches(CEIL_REGEX)) {
            // 单元格名称
            try {
                row = getRow(s);
                column = getColumn(s);
                s = table[row][column];
                // 递归
                try {
                    a = explain(s, isReturnNull);
                } catch (Exception e) {
                    if (!table[row][column].contains("#")) {
                        System.out.println(table[row][column] + "公式：" + e.getMessage());
                        table[row][column] = e.getMessage();
                    }
                    throw new Exception(e.getMessage());
                }
                table[row][column] = table[row][column] == null ? null : a;
            } catch (NumberFormatException e) {
                System.out.println("数字类型错误");
                throw new Exception(ErrorMessageEnum.NAME_ERROR.toString());
            } catch (ArrayIndexOutOfBoundsException e) {
                if (row >= 0) {
                    if (isReturnNull) {
                        a = null;
                    } else {
                        a = "0";
                    }
                } else {
                    System.out.println("数组越界");
                    throw new Exception(ErrorMessageEnum.NAME_ERROR.toString());
                }
            }
        } else if (s.matches(NUMBER_REGEX)) {
            // 数值
            a = s;
        } else {
            // 公式
            if (cycleSet.contains(s)) {
                cycleSet.clear();
                // 程序死循环
//                return ErrorMessageEnum.ENDLESS_LOOP.toString();
                throw new Exception(ErrorMessageEnum.ENDLESS_LOOP.toString());
            }
            cycleSet.add(s);
            a = formula(s);
            cycleSet.clear();
        }

        return a;
    }

    /**
     * 转换成boolean，是数值且不为零 或者 为‘ture’ 为真
     *
     * @param s
     * @return
     */
    private static boolean parseBoolean(String s) {
        return (s != null) && (s.equalsIgnoreCase("true") || (s.matches("\\d+") && !s.equals("0")));
    }

    /**
     * 公式计算
     *
     * @param expression
     * @return
     */
    private static String formula(String expression) throws Exception {
        String result;
        List<String> expressionList = expressionToList(expression);
        //System.out.println("中缀表达式转为list结构="+expressionList);
        // 将中缀表达式转换为后缀表达式
        List<String> suffixList = parseToSuffixExpression(expressionList);
        //System.out.println("对应的后缀表达式列表结构="+suffixList);

        // 根据后缀表达式计算结果
        String calculateResult;
        try {
            calculateResult = calculate(suffixList);
        } catch (EmptyStackException e) {
            System.out.println("EmptyStackException");
            throw new Exception(ErrorMessageEnum.ERROR.toString());
        } catch (Exception e) {
            throw new Exception(e.getMessage());
        }
        if ("true".equals(calculateResult) || "false".equals(calculateResult)) {
            result = calculateResult;
        } else {
            result = new BigDecimal(calculateResult).stripTrailingZeros().toPlainString();
        }
        return result;
    }

    /**
     * 该单元格的处于数组的第几行
     *
     * @param s
     * @return
     */
    private static int getRow(String s) {
        return Integer.parseInt(s.substring(1)) - 1 - separateRow;
    }

    /**
     * 该单元格的处于数组的第几列
     *
     * @param s
     * @return
     */
    private static int getColumn(String s) {
        return s.charAt(0) - 65 - separateColumn;
    }

    private static BigDecimal getSum(String start, String end) throws Exception {
        int x1 = getRow(start);
        int y1 = getColumn(start);
        int x2 = getRow(end);
        int y2 = getColumn(end);
        return getSum(x1, y1, x2, y2);
    }

    private static BigDecimal getSum(int x1, int y1, int x2, int y2) throws Exception {
        BigDecimal sum = BigDecimal.ZERO;
        for (int i = x1; i <= x2; i++) {
            for (int j = y1; j <= y2; j++) {
                try {
                    sum = sum.add(new BigDecimal(explain(table[i][j])));
                } catch (IndexOutOfBoundsException e) {
                    // 当i < 0(即：单元名称不能用字母加小于等于零的数值)
                    assert (i > table.length && j > table[i].length) : (char)(j  + separateColumn + 65) + "" + (i + separateRow + 1)  + " 单元格未赋值";
                    System.out.println("单元名称不能用字母加小于等于零的数值");
                    throw new Exception(ErrorMessageEnum.NAME_ERROR.toString());
                }
            }
        }
        return sum;
    }

    private static BigDecimal getAvg(int x1, int y1, int x2, int y2) throws Exception {
        int count = Math.abs((x2 - x1 + 1) * (y2 - y1 + 1));
        BigDecimal avg = getSum(x1, y1, x2, y2).divide(BigDecimal.valueOf(count), DIGITS, ROUNDING_MODE);
        return avg;
    }

    /**
     * 获取该区域的所有单元格的值
     *
     * @param start
     * @param end
     * @return
     * @throws Exception
     */
    private static List<String> getSelectedCeilValueList(String start, String end) throws Exception {
        List<String> lists = new LinkedList<>();
        int row = getRow(start);
        int column = getColumn(start);
        int row2 = getRow(end);
        int column2 = getColumn(end);
        for (int x = row; x <= row2; x++) {
            for (int y = column; y <= column2; y++) {
                // 该单元格没有赋初值，则不参与排名计算
                try {
                    if (table[x][y] != null) {
                        lists.add(explain(table[x][y]));
                    }
                } catch (IndexOutOfBoundsException e) {
                    System.out.println("索引越界");
                }
            }
        }
        return lists;
    }

    /**
     * 取反 或者 取绝对值
     * @param stack
     * @param item
     * @throws Exception
     */
    private static void getNotOrAbs(Stack<String> stack, String item) throws Exception {
        // 是操作符，取出栈顶一个元素
        BigDecimal num = new BigDecimal(explain(stack.pop()));
        BigDecimal res;
        if (item.equals("-(")) {
            res = BigDecimal.ZERO.subtract(num);
        } else {
            res = num.abs();
        }
        stack.push(res + "");
    }

    /**
     * 最值
     * @throws Exception
     */
    private static void getMostValue(Stack<String> stack, String item) throws Exception {
        List<String> lists = new LinkedList<>();
        BigDecimal mostValue = BigDecimal.ZERO;
        // 逗号个数+1
        int count = 1;
        int compareNum = -1;
        if ("MAX(".equals(item)) {
            compareNum = 1;
        }

        for (int j = 0; j < count; ) {
            if (COMMA.equals(stack.peek())) {
                count++;
                stack.pop();
            } else {
                lists.add(stack.pop());
                j++;
            }
        }
        for (int j = lists.size() - 1; j >= 0; j--) {
            BigDecimal temp = new BigDecimal(explain(lists.get(j)));
            if (j == lists.size() - 1) {
                mostValue = temp;
            } else if (temp.compareTo(mostValue) == compareNum) {
                mostValue = temp;
            }
        }
        stack.push(mostValue + "");
    }

    private static void getIf(Stack<String> stack) throws Exception {
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
    }

    private static void getAndOr(Stack<String> stack, String item) {
        boolean flag = "AND(".equals(item) ? true : false;
        List<String> lists = new LinkedList<>();
        // 逗号个数+1
        int count = 1;
        for (int j = 0; j < count; ) {
            if (COMMA.equals(stack.peek())) {
                count++;
                stack.pop();
            } else {
                lists.add(stack.pop());
                j++;
            }
        }
        for (int j = lists.size() - 1; j >= 0; j--) {
            if (flag ? !parseBoolean(lists.get(j)) : parseBoolean(lists.get(j))) {
                flag = !flag;
                break;
            }
        }
        stack.push(flag + "");
    }

    /**
     * 比较函数（大于等于小于）计算
     * @param stack
     * @param item
     * @throws Exception
     */
    private static void getCompare(Stack<String> stack, String item) throws Exception {
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
    }
    private static void getSum(Stack<String> stack) throws Exception {
        BigDecimal sum = BigDecimal.ZERO;
        List<String> lists = new LinkedList<>();
        // 逗号个数+1
        int count = 1;
        for (int j = 0; j < count; ) {
            if (COMMA.equals(stack.peek())) {
                count++;
                stack.pop();
            } else if (COLON.equals(stack.peek())) {
                stack.pop();
                String end = stack.pop();
                String start = stack.pop();
                if (!start.matches(CEIL_REGEX) || !end.matches(CEIL_REGEX)) {
                    System.out.println("SUM公式错误");
                    throw new Exception(ErrorMessageEnum.NAME_ERROR.toString());
                }
                BigDecimal sumTmp = getSum(start, end);
                lists.add(sumTmp.toString());
                j++;
            } else {
                lists.add(stack.pop());
                j++;
            }
        }
        for (int j = lists.size() - 1; j >= 0; j--) {
            sum = sum.add(new BigDecimal(explain(lists.get(j))));
        }

        stack.push(sum + "");
    }

    private static void getRank(Stack<String> stack) throws Exception {
        // 排位方式
        String rankWay = "";
        List<String> lists = new LinkedList<>();
        // 逗号个数+1
        int count = 1;
        for (int j = 0; j < count; ) {
            if (COMMA.equals(stack.peek())) {
                count++;
                stack.pop();
            } else if (COLON.equals(stack.peek())) {
                stack.pop();
                String end = stack.pop();
                String start = stack.pop();
                if (!start.matches(CEIL_REGEX) || !end.matches(CEIL_REGEX)) {
                    System.out.println("RANK公式错误");
                    throw new Exception(ErrorMessageEnum.NAME_ERROR.toString());
                }
                List<String> ceilList = getSelectedCeilValueList(start, end);
                lists.addAll(ceilList);

                j++;
            } else {
                lists.add(stack.pop());
                j++;
            }
        }

        String firstStr = lists.get(lists.size() - 1);
        lists.remove(lists.size() - 1);
        String endStr = lists.get(0);
        if (!endStr.matches(CEIL_REGEX)) {
            rankWay = lists.get(0);
            lists.remove(0);
        }

        int rank = getRank(firstStr, lists, rankWay);
        stack.push(rank + "");
    }

    /**
     * 获取某个单元格（或数值）在某个范围内的排名，该单元格（或数值）必须在该范围里
     *
     * @param str     某个单元格（或数值）
     * @param list    引用（某些单元格）
     * @param rankWay 排位方式，为空或为0，降序， 其他，升序。
     * @return 排名
     */
    private static Integer getRank(String str, List<String> list, String rankWay) throws Exception {
        str = explain(str, true);
        if (str == null) {
            throw new Exception(ErrorMessageEnum.N_A_ERROR.toString());
        }
        if (!str.matches(NUMBER_REGEX)) {
            System.out.println("RANK所选单元格错误");
            throw new Exception(str);
        }
        BigDecimal tmp = new BigDecimal(str);
        // 比较大小，根据排位方式，选择大于或小于来进行排名比较, -1 降序，1 升序。
        int compareNum = -1;
        if (rankWay != null && !"".equals(rankWay) && !RANK_WAY_DESC.equals(rankWay)) {
            compareNum = 1;
        }
        // 该单元格值是否在引用中
        boolean isExist = false;
        // 排名
        int rank = 1;
        for (int i = 0; i < list.size(); i++) {
            String searchResult = explain(list.get(i), true);
            if (searchResult == null) {
                continue;
            }
            BigDecimal currentDecimal = new BigDecimal(searchResult);
            if (tmp.equals(currentDecimal)) {
                isExist = true;
            }
            if (tmp.compareTo(currentDecimal) == compareNum) {
                rank += 1;
            }
        }
        if (!isExist) {
            System.out.println("排名：数值不在该范围内");
            throw new Exception(ErrorMessageEnum.N_A_ERROR.toString());
        }
        return rank;
    }

    /**
     * 验证公式
     * 括号对称验证（左右括号是否成对），某些公式是否连续（如：加减）
     * 是否包含规定之外的函数名称（需要么？）
     * @param clac
     * @throws Exception
     */
    private static void verifyFormula(String clac) throws Exception {
        if (!isValid(clac)) {
            throw new Exception(ErrorMessageEnum.NAME_ERROR.toString());
        }
//        clac
    }


    /**
     * 括号匹配验证
     * @param s
     * @return
     */
    private static boolean isValid(String s) {
        if (s == null || s.length() == 0) {
            return false;
        }
        Stack stack = new Stack();
        for(int i = 0; i<s.length(); i++) {
            char letter = s.charAt(i);
            switch(letter) {
                case '(':
                case '{':
                case '[':
                    stack.push(letter);
                    break;
                case ')':
                    if(!"(".equals(stack.pop())) {
                        return false;
                    }
                case '}':
                    if(!"{".equals(stack.pop())) {
                        return false;
                    }
                case ']':
                    if(!"[".equals(stack.pop()))  {
                        return false;
                    }
                default:
                    break;
            }
        }
        return stack.size() == 0;
    }

    /**
     * excel 计算公式
     * 行列偏移量都为零
     * @param strArray
     * @return
     */
    public static String[][] excel(String[][] strArray) {
        return excel(strArray, 0, 0);
    }

    /**
     * excel计算公式
     * @param table 计算的数组
     * @param separateRow 公式的 单元格名称，跟现在的数组的行数差
     * @param separateColumn 公式的 单元格名称，跟现在的数组的列数差
     * @return String[][]
     */
    public static String[][] excel(String[][] table, int separateRow, int separateColumn) {
        ExcelCalculation.table = table;
        ExcelCalculation.separateRow = separateRow;
        ExcelCalculation.separateColumn = separateColumn;
        for (int i = 0; i < table.length; i++) {
            for (int j = 0; j < table[i].length; j++) {
                try {
//                    if (table[i][j].contains("=RANK(B5,(A2,A1, B2, C3,A4), 0)")) {
//                        System.out.println(table[i][j]);
//                    }

                    // 验证公式
                    verifyFormula(table[i][j]);
                    table[i][j] = table[i][j] == null ? null : explain(table[i][j]);
                } catch (Exception e) {
                    // 是否是已经处理过的单元格，“#”为异常消息ErrorMessageEnum：名称都以“#”开头
                    if (!table[i][j].contains("#")) {
                        System.out.println(table[i][j] + "公式~~：" + e.getMessage());
                        table[i][j] = e.getMessage();
                    }
                }
            }
        }
        return table;
    }

}

