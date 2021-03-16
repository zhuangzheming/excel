package com.zhuang.designPattern.interpreter;

import java.util.LinkedList;
import java.util.List;

/**
 * 解释器模式
 */
public class Main {
    public static void main(String[] args) {
        Context context = new Context();
        List<AbstractExpression> list = new LinkedList<>();
        list.add(new TerminalExpression());
        list.add(new TerminalExpression());
        list.add(new NonTerminalExpression());
        list.add(new TerminalExpression());

        for (AbstractExpression exp :
                list) {
            exp.interpreter(context);
        }
    }
}
