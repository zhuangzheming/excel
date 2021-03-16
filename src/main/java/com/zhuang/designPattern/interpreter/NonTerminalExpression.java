package com.zhuang.designPattern.interpreter;

public class NonTerminalExpression extends AbstractExpression {
    @Override
    public void interpreter(Context context) {
        System.out.println("非终端解释器");
    }
}
