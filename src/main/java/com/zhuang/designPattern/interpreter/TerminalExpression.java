package com.zhuang.designPattern.interpreter;

public class TerminalExpression extends AbstractExpression {
    @Override
    public void interpreter(Context context) {
        System.out.println("终端解释器");
    }
}
