package com.zhuang.designPattern.visitor;

public class ConcreteElementB extends Elemnet {
    @Override
    public void accept(Visitor visitor) {
        visitor.visitorConcreteElementB(this);
    }

    public void operationB() {

    }
}
