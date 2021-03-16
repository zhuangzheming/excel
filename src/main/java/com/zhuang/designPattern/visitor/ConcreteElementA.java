package com.zhuang.designPattern.visitor;

public class ConcreteElementA extends Elemnet {
    @Override
    public void accept(Visitor visitor) {
        visitor.visitorConcreteElementA(this);
    }

    public void operationA() {

    }
}
