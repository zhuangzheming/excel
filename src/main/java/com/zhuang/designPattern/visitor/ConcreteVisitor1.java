package com.zhuang.designPattern.visitor;

public class ConcreteVisitor1 extends Visitor {
    @Override
    public void visitorConcreteElementA(ConcreteElementA concreteElementA) {
        System.out.println(concreteElementA.getClass().getSimpleName()  + "被" + this.getClass().getSimpleName() + "访问");
    }

    @Override
    public void visitorConcreteElementB(ConcreteElementB concreteElementB) {
        System.out.println(concreteElementB.getClass().getSimpleName()  + "被" + this.getClass().getSimpleName()  + "访问");
    }
}
