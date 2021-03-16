package com.zhuang.designPattern.decorator;

public class ConcreteDecoratorB extends Decorator {
    @Override
    public void operation() {
        super.operation();
        addedBehavior();
        System.out.println("具体B类操作");
    }
    public void addedBehavior() {
        System.out.println("B 类行为");
    }
}
