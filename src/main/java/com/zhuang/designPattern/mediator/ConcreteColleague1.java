package com.zhuang.designPattern.mediator;

public class ConcreteColleague1 extends Colleague {
    public ConcreteColleague1(Mediator mediator) {
        super(mediator);
    }

    public void send(String message) {
        notify(message);
    }

    public void notify(String message) {
        System.out.println("同事1得到消息" + message);
    }
}
