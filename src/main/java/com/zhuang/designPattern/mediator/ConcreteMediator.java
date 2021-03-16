package com.zhuang.designPattern.mediator;

/**
 * 中介者的具体实现类
 * @author zhuang
 */
public class ConcreteMediator extends Mediator {
    private ConcreteColleague concreteColleague;
    private ConcreteColleague1 concreteColleague1;

    public void setConcreteColleague(ConcreteColleague concreteColleague) {
        this.concreteColleague = concreteColleague;
    }

    public void setConcreteColleague1(ConcreteColleague1 concreteColleague1) {
        this.concreteColleague1 = concreteColleague1;
    }

    @Override
    public void send(String message, Colleague colleague) {
        if (colleague == concreteColleague) {
            concreteColleague.send(message);
        } else {
            concreteColleague1.send(message);
        }
    }
}
