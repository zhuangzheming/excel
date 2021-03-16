package com.zhuang.designPattern.mediator;

/**
 * 中介者模式
 * @author zhuang
 */
public class Main {
    public static void main(String[] args) {
        ConcreteMediator mediator = new ConcreteMediator();

        // 让两个具体同事类 认识中介者对象
        ConcreteColleague c = new ConcreteColleague(mediator);
        ConcreteColleague1 c1 = new ConcreteColleague1(mediator);

        // 让中介者 认识各个具体同事类对象
        mediator.setConcreteColleague(c);
        mediator.setConcreteColleague1(c1);

        // 具体同事类对象发送消息 都通过中介者转发
        c.send(" 你好，同事1");
        c1.send(" 你好，同事");
    }
}
