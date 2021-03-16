package com.zhuang.designPattern.mediator;

/**
 * 抽象中介者
 */
public abstract class Mediator {
    public abstract void send(String message, Colleague colleague);
}
