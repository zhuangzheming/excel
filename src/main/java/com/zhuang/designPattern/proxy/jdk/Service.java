package com.zhuang.designPattern.proxy.jdk;

public class Service implements IService {

    @Override
    public void add() {
        System.out.println("service add <<");
    }

    @Override
    public void update() {
        System.out.println("service update <<");
    }
}
