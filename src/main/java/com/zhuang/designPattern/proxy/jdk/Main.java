package com.zhuang.designPattern.proxy.jdk;

import java.lang.reflect.Proxy;

public class Main {
    public static void main(String[] args) {
        IService service = new Service();
        MyInvocationHandler myInvocationHandler = new MyInvocationHandler(service);
        IService serviceProxy = (IService) Proxy.newProxyInstance(service.getClass().getClassLoader(), service.getClass().getInterfaces(), myInvocationHandler);
        serviceProxy.add();
        System.out.println();
        serviceProxy.update();
    }
}
