package com.zhuang.designPattern.proxy.cglib;

/**
 * @author typ
 *
 */
public class Main {
    public static void main(String[] args) {
        CglibProxy proxy = new CglibProxy();
        // base为生成的增强过的目标类
        Base base = Factory.getInstance(proxy);
        base.add();
    }
}