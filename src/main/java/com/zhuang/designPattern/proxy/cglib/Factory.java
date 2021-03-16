package com.zhuang.designPattern.proxy.cglib;

import net.sf.cglib.proxy.Enhancer;

/**
 * 工厂类，生成增强过的目标类（已加入切入逻辑）
 *
 * @author typ
 *
 */
public class Factory {
    /**
     * 获得增强之后的目标类，即添加了切入逻辑advice之后的目标类
     *
     * @param proxy
     * @return
     */
    public static Base getInstance(CglibProxy proxy) {
        Enhancer enhancer = new Enhancer();
        enhancer.setSuperclass(Base.class);
        // 回调方法的参数为代理类对象CglibProxy，最后增强目标类调用的是代理类对象CglibProxy中的intercept方法
        enhancer.setCallback(proxy);
        // 此刻，base不是单纯的目标类，而是增强过的目标类
        Base base = (Base) enhancer.create();
        return base;
    }
}
