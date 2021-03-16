package com.zhuang.myClassLoader;

public class MyThread implements Runnable {
    @Override
    public void run() {
        System.out.println("MyThread: " + Thread.currentThread().getName());
        System.out.println("My Context Class Loader: " + Thread.currentThread().getContextClassLoader());
        try {
            Test.useClassLoader();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
