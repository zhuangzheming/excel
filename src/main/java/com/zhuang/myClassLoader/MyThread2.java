package com.zhuang.myClassLoader;

public class MyThread2  extends  Thread {
    public MyThread2() {}

    @Override
    public void run() {
        System.out.println("MyThread2: " + Thread.currentThread().getName());

        System.out.println("My2 Context Class Loader: " + Thread.currentThread().getContextClassLoader());
        try {
            Test.useClassLoader();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
