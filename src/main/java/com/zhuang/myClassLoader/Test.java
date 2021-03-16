package com.zhuang.myClassLoader;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.HashMap;

public class Test {
    private static String classPath = "D:\\project\\ME\\excel\\src\\main\\java\\com\\zhuang\\myClassLoader\\OutputMessage.class";
    private static MyClassLoader myClassLoader = new MyClassLoader(classPath);
    public static void main(String[] args) throws Exception {
        Thread thread = new Thread(new MyThread());
        thread.start();
         Thread.sleep(1500);
        MyThread2 myThread2 = new MyThread2();
        myThread2.setName("weizhi");
        myThread2.start();
//        String url = "jdbc:mysql://localhost:3306/testdb";
//// 通过java库获取数据库连接
//        Connection conn = java.sql.DriverManager.getConnection(url, "name", "password");


        //这个类class的路径
//        String classPath = "/Users/mac/CommProjects/ClassLoaderDemo/src/com/xuhuawei/classloaderdemo/Log.class";
//        String classPath = "D:\\project\\ME\\excel\\src\\main\\java\\com\\zhuang\\myClassLoader\\OutputMessage.class";
//
//        MyClassLoader myClassLoader = new MyClassLoader(classPath);
//        //类的全称
//        String packageNamePath = "com.zhuang.myClassLoader.OutputMessage";
//
//        //加载Log这个class文件
//        Class<?> OutputMessage = myClassLoader.loadClass(packageNamePath);
//
//        System.out.println("类加载器是:" + OutputMessage.getClassLoader());
//
//        //利用反射获取main方法
//        Method method = OutputMessage.getDeclaredMethod("print", null);
//        Object object = OutputMessage.newInstance();
////        String[] arg = {"ad"};
//        method.invoke(object, null);
    }

    public  static  Object  useClassLoader() throws Exception {
        //这个类class的路径



        //类的全称
        String packageNamePath = "com.zhuang.myClassLoader.OutputMessage";

        //加载Log这个class文件
        Class<?> OutputMessage = myClassLoader.loadClass(packageNamePath);

        System.out.println("类加载器是:" + OutputMessage.getClassLoader());

        //利用反射获取main方法
        Method method = OutputMessage.getDeclaredMethod("print", null);
        Object object = OutputMessage.newInstance();
//        String[] arg = {"ad"};
        method.invoke(object, null);
        return  myClassLoader;
    }
}
