package com.zhuang.classLoad;

public class ClassLoadTree {
    public static void main(String[] args) {
        ClassLoader classLoader = ClassLoadTree.class.getClassLoader();
        while (classLoader != null) {
            System.out.println(classLoader);
            classLoader = classLoader.getParent();
        }
    }
}
