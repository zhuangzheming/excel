package com.zhuang.designPattern.visitor;

import java.util.LinkedList;
import java.util.List;

public class ObjectStructure {
    private List<Elemnet> elemnets = new LinkedList<>();

    public void attach(Elemnet elemnet) {
        elemnets.add(elemnet);
    }

    public void detach(Elemnet elemnet) {
        elemnets.remove(elemnet);
    }

    public void accept(Visitor visitor) {
        for (Elemnet e :
                elemnets) {
            e.accept(visitor);
        }
    }
}
