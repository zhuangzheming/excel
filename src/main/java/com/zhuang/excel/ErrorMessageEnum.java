package com.zhuang.excel;

/**
 * 错误种类
 * @author zhuang
 */
public enum ErrorMessageEnum {
    /**
     * 函数名称错误或单元格名称错误
     */
    NAME_ERROR("#NAME?"),
    /**
     * 死循环
     */
    ENDLESS_LOOP("#ENDLESS_LOOP?"),
    /**
     * 函数公式错误
     */
    FORMULA_ERROR("#FORMULA_ERROR?"),
    /**
     * 除数不能为零
     */
    BY_ZERO_ERROR("#DIV/0!"),
    /**
     * 当在函数或公式中没有可用数值时，产生错误值#N/A。
     */
    N_A_ERROR("#N/A"),
    /**
     * 其他错误
     */
    ERROR("#ERROR?");

    private final String text;
    ErrorMessageEnum(final String text){
        this.text=text;
    }
    @Override
    public String toString(){
        return text;
    }
}