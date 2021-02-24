package com.zhuang.excel;

/**
 * @author zhuang
 */
public class Main {
    public static void main(String[] args) {
        //    static String[][] table = {
//            {"=SUM(1,B1:C1,3)", "1", "2", "=SUM(A1,A1:C1,B1:C1, C1)"}
//    };
//    =COUNT(A1,A2) 为null或不存在（包括未赋值） 记为零
//  AVERAGE(A1:A5, 5)
//  AVERAGE(A1:A5)
//    =RANK(E54,(E54,I54,K54,L54),0)
//    =RANK(F24,$F$24:$N$24,0)     最后一个参数可以省略， =RANK(F24,$F$24:$N$24)
//    =SUM(E2,F2,G2)
//    =SUM(E2:G2)
//    static String[][] table = {
//        {"191.087684"},
//        {"248.880452"},
//        {"16"},
//        {"16.5"},
//        {"=F3/F5"},
//        {"=F4/F6"},
//        {"=F8/F7-1"},
//        {"=IF(F9<-15%,90%,IF(AND(F9>=-15%,F9<0%),95%,IF(AND(F9>=0%,F9<3%),100%,IF(AND(F9>=3%,F9<5%),104%,IF(AND(F9>=5%,F9<8%),108%,110%)))))"},
//        {"=50*F10"},
//        {"3719"},
//        {"3747"},
//        {"=ABS((F12-F13)/F13)"},
//        {"=IF(F13<100,10%,IF(AND(F13>=100,F13<1000,ABS(F12-F13)<=20),5%,IF(AND(F13>=1000,ABS(F12-F13)<=30),2%,0)))"},
//        {"=IF(F14>F15,0%,IF(AND(F14<=F15,F14>0.8*F15),100%,MIN(100%+(ABS(F15-F14)/(F15*20%))*2%,110%)))"},
//        {"=50*F16"},
//        {null},
//        {null},
//        {null},
//        {null},
//        {null},
//        {"=F11+F17+F22"},
//        {"=F23/100", "1.05","1.08","1.07","1.01","1.02","1.07","1.02"},//,"1.082"
//        {"=RANK(F24,$F$24:$N$24,0)"},
//        {"=IF(OR(F25=1,F25=2),1,IF(OR(F25=3,F25=4),0.5,IF(F25=5,0,IF(OR(F25=6,F25=7),-0.5,-1))))"},
//        {"=(F24-1)*1*10"},
//        {"=F26*0.6+F27*0.4"},
//        {null}
//    };
        String[][] table = {
                {null, "=B2+2", "-3", "=D1"},
                {"=A1", null, "=B2", "=E1"},
                {"3.3", "=IF(AND(A3<A4, 1%<2, 3 < 5 * 103%), MIN(A3, A4), 0%)"},
                {"=A6+A5", "=SUM(A1:A5)", "=RANK(A1,A1:B5)", "=RANK(A2,A1:B5, 1)", "=RANK(B5,(C3,A4), 0)", "=RANK(C1,(A2,A1, B2, C1,A4), 0)", "=RANK(0,(A2,A3,A4), 1)"},
                {"5", "=SUM(A1:C3)", "=(SUM(A7,A3,C1)-A7)*2", "=SUM(A7+A8)", "=SUM(A1, A1:C1, B1:C1, C1)"},
                {"6", "-(-3/0)", "-(-(3)*3+2)+4/2", "=B6", "=ABS(-3.4343+MAX(C1,A3,A1))", "-3/2"},
                {"=A6*A3", "=10%/3 + IF(OR(A7<1,A1<0,A3<4),IF(AND(A5<6,A7<3), MAX(A3,3), MIN(A3, 3)), MIN(3,5%))"},
                {"=IF(OR(1<2, 3>3), 1, 2)", "IF(79<(4+4), 3, 4)", "79<(4+4)", "=RANK(B1,(A1:A2, C1, B1:B3, A3, A4), 0)"}
        };

        table = ExcelCalculation.excel(table);
//        table = ExcelCalculation.excel(table, 2, 5);

        System.out.println("table = {");
        for (int i = 0; i < table.length; i++) {
            for (int j = 0; j < table[i].length; j++) {
                System.out.print(table[i][j] + ",\t");
            }
            System.out.println();
        }
        System.out.println("}");
    }
}
