package com.zhuang.algorithm;

/**
 * 最大序列之和
 * 溢出问题
 * @author zhuang
 */
public class MaximalSequence {
    public static void main(String[] args) {
        int[] a = {
          -1, -2, -3, -4, -5, 2
        };

        System.out.println(maxValue(a, 0, a.length - 1));
    }

    public static int maxValue(int[] a, int left, int right) {
        if (left == right) {
            return a[left];
//            if (a[left] > 0) {
//                return a[left];
//            }
//            return 0;
        }

        int center = (left + right) / 2;
        int maxLeftSum = maxValue(a, left, center);
        int maxRightSum = maxValue(a, center + 1, right);

        int maxLeftBorderSum = a[center], leftBorderSum = 0;
        for (int i = center; i >= left; i--) {
            leftBorderSum += a[i];
            if (leftBorderSum > maxLeftBorderSum) {
                maxLeftBorderSum = leftBorderSum;
            }
        }

        int maxRightBorderSum = a[center + 1], rightBorderSum = 0;
        for (int i = center + 1; i <= right; i++) {
            rightBorderSum += a[i];
            if (rightBorderSum > maxRightBorderSum) {
                maxRightBorderSum = rightBorderSum;
            }
        }

        return max3(maxLeftSum, maxRightSum, maxLeftBorderSum + maxRightBorderSum);
    }

    public static int max3(int a, int b, int c) {
        return a > b ? (a > c ? a : c) : (b > c ? b : c);
    }
}
