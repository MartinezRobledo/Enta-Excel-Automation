package com.automationanywhere.botcommand.utilities;

import com.jacob.com.Variant;
import com.jacob.com.SafeArray;

public class SafeArrayConverter {

    public static Object[][] convertSafeArrayTo2DArray(Variant variant) {
        if (variant == null || variant.isNull()) {
            return new Object[0][0];
        }

        SafeArray safeArray = variant.toSafeArray();

        int lBoundRow = safeArray.getLBound(1);
        int uBoundRow = safeArray.getUBound(1);
        int lBoundCol = safeArray.getLBound(2);
        int uBoundCol = safeArray.getUBound(2);

        int rows = uBoundRow - lBoundRow + 1;
        int cols = uBoundCol - lBoundCol + 1;

        Object[][] array = new Object[rows][cols];

        for (int i = lBoundRow; i <= uBoundRow; i++) {
            for (int j = lBoundCol; j <= uBoundCol; j++) {
                array[i - lBoundRow][j - lBoundCol] = safeArray.getVariant(i, j).toJavaObject();
            }
        }

        return array;
    }
}
