package com.automationanywhere.botcommand.utilities;

import com.jacob.com.Variant;
import com.jacob.com.SafeArray;
import com.jacob.com.ComFailException;

public class VariantUtils {

    public static SafeArray convertVariantToSafeArray(Variant variant) {
        try {
            // Intentamos convertir el Variant a un SafeArray
            return variant.toSafeArray();
        } catch (ComFailException e) {
            // Si ocurre una excepción, significa que no es un SafeArray
            return null;
        }
    }
}
