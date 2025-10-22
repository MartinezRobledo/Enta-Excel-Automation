package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelHelpers;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.SafeArray;
import com.jacob.com.Variant;

import java.math.BigDecimal;
import java.util.*;

@BotCommand
@CommandPkg(
        label = "Group & Sum by Reference",
        name = "groupSumByReference",
        description = "Groups by a reference column and sums a value column. Reads from source workbook and writes to destination workbook.",
        icon = "excel.svg"
)
public class GroupSumByReference {

    @Execute
    public void action(
            // ===== SOURCE WORKBOOK / SHEET =====
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Source Workbook Session")
            @NotEmpty @SessionObject ExcelSession sourceExcelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select source sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String selectSourceSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Source Sheet Name")
            String sourceSheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Source Sheet Index (1-based)")
            @NumberInteger @GreaterThanEqualTo("1") Double sourceSheetIndex,

            // ===== COLUMN SELECTION =====
            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "By Letter (A,B,...)", value = "letter")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "By Header name", value = "header"))
            })
            @Pkg(label = "Select columns by", default_value = "letter", default_value_type = DataType.STRING)
            @SelectModes String selectColsBy,

            // --- By Letter ---
            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "References Column Letter (e.g., A)")
            String refColLetter,

            @Idx(index = "3.1.2", type = AttributeType.TEXT)
            @Pkg(label = "Values Column Letter (e.g., B)")
            String valColLetter,

            // --- By Header ---
            @Idx(index = "3.2.1", type = AttributeType.TEXT)
            @Pkg(label = "References Header Name")
            String refHeader,

            @Idx(index = "3.2.2", type = AttributeType.TEXT)
            @Pkg(label = "Values Header Name")
            String valHeader,

            // ===== GROUPING OPTIONS =====
            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "Case-sensitive grouping", default_value = "false", default_value_type = DataType.BOOLEAN)
            Boolean caseSensitive,

            // ===== DESTINATION WORKBOOK / SHEET =====
            @Idx(index = "6", type = AttributeType.SESSION)
            @Pkg(label = "Destination Workbook Session")
            @NotEmpty @SessionObject ExcelSession destExcelSession,

            @Idx(index = "7", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "7.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "7.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select destination sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String selectDestSheetBy,

            @Idx(index = "7.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Destination Sheet Name")
            String destSheetName,

            @Idx(index = "7.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Destination Sheet Index (1-based)")
            @NumberInteger @GreaterThanEqualTo("1") Double destSheetIndex,

            @Idx(index = "8", type = AttributeType.TEXT)
            @Pkg(label = "Destination top-left cell (e.g., H1)") @NotEmpty
            String destTopLeft
    ) {

        // ==== 1) Workbooks y Sheets (tu patrón) ====
        Session sourceSession = ExcelObjects.requireSession(sourceExcelSession);
        Dispatch wb1 = ExcelObjects.requireWorkbook(sourceSession, sourceExcelSession);

        Session destSession = ExcelObjects.requireSession(destExcelSession);
        Dispatch wb2 = ExcelObjects.requireWorkbook(destSession, destExcelSession);

        Dispatch srcSheet = ExcelObjects.requireSheet(wb1, selectSourceSheetBy, sourceSheetName, sourceSheetIndex);
        Dispatch dstSheet = ExcelObjects.requireSheet(wb2, selectDestSheetBy, destSheetName, destSheetIndex);

        // ==== 2) Resolver columnas ====
        final int headerRow = 1; // se asume headers en fila 1
        int refCol, valCol;

        if ("letter".equalsIgnoreCase(selectColsBy)) {
            if (refColLetter == null || refColLetter.trim().isEmpty())
                throw new BotCommandException("Reference column letter is required.");
            if (valColLetter == null || valColLetter.trim().isEmpty())
                throw new BotCommandException("Value column letter is required.");
            refCol = ExcelHelpers.excelColumnLetterToNumber(refColLetter.trim());
            valCol = ExcelHelpers.excelColumnLetterToNumber(valColLetter.trim());
        } else if ("header".equalsIgnoreCase(selectColsBy)) {
            if (refHeader == null || refHeader.trim().isEmpty())
                throw new BotCommandException("Reference header name is required.");
            if (valHeader == null || valHeader.trim().isEmpty())
                throw new BotCommandException("Value header name is required.");
            int totalCols = ExcelHelpers.getLastColumn(srcSheet);
            refCol = ExcelHelpers.headerNameToColumnIndex(srcSheet, refHeader.trim(), headerRow, totalCols);
            valCol = ExcelHelpers.headerNameToColumnIndex(srcSheet, valHeader.trim(), headerRow, totalCols);
        } else {
            throw new BotCommandException("Invalid 'Select columns by' option: " + selectColsBy);
        }

        // ==== 3) Rango de datos ====
        int lastDataRow = ExcelHelpers.getLastDataRow(srcSheet);
         if (lastDataRow <= headerRow) {
             // Nada que agrupar → escribir solo headers
             writeOutput(dstSheet, destTopLeft, Collections.emptyList());
             return;
         }
         if (lastDataRow <= headerRow) {
             // Nada que agrupar → no escribimos nada
             return;
         }

        int startRow = headerRow + 1;
        String refTop = ExcelHelpers.numberToColumnLetter(refCol) + startRow;
        String refBot = ExcelHelpers.numberToColumnLetter(refCol) + lastDataRow;
        String valTop = ExcelHelpers.numberToColumnLetter(valCol) + startRow;
        String valBot = ExcelHelpers.numberToColumnLetter(valCol) + lastDataRow;

        Dispatch refRange = range(srcSheet, refTop, refBot);
        Dispatch valRange = range(srcSheet, valTop, valBot);

        Variant refV = Dispatch.get(refRange, "Value2");
        Variant valV = Dispatch.get(valRange, "Value2");

        // ==== 4) Agregación en memoria (O(n)) ====
        Map<String, BigDecimal> acc = new LinkedHashMap<>(); // preserva orden de aparición
        boolean useCase = (caseSensitive != null) && caseSensitive;

        SafeArray refSA = isVariantArray(refV) ? refV.toSafeArray() : null;
        SafeArray valSA = isVariantArray(valV) ? valV.toSafeArray() : null;

        int rL = (refSA != null) ? refSA.getLBound(1) : 1;
        int rU = (refSA != null) ? refSA.getUBound(1) : 1;
        int cRef = (refSA != null) ? refSA.getLBound(2) : 1;
        int cVal = (valSA != null) ? valSA.getLBound(2) : 1;

        for (int r = rL; r <= rU; r++) {
            Variant refCell = (refSA != null) ? refSA.getVariant(new int[]{r, cRef}) : refV;
            Variant valCell = (valSA != null) ? valSA.getVariant(new int[]{r, cVal}) : valV;

            String key = trimToEmpty(refCell);
            if (!useCase) key = key.toUpperCase(Locale.ROOT); // case-insensitive por defecto

            if (key.isEmpty()) {
                // Omitimos claves vacías; si querés incluirlas lo cambiamos fácil.
                continue;
            }

            BigDecimal amount = parseStrictNumeric(valCell, key);
            acc.merge(key, amount, BigDecimal::add);
        }

        // ==== 5) Escritura en destino (en bloque, con headers) ====
        List<Map.Entry<String, BigDecimal>> rows = new ArrayList<>(acc.entrySet());
        writeOutput(dstSheet, destTopLeft, rows);
    }

    // ===== Helpers =====

    private static boolean isVariantArray(Variant v) {
        return v != null && ((v.getvt() & Variant.VariantArray) != 0);
    }

    private static Dispatch range(Dispatch sheet, String top, String bottom) {
        Dispatch tl = Dispatch.call(sheet, "Range", top).toDispatch();
        Dispatch br = Dispatch.call(sheet, "Range", bottom).toDispatch();
        return Dispatch.call(sheet, "Range", tl, br).toDispatch();
    }

    private static String trimToEmpty(Variant v) {
        if (v == null || v.isNull()) return "";
        Object o = v.toJavaObject();
        return (o == null) ? "" : o.toString().trim();
    }

    /** Falla si el valor no es estrictamente numérico. */
    private static BigDecimal parseStrictNumeric(Variant v, String keyForMsg) {
        if (v == null || v.isNull()) return BigDecimal.ZERO;
        Object o = v.toJavaObject();
        if (o instanceof Number) {
            return BigDecimal.valueOf(((Number) o).doubleValue());
        }
        String s = o.toString().trim();
        if (s.isEmpty()) return BigDecimal.ZERO;
        try {
            // Sin heurísticas regionales: si no es parseable directo, se considera no numérico
            return new BigDecimal(s);
        } catch (Exception ex) {
            throw new BotCommandException("Non-numeric value '" + s + "' for reference '" + keyForMsg + "'.");
        }
    }

    private static void writeOutput(Dispatch dstSheet, String destTopLeft,
                                    List<Map.Entry<String, BigDecimal>> rows) {
        if (destTopLeft == null || destTopLeft.trim().isEmpty())
            throw new BotCommandException("Destination top-left cell cannot be empty.");

        // Si no hay datos, no escribir (y no crear headers)
        if (rows == null || rows.isEmpty()) return;

        // Solo datos: 2 columnas (Reference, Total)
        int outRows = rows.size();
        int outCols = 2;

        Dispatch topLeft  = Dispatch.call(dstSheet, "Range", destTopLeft).toDispatch();
        Dispatch outRange = Dispatch.call(topLeft, "Resize",
                new Variant(outRows), new Variant(outCols)).toDispatch();

        // --- SafeArray 1-based: [1..outRows, 1..2] ---
        // IMPORTANTE: este constructor define LBound=1 para cada dimensión,
        // evitando el desplazamiento a B3 que veías con el constructor 0-based.
        com.jacob.com.SafeArray sa = new com.jacob.com.SafeArray(
                Variant.VariantVariant,
                new int[]{1, 1},
                new int[]{outRows, outCols}
        );

        // Llenar datos: col1 = Reference (key), col2 = Total (valor)
        for (int i = 1; i <= outRows; i++) {
            Map.Entry<String, BigDecimal> e = rows.get(i - 1);
            sa.setVariant(new int[]{i, 1}, new Variant(e.getKey()));
            sa.setVariant(new int[]{i, 2}, new Variant(e.getValue().doubleValue())); // numérico en Excel
        }

        Variant v = new Variant();
        v.putSafeArrayRef(sa);
        Dispatch.put(outRange, "Value2", v);
    }
}