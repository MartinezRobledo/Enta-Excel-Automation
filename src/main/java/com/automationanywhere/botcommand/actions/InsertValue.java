package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Insert Value",
        name = "insertValue",
        description = "Inserta un valor o fórmula en una celda, columna o rango (optimizado)",
        icon = "excel.svg"
)
public class InsertValue {

    private static final int xlCalculationAutomatic = -4105;
    private static final int xlCalculationManual    = -4135;

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty @SessionObject ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
            @NumberInteger @GreaterThanEqualTo("1") @NotEmpty Double sheetIndex,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Value or Formula")
            @NotEmpty String value,

            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "Is Formula?", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean isFormula,

            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "Celda",   value = "cell")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "Columna", value = "column")),
                    @Idx.Option(index = "5.3", pkg = @Pkg(label = "Rango",   value = "range"))
            })
            @Pkg(label = "Insert Mode", default_value = "cell", default_value_type = DataType.STRING)
            String insertMode,

            @Idx(index = "5.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Target Cell (ej A1)")
            @NotEmpty String targetCell,

            @Idx(index = "5.2.1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.2.1.1", pkg = @Pkg(label = "Header", value = "header")),
                    @Idx.Option(index = "5.2.1.2", pkg = @Pkg(label = "Letter", value = "letter"))
            })
            @Pkg(label = "Select Column By", default_value = "letter", default_value_type = DataType.STRING)
            @SelectModes String selectColumnBy,

            @Idx(index = "5.2.1.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Column Header Name")
            @NotEmpty String columnName,

            @Idx(index = "5.2.1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "Column Letter (A, B, ...)")
            @NotEmpty String columnLetter,

            @Idx(index = "5.2.2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.2.2.1", pkg = @Pkg(label = "In Column",      value = "inColumn")),
                    @Idx.Option(index = "5.2.2.2", pkg = @Pkg(label = "End of Column",   value = "endColumn"))
            })
            @Pkg(label = "Select Mode By", default_value = "inColumn", default_value_type = DataType.STRING)
            @SelectModes String selectColModeBy,

            @Idx(index = "5.2.2.1.1", type = AttributeType.NUMBER)
            @Pkg(label = "Start Row (for column insert)", default_value = "2", default_value_type = DataType.NUMBER)
            @NumberInteger @GreaterThanEqualTo("1") Double startRowInput,

            @Idx(index = "5.2.2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Margin top (in rows)", default_value = "0", default_value_type = DataType.NUMBER)
            @NumberInteger @GreaterThanEqualTo("0") @NotEmpty Double marginTopRows,

            @Idx(index = "5.3.1", type = AttributeType.TEXT)
            @Pkg(label = "Target range (ej A1:F12)")
            @NotEmpty String targetRange
    ) {
        try {
            run(excelSession, selectSheetBy, sheetName, sheetIndex, value, isFormula,
                    insertMode, targetCell, selectColumnBy, columnName, columnLetter,
                    selectColModeBy, startRowInput, marginTopRows, targetRange);
        } catch (Exception first) {
            // Retry defensivo por si el hilo no tenía COM inicializado
            try {
                ComThread.InitSTA();
                run(excelSession, selectSheetBy, sheetName, sheetIndex, value, isFormula,
                        insertMode, targetCell, selectColumnBy, columnName, columnLetter,
                        selectColModeBy, startRowInput, marginTopRows, targetRange);
            } catch (Exception second) {
                throw (second instanceof BotCommandException)
                        ? (BotCommandException) second
                        : new BotCommandException("InsertValue failed: " + second.getMessage(), second);
            } finally {
                try { ComThread.Release(); } catch (Exception ignore) {}
            }
        }
    }

    private void run(
            ExcelSession excelSession, String selectSheetBy, String sheetName, Double sheetIndex,
            String value, Boolean isFormula, String insertMode, String targetCell,
            String selectColumnBy, String columnName, String columnLetter,
            String selectColModeBy, Double startRowInput, Double marginTopRows, String targetRange
    ) {
        // Reattach a Excel en este hilo (robusto). Si todavía usás la versión antigua:
        // Session session = ExcelObjects.requireSession(excelSession);
        // Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);

        Dispatch sheet = ExcelObjects.requireSheet(wb, selectSheetBy, sheetName, sheetIndex);
        Dispatch app = Dispatch.get(wb, "Application").toDispatch();

        // Guardar estado y optimizar durante la escritura
        boolean prevUpd = getBool(app, "ScreenUpdating");
        boolean prevEvt = getBool(app, "EnableEvents");
        boolean prevAlr = getBool(app, "DisplayAlerts");
        int prevCalc     = getInt (app, "Calculation");

        putBool(app, "ScreenUpdating", false);
        putBool(app, "EnableEvents",   false);
        putBool(app, "DisplayAlerts",  false);
        putInt (app, "Calculation",    xlCalculationManual);

        try {
            if ("cell".equalsIgnoreCase(insertMode)) {
                Dispatch cell = Dispatch.call(sheet, "Range", targetCell).toDispatch();
                if (Boolean.TRUE.equals(isFormula)) Dispatch.put(cell, "Formula", value);
                else                                 Dispatch.put(cell, "Value2", value);

            } else if ("column".equalsIgnoreCase(insertMode)) {
                Dispatch used = Dispatch.get(sheet, "UsedRange").toDispatch();
                int firstRow = Dispatch.get(used, "Row").getInt();
                int rowsCnt  = Dispatch.get(Dispatch.get(used, "Rows").toDispatch(), "Count").getInt();
                int colsCnt  = Dispatch.get(Dispatch.get(used, "Columns").toDispatch(), "Count").getInt();
                int lastRow  = firstRow + rowsCnt - 1;  // **FIX** índice absoluto correcto

                int colIndex;
                if ("letter".equalsIgnoreCase(selectColumnBy)) {
                    if (columnLetter == null || columnLetter.isEmpty())
                        throw new BotCommandException("Column letter not provided.");
                    colIndex = excelColumnLetterToNumber(columnLetter);
                } else {
                    if (columnName == null || columnName.isEmpty())
                        throw new BotCommandException("Column header not provided.");
                    colIndex = -1;
                    String target = columnName.trim();
                    for (int c = 1; c <= colsCnt; c++) {
                        Dispatch hdrCell = Dispatch.call(sheet, "Cells", firstRow, c).toDispatch();
                        String hdr = safeVariantToString(Dispatch.get(hdrCell, "Value"));
                        if (hdr != null && hdr.trim().equalsIgnoreCase(target)) { colIndex = c; break; }
                    }
                    if (colIndex == -1) throw new BotCommandException("Header not found: " + target);
                }

                if ("inColumn".equalsIgnoreCase(selectColModeBy)) {
                    int startRow = (startRowInput != null) ? startRowInput.intValue() : firstRow + 1;
                    if (startRow > lastRow) return;

                    // *** VECTORIZAR: construir el rango completo y setear en UNA llamada ***
                    int height = (lastRow - startRow + 1);
                    Dispatch start = Dispatch.call(sheet, "Cells", startRow, colIndex).toDispatch();
                    Dispatch rng   = Dispatch.call(start, "Resize", height, 1).toDispatch();
                    if (Boolean.TRUE.equals(isFormula)) Dispatch.put(rng, "Formula", value);
                    else                                 Dispatch.put(rng, "Value2", value);

                } else if ("endColumn".equalsIgnoreCase(selectColModeBy)) {
                    // Buscar última fila con datos en ESA columna (correcto en índices absolutos)
                    int lastDataRow = firstRow - 1;
                    for (int r = firstRow; r <= lastRow; r++) {
                        Dispatch cell = Dispatch.call(sheet, "Cells", r, colIndex).toDispatch();
                        String v = safeVariantToString(Dispatch.get(cell, "Value"));
                        if (v != null && !v.isEmpty()) lastDataRow = r;
                    }
                    int offset = (marginTopRows != null) ? marginTopRows.intValue() : 0;
                    int targetRow = (lastDataRow >= firstRow ? lastDataRow : firstRow - 1) + offset + 1;
                    Dispatch cell = Dispatch.call(sheet, "Cells", targetRow, colIndex).toDispatch();
                    if (Boolean.TRUE.equals(isFormula)) Dispatch.put(cell, "Formula", value);
                    else                                 Dispatch.put(cell, "Value2", value);
                } else {
                    throw new BotCommandException("Invalid Select Mode By: " + selectColModeBy);
                }

            } else if ("range".equalsIgnoreCase(insertMode)) {
                Dispatch rng = Dispatch.call(sheet, "Range", targetRange).toDispatch();
                if (Boolean.TRUE.equals(isFormula)) Dispatch.put(rng, "Formula", value);
                else                                 Dispatch.put(rng, "Value2", value);

            } else {
                throw new BotCommandException("Invalid insert mode: " + insertMode);
            }

        } finally {
            // Restaurar estado Excel
            putInt (app, "Calculation",    prevCalc);
            putBool(app, "DisplayAlerts",  prevAlr);
            putBool(app, "EnableEvents",   prevEvt);
            putBool(app, "ScreenUpdating", prevUpd);
        }
    }

    private static int excelColumnLetterToNumber(String col) {
        int res = 0; col = col.toUpperCase();
        for (int i = 0; i < col.length(); i++) res = res * 26 + (col.charAt(i) - 'A' + 1);
        return res;
    }
    private static String safeVariantToString(Variant v) {
        if (v == null || v.isNull()) return "";
        Object o = v.toJavaObject();
        return (o != null) ? o.toString() : "";
    }
    private static boolean getBool(Dispatch app, String prop) {
        try { return Dispatch.get(app, prop).getBoolean(); } catch (Exception e) { return true; }
    }
    private static int getInt(Dispatch app, String prop) {
        try { return Dispatch.get(app, prop).getInt(); } catch (Exception e) { return xlCalculationAutomatic; }
    }
    private static void putBool(Dispatch app, String prop, boolean v) {
        try { Dispatch.put(app, prop, v); } catch (Exception ignore) {}
    }
    private static void putInt(Dispatch app, String prop, int v) {
        try { Dispatch.put(app, prop, new Variant(v)); } catch (Exception ignore) {}
    }
}