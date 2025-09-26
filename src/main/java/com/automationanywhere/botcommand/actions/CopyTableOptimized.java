package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * Copia una “tabla” definida por un Headers Range (ej.: C9:BM9).
 * Opciones:
 *  - Include headers? (yes/no)
 *  - Rows to copy: All / Only visible
 *  - Copy Mode: Overwrite / Append / Manual (+ celda)
 *  - Copy only values (no formats): checkbox
 *  - Create sheet if not exists?: yes/no
 *  - Source/Destination: sesión de libro y selección de hoja por nombre o índice.
 *
 * NOTAS:
 *  - Se asume UNA instancia de Excel (se re-adjunta en el hilo actual).
 *  - Si “Only visible”, se usan SpecialCells(xlCellTypeVisible) = 12.
 *  - Para “values only”:
 *      * Si copia TODO el bloque: usa Value2 (rápido, sin clipboard).
 *      * Si copia SOLO visibles: copia visibles al clipboard y hace PasteSpecial valores (-4163).
 */
@BotCommand
@CommandPkg(
        label = "Copy Table",
        name = "copyTable",
        description = "Copies a table identified by a headers range with options for headers, visible rows, mode and values-only.",
        icon = "excel.svg",
        return_type = DataType.BOOLEAN,
        return_required = true
)
public class CopyTableOptimized {

    // Constantes Excel
    private static final int xlCellTypeVisible = 12;
    private static final int xlCalculationAutomatic = -4105;
    private static final int xlCalculationManual = -4135;
    private static final int xlPasteValues = -4163;

    @Execute
    public Value<Boolean> action(
            // --- Origen ---
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Source Workbook Session")
            @SessionObject @NotEmpty ExcelSession sourceExcelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select origin sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String selectOriginSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Sheet Name")
            String originSheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Origin Sheet Index (1-based)")
            Double originSheetIndex,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Headers Range (e.g., C9:BM9)")
            @NotEmpty String headersRange,

            @Idx(index = "4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "No", value = "no"))
            })
            @Pkg(label = "Include headers?", default_value = "no", default_value_type = DataType.STRING)
            @SelectModes String includeHeaders,

            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "All rows", value = "all")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "Only visible rows", value = "visible"))
            })
            @Pkg(label = "Rows to copy", default_value = "all", default_value_type = DataType.STRING)
            @SelectModes String rowsMode,

            // --- Destino ---
            @Idx(index = "6", type = AttributeType.SESSION)
            @Pkg(label = "Destination Workbook Session")
            @SessionObject @NotEmpty ExcelSession destExcelSession,

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
            Double destSheetIndex,

            @Idx(index = "8", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "8.1", pkg = @Pkg(label = "Overwrite", value = "overwrite")),
                    @Idx.Option(index = "8.2", pkg = @Pkg(label = "Append", value = "append")),
                    @Idx.Option(index = "8.3", pkg = @Pkg(label = "Manual", value = "manual"))
            })
            @Pkg(label = "Copy mode", default_value = "overwrite", default_value_type = DataType.STRING)
            @SelectModes String copyMode,

            @Idx(index = "8.3.1", type = AttributeType.TEXT)
            @Pkg(label = "Manual destination cell (e.g., A2)")
            String manualCell,

            @Idx(index = "9", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "9.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "9.2", pkg = @Pkg(label = "No", value = "no"))
            })
            @Pkg(label = "Create sheet if not exists?", default_value = "yes", default_value_type = DataType.STRING)
            @SelectModes String createSheet,

            @Idx(index = "10", type = AttributeType.CHECKBOX)
            @Pkg(label = "Copy only values (no formats)", description = "If checked, pastes values only",
                    default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean valuesOnly
    ) {
        // Sin ComScope por preferencia: agregamos retry defensivo si el hilo actual no estaba inicializado para COM.
        try {
            return run(sourceExcelSession, selectOriginSheetBy, originSheetName, originSheetIndex,
                    headersRange, includeHeaders, rowsMode,
                    destExcelSession, selectDestSheetBy, destSheetName, destSheetIndex,
                    copyMode, manualCell, createSheet, valuesOnly);
        } catch (Exception first) {
            try {
                ComThread.InitSTA();
                return run(sourceExcelSession, selectOriginSheetBy, originSheetName, originSheetIndex,
                        headersRange, includeHeaders, rowsMode,
                        destExcelSession, selectDestSheetBy, destSheetName, destSheetIndex,
                        copyMode, manualCell, createSheet, valuesOnly);
            } catch (Exception second) {
                throw (second instanceof BotCommandException)
                        ? (BotCommandException) second
                        : new BotCommandException("CopyTable failed: " + second.getMessage(), second);
            } finally {
                try { ComThread.Release(); } catch (Exception ignore) {}
            }
        }
    }

    private Value<Boolean> run(
            ExcelSession sourceExcelSession, String selSrcBy, String srcName, Double srcIdx,
            String headersRange, String includeHeaders, String rowsMode,
            ExcelSession destExcelSession, String selDstBy, String dstName, Double dstIdx,
            String copyMode, String manualCell, String createSheet, Boolean valuesOnly
    ) {
        // Validaciones mínimas
        if (!"name".equalsIgnoreCase(selSrcBy) && !"index".equalsIgnoreCase(selSrcBy))
            throw new BotCommandException("Invalid 'Select origin sheet by'. Use 'name' or 'index'.");
        if (!"name".equalsIgnoreCase(selDstBy) && !"index".equalsIgnoreCase(selDstBy))
            throw new BotCommandException("Invalid 'Select destination sheet by'. Use 'name' or 'index'.");
        String cm = copyMode == null ? "overwrite" : copyMode.trim().toLowerCase();
        if (!cm.equals("overwrite") && !cm.equals("append") && !cm.equals("manual"))
            throw new BotCommandException("Invalid Copy mode. Use Overwrite, Append or Manual.");
        if (cm.equals("manual") && (manualCell == null || manualCell.trim().isEmpty()))
            throw new BotCommandException("Manual destination cell is required when Copy mode = Manual.");

        // 1) Excel y workbooks (reattach en este hilo)
        Session sourceSession = ExcelObjects.requireSession(sourceExcelSession);
        Dispatch wbSrc = ExcelObjects.requireWorkbook(sourceSession, sourceExcelSession);

        Session destSession = ExcelObjects.requireSession(destExcelSession);
        Dispatch wbDst = ExcelObjects.requireWorkbook(destSession, destExcelSession);

        // 2) Hojas
        Dispatch shSrc = ExcelObjects.requireSheet(wbSrc, selSrcBy, srcName, srcIdx);
        Dispatch shDst = ensureDestSheet(wbDst, selDstBy, dstName, dstIdx, createSheet);

        // 3) Calcular rango “tabla” a copiar a partir de Headers Range
        final String[] parts = headersRange.split(":");
        if (parts.length != 2) throw new BotCommandException("Invalid headers range: " + headersRange);

        final String startCell = parts[0].trim();  // ej. C9
        final String endCell   = parts[1].trim();  // ej. BM9
        final String startCol  = startCell.replaceAll("\\d", "");
        final String endCol    = endCell.replaceAll("\\d", "");
        final int headerRow = parseRow(endCell);

        Dispatch used = Dispatch.get(shSrc, "UsedRange").toDispatch();
        if (used == null || used.m_pDispatch == 0) return new BooleanValue(false);

        int usedFirstRow = Dispatch.get(used, "Row").getInt();
        int usedRows     = Dispatch.get(Dispatch.get(used, "Rows").toDispatch(), "Count").getInt();
        int lastRow      = usedFirstRow + usedRows - 1;

        int topRow = "yes".equalsIgnoreCase(includeHeaders) ? headerRow : (headerRow + 1);
        if (lastRow < topRow) return new BooleanValue(false);

        String srcRangeA1 = startCol + topRow + ":" + endCol + lastRow;
        Dispatch srcRange = Dispatch.call(shSrc, "Range", srcRangeA1).toDispatch();

        // 4) “Only visible rows” (si aplica)
        boolean visibleOnly = "visible".equalsIgnoreCase(rowsMode);
        Dispatch effectiveSrc = srcRange;
        if (visibleOnly) {
            effectiveSrc = specialCellsVisibleOrNull(srcRange);
            if (effectiveSrc == null) return new BooleanValue(false);
        }

        // 5) Resolver destino según Copy Mode
        Dispatch destStart;
        if (cm.equals("overwrite")) {
            Dispatch usedDst = Dispatch.get(shDst, "UsedRange").toDispatch();
            try { if (usedDst != null && usedDst.m_pDispatch != 0) Dispatch.call(usedDst, "Clear"); } catch (Exception ignore) {}
            destStart = Dispatch.call(shDst, "Range", "A1").toDispatch();
        } else if (cm.equals("append")) {
            Dispatch usedDst = Dispatch.get(shDst, "UsedRange").toDispatch();
            boolean destEmpty = isUsedRangeEmpty(usedDst);
            if (destEmpty) {
                destStart = Dispatch.call(shDst, "Range", "A1").toDispatch();
            } else {
                int dstFirstRow = Dispatch.get(usedDst, "Row").getInt();
                int dstRows     = Dispatch.get(Dispatch.get(usedDst, "Rows").toDispatch(), "Count").getInt();
                int dstLastRow  = dstFirstRow + dstRows - 1;
                destStart = Dispatch.call(shDst, "Cells", dstLastRow + 1, 1).toDispatch();
            }
        } else { // manual
            destStart = Dispatch.call(shDst, "Range", manualCell.trim()).toDispatch();
        }

        // 6) Optimización temporal de Excel (una sola Application)
        Dispatch app = Dispatch.get(wbSrc, "Application").toDispatch();
        boolean prevUpd = true, prevEvents = true, prevAlerts = true; int prevCalc = xlCalculationAutomatic;
        try {
            prevUpd   = getBool(app, "ScreenUpdating");
            prevEvents= getBool(app, "EnableEvents");
            prevAlerts= getBool(app, "DisplayAlerts");
            prevCalc  = getInt (app, "Calculation");

            putBool(app, "ScreenUpdating", false);
            putBool(app, "EnableEvents",   false);
            putBool(app, "DisplayAlerts",  false);
            putInt (app, "Calculation",    xlCalculationManual);

            // 7) Copiar
            if (Boolean.TRUE.equals(valuesOnly)) {
                if (visibleOnly) {
                    // Visible + solo valores → clipboard + PasteSpecial valores
                    Dispatch.call(effectiveSrc, "Copy");
                    pasteValuesWithRetry(shDst, destStart);
                    clearCutCopyMode(app);
                } else {
                    // Todo el bloque + valores → Value2 en bloque (rápido, sin clipboard)
                    Variant v = Dispatch.get(srcRange, "Value2"); // matriz 2D o escalar
                    // Dimensionar destino al mismo tamaño que srcRange
                    int srcRows = Dispatch.get(Dispatch.get(srcRange, "Rows").toDispatch(), "Count").getInt();
                    int srcCols = Dispatch.get(Dispatch.get(srcRange, "Columns").toDispatch(), "Count").getInt();
                    Dispatch destRange = Dispatch.call(destStart, "Resize", srcRows, srcCols).toDispatch();
                    Dispatch.put(destRange, "Value2", v);
                }
            } else {
                // Copiar con formato (estándar). Para visibles: el Range ya es “visible” y copia solo visibles.
                Dispatch.call(effectiveSrc, "Copy", destStart);
                clearCutCopyMode(app);
            }

            return new BooleanValue(true);

        } finally {
            // Restaurar estados
            putInt (app, "Calculation",    prevCalc);
            putBool(app, "DisplayAlerts",  prevAlerts);
            putBool(app, "EnableEvents",   prevEvents);
            putBool(app, "ScreenUpdating", prevUpd);
        }
    }

    // ------------- Helpers internos (sin dependencias externas) --------------

    private static int parseRow(String a1) {
        String digits = a1.replaceAll("\\D", "");
        if (digits.isEmpty()) throw new BotCommandException("Invalid A1 with no row number: " + a1);
        return Integer.parseInt(digits);
    }

    private static Dispatch ensureDestSheet(Dispatch wb, String selectBy, String name, Double index, String createSheet) {
        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        int count = Dispatch.get(sheets, "Count").getInt();
        for (int i = 1; i <= count; i++) {
            Dispatch s = Dispatch.call(sheets, "Item", i).toDispatch();
            String nm = Dispatch.get(s, "Name").getString();
            if ("index".equalsIgnoreCase(selectBy) && index != null && i == index.intValue()) return s;
            if ("name".equalsIgnoreCase(selectBy) && name != null && nm.equalsIgnoreCase(name)) return s;
        }
        if ("yes".equalsIgnoreCase(createSheet)) {
            Dispatch s = Dispatch.call(sheets, "Add").toDispatch();
            if (name != null && !name.trim().isEmpty()) {
                try { Dispatch.put(s, "Name", name.trim()); } catch (Exception ignore) {}
            }
            return s;
        }
        throw new BotCommandException("Destination sheet does not exist.");
    }

    private static Dispatch specialCellsVisibleOrNull(Dispatch range) {
        try { return Dispatch.call(range, "SpecialCells", new Variant(xlCellTypeVisible)).toDispatch(); }
        catch (Exception e) { return null; }
    }

    private static boolean isUsedRangeEmpty(Dispatch usedRange) {
        try {
            if (usedRange == null || usedRange.m_pDispatch == 0) return true;
            Variant v = Dispatch.get(usedRange, "Value");
            return v == null || v.isNull();
        } catch (Exception e) {
            return true;
        }
    }

    private static void pasteValuesWithRetry(Dispatch destSheet, Dispatch destStart) {
        // Después de Copy(), pegamos valores especiales con algunos reintentos
        final int maxAttempts = 5;
        final long[] waits = new long[]{70, 110, 170, 240, 320};
        for (int i = 0; i < maxAttempts; i++) {
            try {
                Dispatch.call(destStart, "PasteSpecial", new Variant(xlPasteValues));
                return;
            } catch (Exception e) {
                if (i < waits.length) sleepQuiet(waits[i]);
            }
        }
        // Si no se pudo, dejamos que la acción falle controladamente
        throw new BotCommandException("PasteSpecial(values) failed after retries.");
    }

    private static void clearCutCopyMode(Dispatch app) {
        try { Dispatch.put(app, "CutCopyMode", false); } catch (Exception ignore) {}
    }

    private static void sleepQuiet(long ms) {
        try { Thread.sleep(ms); } catch (InterruptedException ignore) {}
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