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

import static com.automationanywhere.botcommand.utilities.ExcelHelpers.*;
// headerNameToColumnIndex, excelColumnLetterToNumber, getLastDataRowInColumn

@BotCommand
@CommandPkg(
        label = "Copy Column",
        name = "copyColumn",
        description = "Copia una columna por Header o Letter hacia una columna destino (Header o Letter).",
        icon = "excel.svg",
        return_type = DataType.BOOLEAN,
        return_required = true
)
public class CopyColumn {

    // Constantes Excel (alineadas con CopyTableOptimized)
    private static final int xlCellTypeVisible      = 12;
    private static final int xlCalculationAutomatic = -4105;
    private static final int xlCalculationManual    = -4135;
    private static final int xlPasteValues          = -4163;

    @Execute
    public Value<Boolean> action(
            // --- ORIGEN ---
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Source Workbook Session")
            @SessionObject @NotEmpty ExcelSession sourceExcelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name",  value = "name")),
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

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Header", value = "header")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Letter", value = "letter"))
            })
            @Pkg(label = "Select ORIGIN Column By", default_value = "letter", default_value_type = DataType.STRING)
            @SelectModes String selectOriginColumnBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Column Header Name")
            String originColumnHeader,

            @Idx(index = "3.2.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Column Letter (A, B, ...)")
            String originColumnLetter,

            @Idx(index = "4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "No",  value = "no"))
            })
            @Pkg(label = "Include header from ORIGIN?", default_value = "no", default_value_type = DataType.STRING)
            @SelectModes String includeHeader,

            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "All rows",          value = "all")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "Only visible rows", value = "visible"))
            })
            @Pkg(label = "Rows to copy", default_value = "all", default_value_type = DataType.STRING)
            @SelectModes String rowsMode,

            // --- DESTINO ---
            @Idx(index = "6", type = AttributeType.SESSION)
            @Pkg(label = "Destination Workbook Session")
            @SessionObject @NotEmpty ExcelSession destExcelSession,

            @Idx(index = "7", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "7.1", pkg = @Pkg(label = "Name",  value = "name")),
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

            // *** NUEVO BLOQUE: DESTINATION COLUMN ***
            @Idx(index = "8", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "8.1", pkg = @Pkg(label = "Header", value = "header")),
                    @Idx.Option(index = "8.2", pkg = @Pkg(label = "Letter", value = "letter"))
            })
            @Pkg(label = "Select DESTINATION Column By", default_value = "letter", default_value_type = DataType.STRING)
            @SelectModes String selectDestColumnBy,

            @Idx(index = "8.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Destination Column Header Name")
            String destColumnHeader,

            @Idx(index = "8.2.1", type = AttributeType.TEXT)
            @Pkg(label = "Destination Column Letter (A, B, ...)")
            String destColumnLetter,

            @Idx(index = "9", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "9.1", pkg = @Pkg(label = "Overwrite", value = "overwrite")),
                    @Idx.Option(index = "9.2", pkg = @Pkg(label = "Append",    value = "append")),
                    @Idx.Option(index = "9.3", pkg = @Pkg(label = "Manual",    value = "manual"))
            })
            @Pkg(label = "Copy mode", default_value = "overwrite", default_value_type = DataType.STRING)
            @SelectModes String copyMode,

            @Idx(index = "9.3.1", type = AttributeType.TEXT)
            @Pkg(label = "Manual START cell (row is used; column will be the DESTINATION column)",
                    description = "Ej.: A2 → se usa la FILA 2 pero la COLUMNA es la de destino seleccionada")
            String manualCell,

            @Idx(index = "10", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "10.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "10.2", pkg = @Pkg(label = "No",  value = "no"))
            })
            @Pkg(label = "Create sheet if not exists?", default_value = "yes", default_value_type = DataType.STRING)
            @SelectModes String createSheet,

            @Idx(index = "11", type = AttributeType.CHECKBOX)
            @Pkg(label = "Copy only values (no formats)",
                    default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean valuesOnly
    ) {

        String cm = copyMode == null ? "overwrite" : copyMode.trim().toLowerCase();
        if (!cm.equals("overwrite") && !cm.equals("append") && !cm.equals("manual"))
            throw new BotCommandException("Invalid Copy mode. Use Overwrite, Append or Manual.");
        if (cm.equals("manual") && (manualCell == null || manualCell.trim().isEmpty()))
            throw new BotCommandException("Manual START cell is required when Copy mode = Manual.");

        // 1) Workbooks
        Session sourceSession = ExcelObjects.requireSession(sourceExcelSession);
        Dispatch wbSrc        = ExcelObjects.requireWorkbook(sourceSession, sourceExcelSession);

        Session destSession   = ExcelObjects.requireSession(destExcelSession);
        Dispatch wbDst        = ExcelObjects.requireWorkbook(destSession, destExcelSession);

        // 2) Hojas
        Dispatch shSrc = ExcelObjects.requireSheet(wbSrc, selectOriginSheetBy, originSheetName, originSheetIndex);
        Dispatch shDst = ensureDestSheet(wbDst, selectDestSheetBy, destSheetName, destSheetIndex, createSheet);

        // 3) ORIGEN: headerRow y colIndex
        Dispatch usedSrc = Dispatch.get(shSrc, "UsedRange").toDispatch();
        if (usedSrc == null || usedSrc.m_pDispatch == 0) return new BooleanValue(false);
        int srcHeaderRow = Dispatch.get(usedSrc, "Row").getInt(); // 1ª fila de UsedRange
        int srcColsCnt   = Dispatch.get(Dispatch.get(usedSrc, "Columns").toDispatch(), "Count").getInt();

        int srcColIndex;
        if ("letter".equalsIgnoreCase(selectOriginColumnBy)) {
            if (originColumnLetter == null || originColumnLetter.trim().isEmpty())
                throw new BotCommandException("Origin Column letter not provided.");
            srcColIndex = excelColumnLetterToNumber(originColumnLetter.trim());
        } else {
            if (originColumnHeader == null || originColumnHeader.trim().isEmpty())
                throw new BotCommandException("Origin Column header not provided.");
            srcColIndex = headerNameToColumnIndex(shSrc, originColumnHeader.trim(), srcHeaderRow, srcColsCnt);
            if (srcColIndex <= 0) throw new BotCommandException("Origin header not found: " + originColumnHeader);
        }

        int srcLastDataRow = getLastDataRowInColumn(shSrc, srcColIndex);
        if (srcLastDataRow <= 0) return new BooleanValue(false);

        int srcStartRow = "yes".equalsIgnoreCase(includeHeader) ? srcHeaderRow : (srcHeaderRow + 1);
        if (srcLastDataRow < srcStartRow) return new BooleanValue(false);

        Dispatch srcStart = Dispatch.call(shSrc, "Cells", srcStartRow,    srcColIndex).toDispatch();
        Dispatch srcEnd   = Dispatch.call(shSrc, "Cells", srcLastDataRow, srcColIndex).toDispatch();
        Dispatch srcRange = Dispatch.invoke(shSrc, "Range", Dispatch.Get, new Object[]{ srcStart, srcEnd }, new int[1]).toDispatch();

        boolean visibleOnly = "visible".equalsIgnoreCase(rowsMode);
        Dispatch effectiveSrc = srcRange;
        if (visibleOnly) {
            effectiveSrc = specialCellsVisibleOrNull(srcRange);
            if (effectiveSrc == null) return new BooleanValue(false);
        }

        // 4) DESTINO: headerRow y destColIndex
        Dispatch usedDst = Dispatch.get(shDst, "UsedRange").toDispatch();
        int dstHeaderRow = 1; // si no hay UsedRange, asumimos fila 1
        int dstColsCnt   = 16384;
        if (usedDst != null && usedDst.m_pDispatch != 0) {
            dstHeaderRow = Dispatch.get(usedDst, "Row").getInt();
            dstColsCnt   = Dispatch.get(Dispatch.get(usedDst, "Columns").toDispatch(), "Count").getInt();
        }

        int dstColIndex;
        if ("letter".equalsIgnoreCase(selectDestColumnBy)) {
            if (destColumnLetter == null || destColumnLetter.trim().isEmpty())
                throw new BotCommandException("Destination Column letter not provided.");
            dstColIndex = excelColumnLetterToNumber(destColumnLetter.trim());
        } else {
            if (destColumnHeader == null || destColumnHeader.trim().isEmpty())
                throw new BotCommandException("Destination Column header not provided.");
            dstColIndex = headerNameToColumnIndex(shDst, destColumnHeader.trim(), dstHeaderRow, dstColsCnt);
            if (dstColIndex <= 0) throw new BotCommandException("Destination header not found: " + destColumnHeader);
        }

        // 5) Resolver destStart según Copy Mode (y reglas pedidas)
        Dispatch destStart;
        if (cm.equals("overwrite")) {
            // Limpiar SOLO la columna destino
            Dispatch dstColRange = Dispatch.call(shDst, "Columns", dstColIndex).toDispatch();
            try { Dispatch.call(dstColRange, "Clear"); } catch (Exception ignore) {}
            int startRow = "yes".equalsIgnoreCase(includeHeader) ? dstHeaderRow : (dstHeaderRow + 1);
            destStart = Dispatch.call(shDst, "Cells", startRow, dstColIndex).toDispatch();

        } else if (cm.equals("append")) {
            int dstLastDataRow = getLastDataRowInColumn(shDst, dstColIndex);
            int startRow;
            if (dstLastDataRow > 0) {
                startRow = dstLastDataRow + 1;
            } else {
                startRow = "yes".equalsIgnoreCase(includeHeader) ? dstHeaderRow : (dstHeaderRow + 1);
            }
            destStart = Dispatch.call(shDst, "Cells", startRow, dstColIndex).toDispatch();

        } else { // manual
            // Usar la FILA de manualCell, pero la COLUMNA SIEMPRE es la de destino
            int manualRow = parseRow(manualCell.trim());
            if (manualRow < 1) throw new BotCommandException("Invalid manual START cell row: " + manualCell);
            destStart = Dispatch.call(shDst, "Cells", manualRow, dstColIndex).toDispatch();
        }

        // 6) Optimización Excel
        Dispatch app = Dispatch.get(wbSrc, "Application").toDispatch();
        boolean prevUpd = true, prevEvents = true, prevAlerts = true;
        int prevCalc = xlCalculationAutomatic;

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
                    // Visibles + solo valores → clipboard + PasteSpecial valores
                    Dispatch.call(effectiveSrc, "Copy");
                    pasteValuesWithRetry(shDst, destStart);
                    clearCutCopyMode(app);
                } else {
                    // Bloque completo → Value2 directo (1 sola columna)
                    Variant v = Dispatch.get(srcRange, "Value2");
                    int srcRows = Dispatch.get(Dispatch.get(srcRange, "Rows").toDispatch(), "Count").getInt();
                    Dispatch destRange = Dispatch.call(destStart, "Resize", srcRows, 1).toDispatch();
                    Dispatch.put(destRange, "Value2", v);
                }
            } else {
                // Copiar con formato (si visibleOnly, el range ya es 'visible')
                Dispatch.call(effectiveSrc, "Copy", destStart);
                clearCutCopyMode(app);
            }

            return new BooleanValue(true);
        } finally {
            putInt (app, "Calculation",   prevCalc);
            putBool(app, "DisplayAlerts", prevAlerts);
            putBool(app, "EnableEvents",  prevEvents);
            putBool(app, "ScreenUpdating",prevUpd);
        }
    }

    // ===== Helpers internos (como en CopyTableOptimized) =====

    private static int parseRow(String a1) {
        String digits = a1.replaceAll("\\D", "");
        if (digits.isEmpty()) throw new BotCommandException("Invalid A1 with no row number: " + a1);
        return Integer.parseInt(digits);
    }

    private static Dispatch ensureDestSheet(Dispatch wb, String selectBy, String name, Double index, String createSheet) {
        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        int count = Dispatch.get(sheets, "Count").getInt();
        for (int i = 1; i <= count; i++) {
            Dispatch s  = Dispatch.call(sheets, "Item", i).toDispatch();
            String nm   = Dispatch.get(s, "Name").getString();
            if ("index".equalsIgnoreCase(selectBy) && index != null && i == index.intValue()) return s;
            if ("name".equalsIgnoreCase(selectBy)  && name  != null && nm.equalsIgnoreCase(name)) return s;
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

    private static void pasteValuesWithRetry(Dispatch destSheet, Dispatch destStart) {
        final int maxAttempts = 5;
        final long[] waits = new long[]{70, 110, 170, 240, 320};
        for (int i = 0; i < maxAttempts; i++) {
            try {
                Dispatch.call(destStart, "PasteSpecial", new Variant(xlPasteValues));
                return;
            } catch (Exception e) {
                if (i < waits.length) {
                    try { Thread.sleep(waits[i]); } catch (InterruptedException ignore) {}
                }
            }
        }
        throw new BotCommandException("PasteSpecial(values) failed after retries.");
    }

    private static void clearCutCopyMode(Dispatch app) {
        try { Dispatch.put(app, "CutCopyMode", false); } catch (Exception ignore) {}
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
