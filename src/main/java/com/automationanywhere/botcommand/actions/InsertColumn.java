package com.automationanywhere.botcommand.actions;

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

import static com.automationanywhere.botcommand.utilities.ExcelHelpers.excelColumnLetterToNumber;

@BotCommand
@CommandPkg(
        label = "Insert Column",
        name = "insertColumn",
        description = "Inserta una columna por letra y, opcionalmente, asigna un header en una fila específica",
        icon = "excel.svg"
)
public class InsertColumn {

    // Constantes Excel
    private static final int xlCalculationAutomatic = -4105;
    private static final int xlCalculationManual    = -4135;

    @Execute
    public void action(
            // --- Workbook / Sheet ---
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty @SessionObject ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name",  value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name") String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
            @NumberInteger @GreaterThanEqualTo("1") Double sheetIndex,

            // --- Insert ---
            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Column Letter (A, B, ...)", description = "Inserta ANTES de esta letra (estilo Columns(\"C:C\").Insert)")
            @NotEmpty String columnLetter,

            // --- Header (opcional) ---
            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "Set header", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean setHeader,

            @Idx(index = "4.1", type = AttributeType.TEXT)
            @Pkg(label = "Header Text")
            String headerText,

            @Idx(index = "4.2", type = AttributeType.NUMBER)
            @Pkg(label = "Header Row (1-based)", default_value = "1", default_value_type = DataType.NUMBER)
            @NumberInteger @GreaterThanEqualTo("1") Double headerRowInput
    ) {
        // Retry defensivo si el hilo no tenía COM inicializado (mismo patrón que tus otras acciones)
        try {
            run(excelSession, selectSheetBy, sheetName, sheetIndex, columnLetter, setHeader, headerText, headerRowInput);
        } catch (Exception first) {
            try {
                ComThread.InitSTA();
                run(excelSession, selectSheetBy, sheetName, sheetIndex, columnLetter, setHeader, headerText, headerRowInput);
            } catch (Exception second) {
                throw (second instanceof BotCommandException)
                        ? (BotCommandException) second
                        : new BotCommandException("InsertColumn failed: " + second.getMessage(), second);
            } finally {
                try { ComThread.Release(); } catch (Exception ignore) {}
            }
        }
    }

    private void run(
            ExcelSession excelSession,
            String selectSheetBy, String sheetName, Double sheetIndex,
            String columnLetter,
            Boolean setHeader, String headerText, Double headerRowInput
    ) {
        if (!"name".equalsIgnoreCase(selectSheetBy) && !"index".equalsIgnoreCase(selectSheetBy)) {
            throw new BotCommandException("Invalid 'Select sheet by'. Use 'name' or 'index'.");
        }
        if (columnLetter == null || columnLetter.trim().isEmpty()) {
            throw new BotCommandException("Column letter is required.");
        }
        final String colLetter = columnLetter.trim().toUpperCase();

        // 1) Excel objects
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb     = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheet  = ExcelObjects.requireSheet(wb, selectSheetBy, sheetName, sheetIndex);

        // 2) Proteger: si la hoja está protegida, no se puede insertar
        boolean protectedContents = getBool(sheet, "ProtectContents");
        if (protectedContents) {
            throw new BotCommandException("La hoja está protegida (ProtectContents=true). No se puede insertar columna.");
        }

        // 3) App state
        Dispatch app = Dispatch.get(wb, "Application").toDispatch();
        boolean prevUpd = getBool(app, "ScreenUpdating");
        boolean prevEvt = getBool(app, "EnableEvents");
        boolean prevAlr = getBool(app, "DisplayAlerts");
        int     prevCal = getInt (app, "Calculation");

        putBool(app, "ScreenUpdating", false);
        putBool(app, "EnableEvents",   false);
        putBool(app, "DisplayAlerts",  false);
        putInt (app, "Calculation",    xlCalculationManual);

        try {
            // 4) Insertar columna ANTES de la letra indicada, forma robusta:
            //    Cells(1, colIndex).EntireColumn.Insert
            int colIndex = excelColumnLetterToNumber(colLetter);

            // Validación de rango (1..16384 -> A..XFD)
            if (colIndex < 1 || colIndex > 16384) {
                throw new BotCommandException("Column letter fuera de rango: " + colLetter + " (Excel permite A..XFD).");
            }

            // Obtenemos una celda de esa columna y su columna completa
            Dispatch anyCellInTargetCol = Dispatch.call(sheet, "Cells", 1, colIndex).toDispatch();
            Dispatch entireColumn       = Dispatch.get(anyCellInTargetCol, "EntireColumn").toDispatch();

            // Insert sin parámetros: inserta la columna y desplaza a la derecha
            Dispatch.call(entireColumn, "Insert");

            // 5) Header opcional
            if (Boolean.TRUE.equals(setHeader)) {
                if (headerText == null) headerText = "";
                int headerRow = (headerRowInput == null) ? 1 : headerRowInput.intValue();
                if (headerRow < 1) headerRow = 1;

                // Tras Insert, la nueva columna ocupa el índice ORIGINAL de la letra
                int insertedColIndex = excelColumnLetterToNumber(colLetter);
                Dispatch cell = Dispatch.call(sheet, "Cells", headerRow, insertedColIndex).toDispatch();
                Dispatch.put(cell, "Value2", headerText);
            }

        } catch (Exception e) {
            throw new BotCommandException("InsertColumn failed: " + e.getMessage(), e);
        } finally {
            // 6) Restaurar estado
            putInt (app, "Calculation",   prevCal);
            putBool(app, "DisplayAlerts", prevAlr);
            putBool(app, "EnableEvents",  prevEvt);
            putBool(app, "ScreenUpdating",prevUpd);
        }
    }

    // --- Helpers de acceso seguro a propiedades COM ---
    private static boolean getBool(Dispatch obj, String prop) {
        try { return Dispatch.get(obj, prop).getBoolean(); } catch (Exception e) { return true; }
    }
    private static int getInt(Dispatch obj, String prop) {
        try { return Dispatch.get(obj, prop).getInt(); } catch (Exception e) { return xlCalculationAutomatic; }
    }
    private static void putBool(Dispatch obj, String prop, boolean v) {
        try { Dispatch.put(obj, prop, v); } catch (Exception ignore) {}
    }
    private static void putInt(Dispatch obj, String prop, int v) {
        try { Dispatch.put(obj, prop, new Variant(v)); } catch (Exception ignore) {}
    }
}