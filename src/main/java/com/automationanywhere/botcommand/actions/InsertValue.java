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
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;

import static com.automationanywhere.botcommand.utilities.ExcelHelpers.*;

@BotCommand
@CommandPkg(
        label = "Set Value",
        name = "insertValue",
        description = "Inserta un valor o fórmula en una celda, columna o rango (optimizado)",
        icon = "excel.svg"
)
public class InsertValue {

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

                int colIndex;
                if ("letter".equalsIgnoreCase(selectColumnBy)) {
                    if (columnLetter == null || columnLetter.isEmpty())
                        throw new BotCommandException("Column letter not provided.");
                    colIndex = excelColumnLetterToNumber(columnLetter);
                } else {
                    String target = columnName.trim();
                    colIndex = ExcelHelpers.headerNameToColumnIndex(sheet, target, firstRow, colsCnt);
                }

            if ("inColumn".equalsIgnoreCase(selectColModeBy)) {
                // 1) Dos esquinas del rango
                Dispatch start = Dispatch.call(sheet, "Cells", startRowInput.intValue(), colIndex).toDispatch();
                Dispatch end   = Dispatch.call(sheet, "Cells", rowsCnt,    colIndex).toDispatch();

                // 2) Construir el rango con Range(Cell1, Cell2) en lugar de Resize
                // (algunas versiones requieren 'invoke' para la sobrecarga de 2 args)
                Dispatch rng = Dispatch.invoke(
                        sheet, "Range", Dispatch.Get,
                        new Object[]{ start, end }, new int[1]
                ).toDispatch();

                // 3) Escribir (idéntico a 'range')
                if (Boolean.TRUE.equals(isFormula)) {
                    Dispatch.put(rng, "Formula", value);   // también podés usar AutoFill si querés
                } else {
                    Dispatch.put(rng, "Value2",  value);   // escalar se replica a todo el rango
                }

            } else if ("endColumn".equalsIgnoreCase(selectColModeBy)) {
                    // Buscar última fila con datos en ESA columna (correcto en índices absolutos)
                    // 1) el helper devuelve un int (índice de última fila con datos en esa columna)
                    int lastDataRow = ExcelHelpers.getLastDataRowInColumn(sheet, colIndex);

                    // nada que hacer si la columna está vacía o si el startRow excede
                    if (lastDataRow == 0 || firstRow > lastDataRow) {
                        return;
                    }

                    int offset = (marginTopRows != null) ? marginTopRows.intValue() : 0;
                    int targetRow = lastDataRow + offset + 1;

                    Dispatch cell = Dispatch.call(sheet, "Cells", targetRow, colIndex).toDispatch();
                    if (Boolean.TRUE.equals(isFormula))
                        Dispatch.put(cell, "Formula", value);
                    else
                        Dispatch.put(cell, "Value2", value);
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
}