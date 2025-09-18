package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Insert Value",
        name = "insertValue",
        description = "Inserta un valor o fórmula en una celda, columna o rango",
        icon = "excel.svg"
)
public class InsertValue {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty
            String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double sheetIndex,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Value or Formula")
            @NotEmpty
            String value,

            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "Is Formula?", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean isFormula,

            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "Celda", value = "cell")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "Columna", value = "column")),
                    @Idx.Option(index = "5.3", pkg = @Pkg(label = "Rango", value = "range"))
            })
            @Pkg(label = "Insert Mode", default_value = "cell", default_value_type = DataType.STRING)
            String insertMode,

            @Idx(index = "5.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Target Cell (ej A1)")
            @NotEmpty
            String targetCell,

            @Idx(index = "5.2.1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.2.1.1", pkg = @Pkg(label = "Header", value = "header")),
                    @Idx.Option(index = "5.2.1.2", pkg = @Pkg(label = "Letter", value = "letter"))
            })
            @Pkg(label = "Select Column By", default_value = "letter", default_value_type = DataType.STRING)
            @SelectModes
            String selectColumnBy,

            @Idx(index = "5.2.1.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Column Header Name")
            @NotEmpty
            String columnName,

            @Idx(index = "5.2.1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "Column Letter (A, B, ...)")
            @NotEmpty
            String columnLetter,

            @Idx(index = "5.2.2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.2.2.1", pkg = @Pkg(label = "In Column", value = "inColumn")),
                    @Idx.Option(index = "5.2.2.2", pkg = @Pkg(label = "End of Column", value = "endColumn"))
            })
            @Pkg(label = "Select Mode By", default_value = "inColumn", default_value_type = DataType.STRING)
            @SelectModes
            String selectColModeBy,

            @Idx(index = "5.2.2.1.1", type = AttributeType.NUMBER)
            @Pkg(label = "Start Row (for column insert)", default_value = "2", default_value_type = DataType.NUMBER)
            @NumberInteger
            @GreaterThanEqualTo("1")
            Double startRowInput,

            @Idx(index = "5.2.2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Margin top (in rows)", default_value = "0", default_value_type = DataType.NUMBER)
            @NumberInteger
            @GreaterThanEqualTo("0")
            @NotEmpty
            Double marginTopRows,

            @Idx(index = "5.3.1", type = AttributeType.TEXT)
            @Pkg(label = "Target range (ej A1:F12)")
            @NotEmpty
            String targetRange
    ) {

        Session session = excelSession.getSession();
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found o closed");

        if (session.openWorkbooks.isEmpty())
            throw new BotCommandException("No workbook is open in this session.");

        Dispatch wb = session.openWorkbooks.values().iterator().next();

        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        Dispatch sheet;
        if ("index".equalsIgnoreCase(selectSheetBy)) {
            sheet = Dispatch.call(sheets, "Item", sheetIndex.intValue()).toDispatch();
        } else {
            sheet = Dispatch.call(sheets, "Item", sheetName).toDispatch();
        }

        if ("cell".equalsIgnoreCase(insertMode)) {
            Dispatch cell = Dispatch.call(sheet, "Range", targetCell).toDispatch();
            if (Boolean.TRUE.equals(isFormula)) {
                Dispatch.put(cell, "Formula", value);
            } else {
                Dispatch.put(cell, "Value", value);
            }

        } else if ("column".equalsIgnoreCase(insertMode)) {

            Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
            int firstRow = Dispatch.get(usedRange, "Row").getInt();
            int totalRows = Dispatch.get(Dispatch.get(usedRange, "Rows").toDispatch(), "Count").getInt();
            int totalCols = Dispatch.get(Dispatch.get(usedRange, "Columns").toDispatch(), "Count").getInt();

            int colIndex;
            if ("letter".equalsIgnoreCase(selectColumnBy)) {
                if (columnLetter == null || columnLetter.isEmpty())
                    throw new BotCommandException("Column letter not provided.");
                colIndex = excelColumnLetterToNumber(columnLetter);
            } else { // header
                if (columnName == null || columnName.isEmpty())
                    throw new BotCommandException("Column header not provided.");
                colIndex = -1;
                String headerTarget = columnName.trim();
                for (int c = 1; c <= totalCols; c++) {
                    Dispatch cell = Dispatch.call(sheet, "Cells", firstRow, c).toDispatch();
                    String val = safeVariantToString(Dispatch.get(cell, "Value"));
                    if (val != null && val.trim().equalsIgnoreCase(headerTarget)) {
                        colIndex = c;
                        break;
                    }
                }
                if (colIndex == -1) throw new BotCommandException("Header not found: " + headerTarget);
            }

            if ("inColumn".equalsIgnoreCase(selectColModeBy)) {
                // Insertar desde una fila fija hacia abajo
                int startRow = startRowInput != null ? startRowInput.intValue() : firstRow + 1;

                for (int r = startRow; r <= totalRows; r++) {
                    Dispatch cell = Dispatch.call(sheet, "Cells", r, colIndex).toDispatch();
                    if (Boolean.TRUE.equals(isFormula)) {
                        Dispatch.put(cell, "Formula", value);
                    } else {
                        Dispatch.put(cell, "Value", value);
                    }
                }

            } else if ("endColumn".equalsIgnoreCase(selectColModeBy)) {
                // Buscar la última fila con datos en esa columna
                int lastDataRow = firstRow;
                for (int r = firstRow; r <= totalRows; r++) {
                    Dispatch cell = Dispatch.call(sheet, "Cells", r, colIndex).toDispatch();
                    String val = safeVariantToString(Dispatch.get(cell, "Value"));
                    if (val != null && !val.isEmpty()) {
                        lastDataRow = r;
                    }
                }

                int targetRow = lastDataRow + (marginTopRows != 0 ? marginTopRows.intValue() + 1 : 1);

                Dispatch cell = Dispatch.call(sheet, "Cells", targetRow, colIndex).toDispatch();
                if (Boolean.TRUE.equals(isFormula)) {
                    Dispatch.put(cell, "Formula", value);
                } else {
                    Dispatch.put(cell, "Value", value);
                }
            }

        } else if ("range".equalsIgnoreCase(insertMode)) {
            Dispatch range = Dispatch.call(sheet, "Range", targetRange).toDispatch();
            if (Boolean.TRUE.equals(isFormula)) {
                Dispatch.put(range, "Formula", value);
            } else {
                Dispatch.put(range, "Value", value);
            }
        } else {
            throw new BotCommandException("Invalid insert mode: " + insertMode);
        }
    }

    private static int excelColumnLetterToNumber(String col) {
        int res = 0;
        col = col.toUpperCase();
        for (int i = 0; i < col.length(); i++) {
            res = res * 26 + (col.charAt(i) - 'A' + 1);
        }
        return res;
    }

    private static String safeVariantToString(Variant v) {
        if (v == null || v.isNull()) return "";
        Object o = v.toJavaObject();
        return o != null ? o.toString() : "";
    }
}
