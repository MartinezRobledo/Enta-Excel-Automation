package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.ExcelSessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.GreaterThanEqualTo;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.NumberInteger;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Insert Formula",
        name = "insertFormula",
        description = "Inserta una fórmula en una celda o en toda una columna",
        icon = "excel.svg"
)
public class InsertFormula {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Session ID", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty
            String sessionId,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Workbook Path")
            @NotEmpty
            String workbookName,

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select origin sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Sheet Name")
            @NotEmpty
            String sheetName,

            @Idx(index = "3.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Origin Sheet Index (1-based)")
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double sheetIndex,

            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Formula")
            @NotEmpty
            String formula,

            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "Celda", value = "cell")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "Columna", value = "column"))
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

            @Idx(index = "5.2.1.1.2", type = AttributeType.CHECKBOX)
            @Pkg(label = "Trim header", default_value = "true", default_value_type = DataType.BOOLEAN)
            Boolean trimHeader
    ) {

        ExcelSession session = ExcelSessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found: " + sessionId);

        Dispatch wb = session.openWorkbooks.get(workbookName);
        if (wb == null)
            throw new BotCommandException("Workbook not open: " + workbookName);

        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        Dispatch sheet;
        if ("index".equalsIgnoreCase(selectSheetBy)) {
            sheet = Dispatch.call(sheets, "Item", sheetIndex.intValue()).toDispatch();
        } else {
            sheet = Dispatch.call(sheets, "Item", sheetName).toDispatch();
        }

        if ("cell".equalsIgnoreCase(insertMode)) {
            Dispatch cell = Dispatch.call(sheet, "Range", targetCell).toDispatch();
            Dispatch.put(cell, "Formula", formula);

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
                String headerTarget = trimHeader && columnName != null ? columnName.trim() : columnName;
                for (int c = 1; c <= totalCols; c++) {
                    Dispatch cell = Dispatch.call(sheet, "Cells", firstRow, c).toDispatch();
                    String val = safeVariantToString(Dispatch.get(cell, "Value"));
                    if (trimHeader && val != null) val = val.trim();
                    if (headerTarget.equalsIgnoreCase(val)) {
                        colIndex = c;
                        break;
                    }
                }
                if (colIndex == -1) throw new BotCommandException("Header not found: " + headerTarget);
            }

            // Inserto la fórmula en la primera celda vacía de la columna
            int startRow = firstRow + 1;
            for (int r = firstRow + 1; r <= totalRows; r++) {
                Dispatch cell = Dispatch.call(sheet, "Cells", r, colIndex).toDispatch();
                String val = safeVariantToString(Dispatch.get(cell, "Value"));
                if (val == null || val.isEmpty()) {
                    startRow = r;
                    Dispatch firstCell = Dispatch.call(sheet, "Cells", startRow, colIndex).toDispatch();
                    Dispatch.put(firstCell, "Formula", formula);

                    // Determinar rango final de autofill
                    Dispatch lastCell = Dispatch.call(sheet, "Cells", totalRows, colIndex).toDispatch();
                    Dispatch fillRange = Dispatch.call(sheet, "Range", firstCell, lastCell).toDispatch();

                    // Usar AutoFill para replicar la fórmula con referencias relativas
                    Dispatch.call(firstCell, "AutoFill", fillRange, 1 /* xlFillDefault */);
                    break;
                }
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
