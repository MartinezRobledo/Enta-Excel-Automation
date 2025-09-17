package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Range to CSV String",
        name = "rangeToCSV",
        description = "Recibe un rango de Excel y retorna un string con todos los valores separados por coma",
        icon = "excel.svg",
        return_type = DataType.STRING,
        return_required = true
)
public class RangeToCSV {

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Session ID", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty String sessionId,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Workbook Name")
            @NotEmpty String workbookName,

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty
            String sheetName,

            @Idx(index = "3.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)")
            @NotEmpty
            Double sheetIndex,

            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Range (e.g., A1:C5)")
            @NotEmpty String rangeStr,

            @Idx(index = "5", type = AttributeType.CHECKBOX)
            @Pkg(label = "Ignorar columnas vacias", description = "Si está marcado, las celdas vacías no se incluirán en el CSV",
                    default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean ignoreEmpty
    ) {

        Session session = SessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Excel session not found: " + sessionId);

        Dispatch workbook = session.openWorkbooks.get(workbookName);
        if (workbook == null)
            throw new BotCommandException("Workbook not open: " + workbookName);

        Dispatch sheets = Dispatch.get(workbook, "Sheets").toDispatch();
        Dispatch sheet;

        if ("index".equalsIgnoreCase(selectSheetBy)) {
            sheet = Dispatch.call(sheets, "Item", sheetIndex.intValue()).toDispatch();
        } else {
            sheet = Dispatch.call(sheets, "Item", sheetName).toDispatch();
        }

        Dispatch range = Dispatch.call(sheet, "Range", rangeStr).toDispatch();

        StringBuilder sb = new StringBuilder();

        Dispatch rows = Dispatch.get(range, "Rows").toDispatch();
        int rowCount = Dispatch.get(rows, "Count").getInt();

        Dispatch cols = Dispatch.get(range, "Columns").toDispatch();
        int colCount = Dispatch.get(cols, "Count").getInt();

        for (int r = 1; r <= rowCount; r++) {
            for (int c = 1; c <= colCount; c++) {
                Dispatch cell = Dispatch.call(range, "Cells", r, c).toDispatch();
                Variant value = Dispatch.get(cell, "Value");
                String valStr = value != null && !value.isNull() ? value.toString().trim() : "";

                if (ignoreEmpty != null && ignoreEmpty) {
                    if (!valStr.isEmpty()) {
                        if (sb.length() > 0) sb.append(",");
                        sb.append(valStr);
                    }
                } else {
                    sb.append(valStr);
                    if (c < colCount || r < rowCount) {
                        sb.append(",");
                    }
                }
            }
        }

        return new StringValue(sb.toString());
    }
}
