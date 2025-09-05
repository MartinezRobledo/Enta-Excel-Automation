package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
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
        label = "Get Headers Range",
        name = "getHeadersRange",
        description = "Finds the row where a reference header exists and returns the contiguous headers range in that row. Returns empty string if not found.",
        icon = "excel.svg"
)
public class GetHeadersRange {

    @Execute
    public Value<String> action(
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
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty
            String sheetName,

            @Idx(index = "3.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index", description = "1-based index")
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double sheetIndex,

            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Reference Header", description = "Exact header text to find")
            @NotEmpty
            String referenceHeader,

            @Idx(index = "5", type = AttributeType.VARIABLE)
            @Pkg(label = "Select variable")
            @NotEmpty
            Value<String> varOutput
    ) {
        ExcelSession session = ExcelSessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found: " + sessionId);

        Dispatch workbook = session.openWorkbooks.get(workbookName);
        if (workbook == null)
            throw new BotCommandException("Workbook not open: " + workbookName);

        // Hoja
        Dispatch sheets = Dispatch.get(workbook, "Sheets").toDispatch();
        Dispatch sheet = "index".equalsIgnoreCase(selectSheetBy)
                ? Dispatch.call(sheets, "Item", sheetIndex.intValue()).toDispatch()
                : Dispatch.call(sheets, "Item", sheetName).toDispatch();

        // UsedRange y límites (ojo: Rows/Columns devuelven Range; hay que pedir Count)
        Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();
        int usedFirstRow = Dispatch.get(usedRange, "Row").getInt();
        int usedFirstCol = Dispatch.get(usedRange, "Column").getInt();
        int usedRows = Dispatch.get(Dispatch.get(usedRange, "Rows").toDispatch(), "Count").getInt();
        int usedCols = Dispatch.get(Dispatch.get(usedRange, "Columns").toDispatch(), "Count").getInt();
        int usedLastCol = usedFirstCol + usedCols - 1;

        // Buscar el header
        Dispatch findResult;
        try {
            findResult = Dispatch.call(usedRange, "Find", referenceHeader).toDispatch();
        } catch (Exception e) {
            // Excel puede lanzar excepción si no encuentra; tratamos como no encontrado
            findResult = null;
        }

        if (findResult == null || findResult.m_pDispatch == 0) {
            varOutput.set("");
            return new StringValue("");
        }

        int headerRow = Dispatch.get(findResult, "Row").getInt();
        int foundCol = Dispatch.get(findResult, "Column").getInt();

        // Expandir a izquierda hasta celda vacía (limitado por UsedRange)
        int startCol = foundCol;
        while (startCol - 1 >= usedFirstCol) {
            Dispatch leftCell = Dispatch.call(sheet, "Cells", headerRow, startCol - 1).toDispatch();
            Variant v = Dispatch.get(leftCell, "Value");
            if (v.isNull() || v.toString().trim().isEmpty()) break;
            startCol--;
        }

        // Expandir a derecha hasta celda vacía (limitado por UsedRange)
        int endCol = foundCol;
        while (endCol + 1 <= usedLastCol) {
            Dispatch rightCell = Dispatch.call(sheet, "Cells", headerRow, endCol + 1).toDispatch();
            Variant v = Dispatch.get(rightCell, "Value");
            if (v.isNull() || v.toString().trim().isEmpty()) break;
            endCol++;
        }

        String startColLetter = getColumnLetter(startCol);
        String endColLetter = getColumnLetter(endCol);
        String range = startColLetter + headerRow + ":" + endColLetter + headerRow;

        varOutput.set(range);
        return new StringValue(range);
    }

    private String getColumnLetter(int columnNumber) {
        StringBuilder sb = new StringBuilder();
        while (columnNumber > 0) {
            int rem = (columnNumber - 1) % 26;
            sb.insert(0, (char) ('A' + rem));
            columnNumber = (columnNumber - 1) / 26;
        }
        return sb.toString();
    }
}
