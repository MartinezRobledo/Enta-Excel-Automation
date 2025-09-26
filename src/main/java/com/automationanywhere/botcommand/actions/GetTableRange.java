package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
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

@BotCommand
@CommandPkg(
        label = "Get Full Data Range (from headers)",
        name = "getFullDataRangeFromHeaders",
        description = "Given a headers range (e.g., C9:BM9), returns the full data rectangle including or excluding headers, down to the last used row.",
        icon = "excel.svg",
        return_type = DataType.STRING,
        return_required = true
)
public class GetTableRange {

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session") @SessionObject @NotEmpty ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name") String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (1-based)") Double sheetIndex,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Headers Range (e.g., C9:BM9)") @NotEmpty String headersRange,

            @Idx(index = "4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "No", value = "no"))
            })
            @Pkg(label = "Include headers?", default_value = "no", default_value_type = DataType.STRING)
            @SelectModes String includeHeaders
    ) {
        try {
            return run(excelSession, selectSheetBy, sheetName, sheetIndex, headersRange, includeHeaders);
        } catch (Exception first) {
            try {
                ComThread.InitSTA();
                return run(excelSession, selectSheetBy, sheetName, sheetIndex, headersRange, includeHeaders);
            } catch (Exception second) {
                throw (second instanceof BotCommandException)
                        ? (BotCommandException) second
                        : new BotCommandException("Failed to get full data range: " + second.getMessage(), second);
            } finally {
                try { ComThread.Release(); } catch (Exception ignore) {}
            }
        }
    }

    private Value<String> run(ExcelSession excelSession, String selectBy, String sheetName, Double sheetIndex,
                              String headersRange, String includeHeaders) {

        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheet = ExcelObjects.requireSheet(wb, selectBy, sheetName, sheetIndex);

        String[] parts = headersRange.split(":");
        if (parts.length != 2) throw new BotCommandException("Invalid headers range: " + headersRange);

        String startCell = parts[0].trim();        // e.g. C9
        String endCell   = parts[1].trim();        // e.g. BM9
        String startCol  = startCell.replaceAll("\\d", "");
        String endCol    = endCell.replaceAll("\\d", "");

        int headerRow;
        try {
            headerRow = Integer.parseInt(endCell.replaceAll("\\D", ""));
        } catch (Exception e) {
            throw new BotCommandException("Invalid headers end cell (row not found): " + endCell);
        }

        Dispatch used = Dispatch.get(sheet, "UsedRange").toDispatch();
        int usedFirstRow = Dispatch.get(used, "Row").getInt();
        int usedRows = Dispatch.get(Dispatch.get(used, "Rows").toDispatch(), "Count").getInt();
        int lastRow = usedFirstRow + usedRows - 1;

        int topRow = "yes".equalsIgnoreCase(includeHeaders) ? headerRow : headerRow + 1;
        if (lastRow < topRow) return new StringValue("");

        String fullRange = startCol + topRow + ":" + endCol + lastRow;
        return new StringValue(fullRange);
    }
}