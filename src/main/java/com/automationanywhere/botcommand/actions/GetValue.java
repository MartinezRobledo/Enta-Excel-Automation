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
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Get Value",
        name = "getValue",
        description = "Returns the value of a cell given its address (e.g., A1). Optionally returns formatted text.",
        icon = "excel.svg",
        return_type = DataType.STRING,
        return_required = true
)
public class GetValue {

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @SessionObject @NotEmpty ExcelSession excelSession,

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
            @Pkg(label = "Cell Address (e.g., A1)") @NotEmpty String cellAddress,

            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "Return formatted text", description = "If checked, returns the cell as displayed in Excel",
                    default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean formatted
    ) {
        try {
            return run(excelSession, selectSheetBy, sheetName, sheetIndex, cellAddress, formatted);
        } catch (Exception first) {
            try {
                ComThread.InitSTA();
                return run(excelSession, selectSheetBy, sheetName, sheetIndex, cellAddress, formatted);
            } catch (Exception second) {
                throw (second instanceof BotCommandException)
                        ? (BotCommandException) second
                        : new BotCommandException("Failed to get cell value: " + second.getMessage(), second);
            } finally {
                try { ComThread.Release(); } catch (Exception ignore) {}
            }
        }
    }

    private Value<String> run(ExcelSession excelSession, String selectBy, String sheetName, Double sheetIndex,
                              String cellAddress, Boolean formatted) {
        if (cellAddress == null || cellAddress.trim().isEmpty()) {
            throw new BotCommandException("Cell address cannot be empty.");
        }

        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);
        Dispatch sheet = ExcelObjects.requireSheet(wb, selectBy, sheetName, sheetIndex);

        Dispatch cell = Dispatch.call(sheet, "Range", cellAddress.trim()).toDispatch();
        Variant val = formatted ? Dispatch.get(cell, "Text") : Dispatch.get(cell, "Value2");

        String result = (val == null || val.isNull()) ? "" : val.toString();
        return new StringValue(result);
    }
}