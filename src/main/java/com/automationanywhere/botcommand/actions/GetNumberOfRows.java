package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.NumberValue;
import com.automationanywhere.botcommand.utilities.ExcelSessionManager;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.ExcelHelpers;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.GreaterThanEqualTo;
import com.automationanywhere.commandsdk.annotations.rules.NumberInteger;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;

@BotCommand
@CommandPkg(
        label = "Get Number of Rows",
        name = "getNumberOfRows",
        description = "Returns the number of rows with data in a sheet",
        icon = "excel.svg"
)
public class GetNumberOfRows {

    @Execute
    public Value<Double> action(
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
            @Pkg(label = "Origin Sheet Index", description = "1-based index")
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double sheetIndex,

            @Idx(index = "4", type = AttributeType.VARIABLE)
            @Pkg(label = "Select variable")
            @NotEmpty
            Value<Double> varOutput
    ) {
        ExcelSession session = ExcelSessionManager.getSession(sessionId);
        if(session == null || session.excelApp == null) {
            throw new BotCommandException("Session not found: " + sessionId);
        }

        Dispatch workbook = session.openWorkbooks.get(workbookName);
        if(workbook == null) {
            throw new BotCommandException("Workbook not open in session: " + workbookName);
        }

        Dispatch sheets = Dispatch.get(workbook, "Sheets").toDispatch();
        Dispatch sheet;

        if("index".equalsIgnoreCase(selectSheetBy)) {
            if(sheetIndex == null) {
                throw new BotCommandException("Sheet index must be provided.");
            }
            sheet = Dispatch.call(sheets, "Item", sheetIndex.intValue()).toDispatch();
        } else {
            if(sheetName == null || sheetName.isEmpty()) {
                throw new BotCommandException("Sheet name must be provided.");
            }
            sheet = Dispatch.call(sheets, "Item", sheetName).toDispatch();
        }

        int rows = ExcelHelpers.getLastRow(sheet);
        varOutput.set(((double) rows));
        return new NumberValue((double) rows);
    }
}
