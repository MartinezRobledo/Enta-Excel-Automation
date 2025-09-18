package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.NumberValue;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.ExcelHelpers;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
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
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select origin sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Sheet Name")
            @NotEmpty
            String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Origin Sheet Index", description = "1-based index")
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double sheetIndex,

            @Idx(index = "3", type = AttributeType.VARIABLE)
            @Pkg(label = "Select variable")
            @NotEmpty
            Value<Double> varOutput
    ) {
        Session session = excelSession.getSession();
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found o closed");

        if (session.openWorkbooks.isEmpty())
            throw new BotCommandException("No workbook is open in this session.");

        Dispatch wb = session.openWorkbooks.values().iterator().next();

        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
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
