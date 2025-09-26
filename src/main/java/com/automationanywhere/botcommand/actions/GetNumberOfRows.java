package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.NumberValue;
import com.automationanywhere.botcommand.utilities.*;
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
        // 1) Sesión + workbook correctos
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);

        // 2) Resolver hoja (valida nombre/índice y lanza errores claros)
        Dispatch sheet = ExcelObjects.requireSheet(wb, selectSheetBy, sheetName, sheetIndex);
        try { Dispatch.call(sheet, "Activate"); } catch (Exception ignore) {}

        int rows = ExcelHelpers.getLastDataRow(sheet);
        varOutput.set(((double) rows));
        return new NumberValue((double) rows);
    }
}
