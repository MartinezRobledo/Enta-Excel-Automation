package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.NumberValue;
import com.automationanywhere.botcommand.utilities.ExcelHelpers;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;

@BotCommand
@CommandPkg(
        label = "Get Number of Rows in column",
        name = "getNumberOfRowsInColumn",
        description = "Returns the number of rows with data in a column",
        return_type = DataType.NUMBER,
        return_label = "Number of rows in column",
        return_required = true,
        icon = "excel.svg"
)
public class GetNumberOfRowsByColumn {

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

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Header", value = "header")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Letter", value = "letter"))
            })
            @Pkg(label = "Select Column By", default_value = "letter", default_value_type = DataType.STRING)
            @SelectModes String selectColumnBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Column Header Name")
            @NotEmpty String columnName,

            @Idx(index = "3.2.1", type = AttributeType.TEXT)
            @Pkg(label = "Column Letter (A, B, ...)")
            @NotEmpty String columnLetter
    ) {
        // 1) Sesión + workbook correctos
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);

        // 2) Resolver hoja (valida nombre/índice y lanza errores claros)
        Dispatch sheet = ExcelObjects.requireSheet(wb, selectSheetBy, sheetName, sheetIndex);
        try { Dispatch.call(sheet, "Activate"); } catch (Exception ignore) {}

        int rows;
        if("letter".equalsIgnoreCase(selectColumnBy))
            rows = ExcelHelpers.getLastDataRowInColumn(sheet, columnLetter);
        else {
            int cantCols = ExcelHelpers.getLastColumn(sheet);
            int colIndex = ExcelHelpers.headerNameToColumnIndex(sheet, columnName, 1, cantCols);
            rows = ExcelHelpers.getLastDataRowInColumn(sheet, colIndex);
        }

        return new NumberValue((double) rows);
    }
}
