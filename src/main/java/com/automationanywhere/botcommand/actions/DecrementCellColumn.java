package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelHelpers;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

@BotCommand
@CommandPkg(
        label = "Decrement Column in Cell Address",
        name = "decrementCellColumn",
        description = "Decrements the column of an A1-style cell reference (e.g., C5 -> B5). Preserves $ if present.",
        return_type = DataType.STRING,
        return_label = "New cell address",
        return_required = true,
        icon = "excel.svg"
)
public class DecrementCellColumn {

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Cell (A1-style)", description = "Examples: C5, $C$5, C$5, $C5")
            @NotEmpty
            String cellAddress,

            @Idx(index = "2", type = AttributeType.NUMBER)
            @Pkg(label = "Decrement by (columns)", default_value = "1", default_value_type = DataType.NUMBER)
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double stepCols
    ) {
        try {
            // Parse A1, restar columnas y reconstruir preservando $.
            ExcelHelpers.A1Ref r = ExcelHelpers.parseA1(cellAddress);
            long newCol = (long) r.col - stepCols.intValue();

            if (newCol < 1 || newCol > ExcelHelpers.EXCEL_MAX_COLS) {
                throw new BotCommandException("Column out of range for Excel after decrement: " + newCol);
            }

            ExcelHelpers.A1Ref r2 = new ExcelHelpers.A1Ref(r.row, (int) newCol, r.absRow, r.absCol);
            String out = ExcelHelpers.buildA1(r2);
            return new StringValue(out);

        } catch (BotCommandException e) {
            throw e;
        } catch (Exception e) {
            throw new BotCommandException(
                    "Failed to decrement column for '" + cellAddress + "': " + e.getMessage(), e);
        }
    }
}
