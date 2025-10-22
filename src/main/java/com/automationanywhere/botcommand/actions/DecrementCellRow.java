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
        label = "Decrement Row in Cell Address",
        name = "decrementCellRow",
        description = "Decrements the row of an A1-style cell reference (e.g., A3 -> A2). Preserves $ if present.",
        return_type = DataType.STRING,
        return_label = "New cell address",
        return_required = true,
        icon = "excel.svg"
)
public class DecrementCellRow {

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Cell (A1-style)", description = "Examples: A3, $A$3, A$3, $A3")
            @NotEmpty
            String cellAddress,

            @Idx(index = "2", type = AttributeType.NUMBER)
            @Pkg(label = "Decrement by (rows)", default_value = "1", default_value_type = DataType.NUMBER)
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double stepRows
    ) {
        try {
            // Parse A1, restar filas y reconstruir preservando $.
            ExcelHelpers.A1Ref r = ExcelHelpers.parseA1(cellAddress);
            long newRow = (long) r.row - stepRows.intValue();

            if (newRow < 1 || newRow > ExcelHelpers.EXCEL_MAX_ROWS) {
                throw new BotCommandException("Row out of range for Excel after decrement: " + newRow);
            }

            ExcelHelpers.A1Ref r2 = new ExcelHelpers.A1Ref((int) newRow, r.col, r.absRow, r.absCol);
            String out = ExcelHelpers.buildA1(r2);
            return new StringValue(out);

        } catch (BotCommandException e) {
            throw e;
        } catch (Exception e) {
            throw new BotCommandException(
                    "Failed to decrement row for '" + cellAddress + "': " + e.getMessage(), e);
        }
    }
}