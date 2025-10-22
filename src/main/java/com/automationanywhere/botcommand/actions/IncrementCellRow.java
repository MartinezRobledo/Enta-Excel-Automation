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
        label = "Increment Row in Cell Address",
        name = "incrementCellRow",
        description = "Increments the row of an A1-style cell reference (e.g., A2 -> A3). Preserves $ if present.",
        return_type = DataType.STRING,
        return_label = "New cell address",
        return_required = true,
        icon = "excel.svg"
)
public class IncrementCellRow {

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Cell (A1-style)", description = "Examples: A2, $A$2, A$2, $A2")
            @NotEmpty
            String cellAddress,

            @Idx(index = "2", type = AttributeType.NUMBER)
            @Pkg(label = "Increment by (rows)", default_value = "1", default_value_type = DataType.NUMBER)
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double stepRows
    ) {
        try {
            String out = ExcelHelpers.incrementRowInA1(cellAddress, stepRows.intValue());
            return new StringValue(out);
        } catch (BotCommandException e) {
            throw e;
        } catch (Exception e) {
            throw new BotCommandException("Failed to increment row for '" + cellAddress + "': " + e.getMessage(), e);
        }
    }
}