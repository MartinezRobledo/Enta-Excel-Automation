package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;

@BotCommand
@CommandPkg(
        label = "Unmerge Range",
        name = "unmergeRange",
        description = "Unmerge all merged cells in the given range",
        icon = "excel.svg",
        return_type = DataType.BOOLEAN,
        return_required = true
)
public class UnmergeRange {

    @Execute
    public com.automationanywhere.botcommand.data.Value<Boolean> action(
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
            @Pkg(label = "Range (e.g., A1:C10)") @NotEmpty String rangeA1
    ) {
        try {
            return run(excelSession, selectSheetBy, sheetName, sheetIndex, rangeA1);
        } catch (Exception first) {
            try {
                ComThread.InitSTA();
                return run(excelSession, selectSheetBy, sheetName, sheetIndex, rangeA1);
            } catch (Exception second) {
                throw (second instanceof BotCommandException)
                        ? (BotCommandException) second
                        : new BotCommandException("Failed to unmerge range: " + second.getMessage(), second);
            } finally {
                try { ComThread.Release(); } catch (Exception ignore) {}
            }
        }
    }

    private Value<Boolean> run(
            ExcelSession excelSession, String selectBy, String sheetName, Double sheetIndex, String rangeA1) {

        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);

        Dispatch sheet = ExcelObjects.requireSheet(wb, selectBy, sheetName, sheetIndex);

        Dispatch rng = Dispatch.call(sheet, "Range", rangeA1.trim()).toDispatch();
        if (rng == null || rng.m_pDispatch == 0) {
            throw new BotCommandException("Invalid range: " + rangeA1);
        }

        try { Dispatch.call(rng, "UnMerge"); } catch (Exception ignore) {}
        return new com.automationanywhere.botcommand.data.impl.BooleanValue(true);
    }
}
