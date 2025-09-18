package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;

@BotCommand
@CommandPkg(
        label = "Save Workbook As",
        name = "saveWorkbookAs",
        description = "Saves an open workbook to a new file path",
        icon = "excel.svg"
)
public class SaveWorkbookAs {

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.FILE)
            @Pkg(label = "Destination Path")
            @NotEmpty
            String destPath
    ) {
        Session session = excelSession.getSession();
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found o closed");

        if (session.openWorkbooks.isEmpty())
            throw new BotCommandException("No workbook is open in this session.");

        Dispatch wb = session.openWorkbooks.values().iterator().next();

        Dispatch.call(wb, "SaveAs", destPath);
        return new StringValue("Workbook saved as: " + destPath);
    }
}
