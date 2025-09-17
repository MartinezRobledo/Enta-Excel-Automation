package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
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
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Session ID", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty
            String sessionId,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Workbook Name")
            @NotEmpty
            String workbookName,

            @Idx(index = "3", type = AttributeType.FILE)
            @Pkg(label = "Destination Path")
            @NotEmpty
            String destPath
    ) {
        Session session = SessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found: " + sessionId);

        Dispatch wb = session.openWorkbooks.get(workbookName);
        if (wb == null)
            throw new BotCommandException("Workbook not open: " + workbookName);

        Dispatch.call(wb, "SaveAs", destPath);
        return new StringValue("Workbook saved as: " + destPath);
    }
}
