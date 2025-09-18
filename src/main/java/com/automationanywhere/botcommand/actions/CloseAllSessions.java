package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.Map;


@BotCommand
@CommandPkg(
        label = "Close All Sessions",
        name = "closeAllWorkbookSessions",
        description = "Closes all active workbook sessions",
        icon = "excel.svg"
)
public class CloseAllSessions {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.CHECKBOX)
            @Pkg(label = "Save all before closing", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean saveBeforeClose
    ) {
        for (Map.Entry<String, Session> entry : SessionManager.getSessions().entrySet()) {
            String sessionId = entry.getKey();
            Session session = entry.getValue();

            try {
                for (Dispatch wb : session.openWorkbooks.values()) {
                    if (Boolean.TRUE.equals(saveBeforeClose)) {
                        Dispatch.call(wb, "Save");
                    }
                    Dispatch.call(wb, "Close", new Variant(false));
                }

                Dispatch.call(session.excelApp, "Quit");

            } catch (Exception e) {
                // Ignorar errores
            } finally {
                SessionManager.removeSession(sessionId);
            }
        }

        try {
            Runtime.getRuntime().exec("taskkill /F /IM EXCEL.EXE");
        } catch (Exception ignored) {}
    }
}
