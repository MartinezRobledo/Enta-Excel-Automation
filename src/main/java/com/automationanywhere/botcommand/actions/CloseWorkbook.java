package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Close Workbook",
        name = "closeWorkbookSession",
        description = "Closes a workbook and its associated session",
        icon = "excel.svg"
)
public class CloseWorkbook {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.CHECKBOX)
            @Pkg(label = "Save before closing", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean saveBeforeClose,

            @Idx(index = "3", type = AttributeType.CHECKBOX)
            @Pkg(label = "Keep open if session is global", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean keepOpenGlobal
    ) {
        Session session = excelSession.getSession();
        String sessionId = excelSession.getSessionId();

        if (session == null || session.excelApp == null) {
            throw new BotCommandException("Session not found: " + sessionId);
        }

        boolean shouldClose = !session.global || Boolean.FALSE.equals(keepOpenGlobal);

        if (shouldClose) {
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
                excelSession.close();
                try {
                    Runtime.getRuntime().exec("taskkill /F /IM EXCEL.EXE");
                } catch (Exception ignored) {}
            }
        }
    }
}
