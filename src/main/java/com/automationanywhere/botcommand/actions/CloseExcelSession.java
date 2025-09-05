package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSessionManager;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;


@BotCommand
@CommandPkg(
        label = "Close Session",
        name = "closeExcelSession",
        description = "Closes an Excel COM session",
        icon = "excel.svg"
)
public class CloseExcelSession {

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Session ID", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty
            String sessionId,

            @Idx(index = "2", type = AttributeType.CHECKBOX)
            @Pkg(label = "Save all open workbooks before closing", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean saveAll
    ) {
        ExcelSession session = ExcelSessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null) {
            throw new BotCommandException("Session not found: " + sessionId);
        }

        try {
            Dispatch workbooks = session.excelApp.getProperty("Workbooks").toDispatch();

            int count = Dispatch.get(workbooks, "Count").getInt();
            for (int i = count; i >= 1; i--) { // recorrer de atrás hacia adelante
                Dispatch wb = Dispatch.call(workbooks, "Item", i).toDispatch();
                if (Boolean.TRUE.equals(saveAll)) {
                    Dispatch.call(wb, "Save");
                }
                // cerrar sin guardar (ignorar cambios si saveAll=false)
                Dispatch.call(wb, "Close", new Variant(Boolean.FALSE));
            }

            Dispatch.call(session.excelApp, "Quit");

        } catch (Exception e) {
            // ignorar errores y pasar a taskkill
        } finally {
            ExcelSessionManager.removeSession(sessionId);
            try {
                Runtime.getRuntime().exec("taskkill /F /IM EXCEL.EXE");
            } catch (Exception ex) {
                // ignorar errores
            }
        }

        return new StringValue("Excel session closed: " + sessionId);
    }
}
