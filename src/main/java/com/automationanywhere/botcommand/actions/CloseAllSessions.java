package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.Set;

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
        try {
            Set<Session> distinct = SessionManager.getDistinctSessions();

            for (Session session : distinct) {
                if (session == null || session.excelApp == null) continue;

                try {
                    // Cerrar todos los libros abiertos en esta instancia
                    for (var wbEntry : session.openWorkbooks.entrySet()) {
                        try {
                            if (Boolean.TRUE.equals(saveBeforeClose)) {
                                Dispatch.call(wbEntry.getValue(), "Save");
                            }
                            Dispatch.call(wbEntry.getValue(), "Close", new Variant(false));
                        } catch (Exception ignore) {}
                    }
                    session.openWorkbooks.clear();

                    // Cerrar Excel
                    try {
                        Dispatch.call(session.excelApp, "Quit");
                    } catch (Exception ignore) {}

                } catch (Exception ignore) {
                } finally {
                    // Remover todas las refs a esta instancia
                    SessionManager.removeAllByInstance(session);
                }
            }

            // ⚠️ Evitar matar procesos a la fuerza (puede cerrar instancias ajenas a este bot).
            // Si querés mantenerlo por compat, dejalo, pero no es recomendable:
            // Runtime.getRuntime().exec("taskkill /F /IM EXCEL.EXE");

        } catch (Exception ignored) {}
    }
}