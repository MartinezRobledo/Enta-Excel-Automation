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
        description = "Closes a workbook from the shared Excel session",
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
        if (excelSession == null) {
            throw new BotCommandException("Workbook Session is null.");
        }

        Session session = excelSession.getSession();
        String sessionId = excelSession.getSessionId();
        String workbookKey = excelSession.getWorkbookKey();

        if (session == null || session.excelApp == null) {
            throw new BotCommandException("Session not found: " + sessionId);
        }

        // Obtener SOLO el workbook asociado a esta ExcelSession
        Dispatch wb = session.openWorkbooks.get(workbookKey);
        if (wb == null) {
            throw new BotCommandException("Workbook not tracked/open in session: " + workbookKey);
        }

        // Si es una sesión global y el usuario eligió mantener abierto, NO cerrar el workbook ni Excel
        if (Boolean.TRUE.equals(session.global) && Boolean.TRUE.equals(keepOpenGlobal)) {
            // (Opcional) Si querés permitir guardar sin cerrar en este caso:
            // if (Boolean.TRUE.equals(saveBeforeClose)) { try { Dispatch.call(wb, "Save"); } catch (Exception ignore) {} }
            return; // no cambiamos mapas ni sessionIds
        }

        // Caso normal: cerrar SOLO este workbook
        boolean shouldCloseExcelAfter = false;

        try {
            if (Boolean.TRUE.equals(saveBeforeClose)) {
                try { Dispatch.call(wb, "Save"); } catch (Exception ignore) {}
            }
            try { Dispatch.call(wb, "Close", new Variant(false)); } catch (Exception ignore) {}
        } finally {
            // Remover del mapa si lo cerramos
            session.openWorkbooks.remove(workbookKey);
        }

        // Si no quedan libros, evaluar cierre de Excel
        if (session.openWorkbooks.isEmpty()) {
            // Cerrar Excel si: no es global, o es global pero NO se pidió mantener abierto
            shouldCloseExcelAfter = !(Boolean.TRUE.equals(session.global) && Boolean.TRUE.equals(keepOpenGlobal));
        }

        if (shouldCloseExcelAfter) {
            try {
                Dispatch.call(session.excelApp, "Quit");
            } catch (Exception ignore) {
            } finally {
                // Remover TODAS las referencias a esta instancia de Session
                SessionManager.removeAllByInstance(session);
            }
        } else {
            // Remover solo este sessionId si cerramos el workbook pero Excel sigue abierto
            SessionManager.removeSessionIdOnly(sessionId);
        }
    }
}