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

        // Si es una sesión global y el usuario eligió mantener abierto, NO cerrar el workbook ni Excel, ni tocar mapas
        if (Boolean.TRUE.equals(session.global) && Boolean.TRUE.equals(keepOpenGlobal)) {
            // (Opcional) Si quisieras permitir guardar sin cerrar en este caso:
            // if (Boolean.TRUE.equals(saveBeforeClose)) { try { Dispatch.call(wb, "Save"); } catch (Exception ignore) {} }
            return;
        }

        boolean closeSucceeded = false;

        // Desactivar alerts para evitar prompts (e.g. "¿Guardar cambios?")
        Dispatch app = session.excelApp;
        Boolean prevAlerts = null;
        try {
            try {
                prevAlerts = Dispatch.get(app, "DisplayAlerts").getBoolean();
            } catch (Exception ignore) { /* si falla, lo dejamos nulo */ }

            try {
                if (prevAlerts != null) {
                    Dispatch.put(app, "DisplayAlerts", false);
                }
            } catch (Exception ignore) { }

            // Guardar si aplica
            if (Boolean.TRUE.equals(saveBeforeClose)) {
                try {
                    Dispatch.call(wb, "Save");
                } catch (Exception ignore) { }
            }

            // Cerrar workbook: usamos el flag de guardado para reforzar
            try {
                Dispatch.call(wb, "Close", new Variant(Boolean.TRUE.equals(saveBeforeClose)));
                closeSucceeded = true;
            } catch (Exception e) {
                // Propagamos el error sin tocar el mapa; el libro sigue registrado
                throw e;
            }
        } finally {
            // Restaurar DisplayAlerts si lo pudimos leer antes
            if (prevAlerts != null) {
                try { Dispatch.put(app, "DisplayAlerts", prevAlerts); } catch (Exception ignore) { }
            }
        }

        // Remover del mapa SOLO si efectivamente se cerró el libro
        if (closeSucceeded) {
            session.openWorkbooks.remove(workbookKey);
        }

        // Primero remover SOLO este sessionId de SessionManager
        SessionManager.removeSessionIdOnly(sessionId);

        // Evaluar cierre de Excel: debe cumplirse TODO:
        // 1) Cerró el libro actual
        // 2) Ya no quedan libros en 'openWorkbooks' de esta sesión
        // 3) O la sesión NO es global, o es global pero NO se pidió mantener abierto
        // 4) Ya NO quedan referencias a esta instancia en el SessionManager
        boolean shouldCloseExcelAfter =
                closeSucceeded
                        && session.openWorkbooks.isEmpty()
                        && !(Boolean.TRUE.equals(session.global) && Boolean.TRUE.equals(keepOpenGlobal))
                        && !SessionManager.hasRefs(session);

        if (shouldCloseExcelAfter) {
            try {
                Dispatch.call(session.excelApp, "Quit");
            } catch (Exception ignore) {
            } finally {
                // En este punto, al no quedar refs, removeAllByInstance sería redundante.
                // Si tu implementación requiere limpieza extra del manager, podrías invocarla aquí.
                try { com.jacob.com.ComThread.Release(); } catch (Exception ignore) {}
            }
        }
    }
}