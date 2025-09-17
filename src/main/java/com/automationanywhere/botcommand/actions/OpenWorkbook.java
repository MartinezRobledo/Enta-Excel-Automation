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

import java.io.File;

@BotCommand
@CommandPkg(
        label = "Open Excel Workbook",
        name = "openWorkbook",
        description = "Opens an Excel workbook in an existing COM session",
        icon = "excel.svg",
        return_type = DataType.STRING,
        return_label = "Workbook Name",
        return_description = "The full name of the workbook as Excel recognizes it"
)
public class OpenWorkbook {

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.FILE)
            @Pkg(label = "Workbook Path", description = "Full path to the Excel workbook")
            @NotEmpty
            String workbookPath,

            @Idx(index = "2", type = AttributeType.CHECKBOX)
            @Pkg(label = "Visible", description = "Make Excel visible?", default_value_type = DataType.BOOLEAN, default_value = "true")
            Boolean visible,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Session ID", description = "Excel session to use", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty
            String sessionId
    ) {

        Session session = SessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null) {
            throw new BotCommandException("Session not found: " + sessionId);
        }

        File file = new File(workbookPath);
        if (!file.exists()) {
            throw new BotCommandException("Workbook file does not exist: " + workbookPath);
        }

        Dispatch workbooks = session.excelApp.getProperty("Workbooks").toDispatch();
        Dispatch wb = null;

        // Buscar si ya está en la sesión
        for (Dispatch existingWb : session.openWorkbooks.values()) {
            String fullName = Dispatch.get(existingWb, "FullName").getString();
            if (fullName.equalsIgnoreCase(workbookPath)) {
                wb = existingWb;
                break;
            }
        }

        // Buscar si ya está abierto en Excel
        if (wb == null) {
            int count = Dispatch.get(workbooks, "Count").getInt();
            for (int i = 1; i <= count; i++) {
                Dispatch existingWb = Dispatch.call(workbooks, "Item", i).toDispatch();
                String fullName = Dispatch.get(existingWb, "FullName").getString();
                if (fullName.equalsIgnoreCase(workbookPath)) {
                    wb = existingWb;
                    break;
                }
            }
        }

        // Si no está abierto, abrirlo con reintentos
        if (wb == null) {
            wb = Dispatch.call(workbooks, "Open", workbookPath).toDispatch();

            int retries = 960; // máx 480 segundos
            while (retries-- > 0) {
                try {
                    String fullName = Dispatch.get(wb, "FullName").getString();
                    if (fullName != null && !fullName.isEmpty()) {
                        break;
                    }
                } catch (Exception ignored) {}
                try { Thread.sleep(500); } catch (InterruptedException ignored) {}
            }
        }

        // FullName de Excel como clave
        String fullName = Dispatch.get(wb, "FullName").getString();
        session.openWorkbooks.put(fullName, wb);

        // Ajustar visibilidad
        session.excelApp.setProperty("Visible", Boolean.TRUE.equals(visible));

        // WorkbookName para usar en otras acciones
        return new StringValue(fullName);
    }
}
