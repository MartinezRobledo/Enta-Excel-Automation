package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSessionManager;
import com.automationanywhere.botcommand.utilities.ExcelSession;
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
        icon = "excel.svg"
)
public class OpenWorkbook {

    @Execute
    public void action(
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

        ExcelSession session = ExcelSessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null) {
            throw new BotCommandException("Session not found: " + sessionId);
        }

        File file = new File(workbookPath);
        if (!file.exists()) {
            throw new BotCommandException("Workbook file does not exist: " + workbookPath);
        }

        Dispatch workbooks = session.excelApp.getProperty("Workbooks").toDispatch();

        // ✅ Si ya lo tenemos en openWorkbooks, simplemente terminamos
        if (session.openWorkbooks.containsKey(workbookPath)) {
            session.excelApp.setProperty("Visible", Boolean.TRUE.equals(visible));
            return;
        }

        // ✅ Buscar si está abierto en Excel
        int count = Dispatch.get(workbooks, "Count").getInt();
        for (int i = 1; i <= count; i++) {
            Dispatch wb = Dispatch.call(workbooks, "Item", i).toDispatch();
            String fullName = Dispatch.get(wb, "FullName").getString();
            if (fullName.equalsIgnoreCase(workbookPath)) {
                // ✅ Adjuntar el workbook existente en la sesión
                session.openWorkbooks.put(workbookPath, wb);
                session.excelApp.setProperty("Visible", Boolean.TRUE.equals(visible));
                return;
            }
        }

        // ✅ Si no está abierto, lo abrimos
        try {
            Dispatch wb = Dispatch.call(workbooks, "Open", workbookPath).toDispatch();
            session.excelApp.setProperty("Visible", Boolean.TRUE.equals(visible));
            session.openWorkbooks.put(workbookPath, wb);
        } catch (Exception e) {
            throw new BotCommandException("Failed to open workbook: " + e.getMessage());
        }
    }

}