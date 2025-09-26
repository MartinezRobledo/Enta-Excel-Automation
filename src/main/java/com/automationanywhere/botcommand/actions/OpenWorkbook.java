package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.*;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.Map;
import java.util.UUID;

@BotCommand
@CommandPkg(
        label = "Open Workbook",
        name = "openWorkbook",
        description = "Opens an Excel workbook and creates/attaches a session",
        icon = "excel.svg",
        return_type = DataType.SESSION,
        return_label = "Workbook Session",
        return_description = "Session variable associated with the opened workbook"
)
public class OpenWorkbook {

    @Execute
    public SessionValue action(
            @Idx(index = "1", type = AttributeType.FILE)
            @Pkg(label = "Workbook Path", description = "Full path to the Excel workbook")
            @NotEmpty
            String workbookPath,

            @Idx(index = "2", type = AttributeType.CHECKBOX)
            @Pkg(label = "Create as Global Session", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean createGlobal,

            @Idx(index = "3", type = AttributeType.CHECKBOX)
            @Pkg(label = "Attach if session already exists", default_value_type = DataType.BOOLEAN, default_value = "true")
            Boolean attachIfExists,

            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "Make Excel Visible", default_value_type = DataType.BOOLEAN, default_value = "true")
            Boolean visible
    ) {
        File file = new File(workbookPath);
        if (!file.exists()) {
            throw new BotCommandException("Workbook file does not exist: " + workbookPath);
        }

        // Mantener el sessionId por path (compat)
        String sessionId = "WB_" + Integer.toHexString(workbookPath.toLowerCase().hashCode());
        String workbookKey = SessionHelper.toWorkbookKey(workbookPath);

        // Si ya existe el mismo sessionId y attachIfExists=true, devolverlo (compat).
        Session existingById = SessionManager.getSession(sessionId);
        if (existingById != null) {
            if (Boolean.TRUE.equals(attachIfExists)) {
                // Aseguramos que si el libro no estaba en el mapa (raro), lo agregamos.
                if (!existingById.openWorkbooks.containsKey(workbookKey)) {
                    Dispatch workbooks = existingById.excelApp.getProperty("Workbooks").toDispatch();
                    Dispatch wb = Dispatch.call(workbooks, "Open", workbookPath).toDispatch();
                    existingById.openWorkbooks.put(workbookKey, wb);
                }
                return SessionValue.builder()
                        .withSessionObject(new ExcelSession(sessionId, existingById, workbookKey))
                        .build();
            } else {
                throw new BotCommandException("Session already exists for workbook: " + workbookPath);
            }
        }

        try {
            // Buscar si ya hay ALGUNA Session existente (compartimos Excel.Application)
            Session shared = null;
            for (Map.Entry<String, Session> e : SessionManager.getSessions().entrySet()) {
                if (e.getValue() != null && e.getValue().excelApp != null) {
                    shared = e.getValue();
                    break;
                }
            }

            if (shared == null) {
                // Primera vez: cargar JACOB, inicializar COM y crear Excel
                boolean is64Bit = System.getProperty("os.arch").contains("64");
                String dllName = is64Bit ? "BridgeCOM64.dll" : "BridgeCOM32.dll";

                InputStream dllStream = this.getClass().getClassLoader().getResourceAsStream("bridges/" + dllName);
                if (dllStream == null) {
                    throw new BotCommandException("DLL not found in resources/bridges/: " + dllName);
                }

                String tempDir = System.getProperty("java.io.tmpdir");
                File dllFile;

                if (Boolean.TRUE.equals(createGlobal)) {
                    dllFile = new File(tempDir, dllName);
                    if (!dllFile.exists()) {
                        try (FileOutputStream out = new FileOutputStream(dllFile)) {
                            byte[] buffer = new byte[1024];
                            int read;
                            while ((read = dllStream.read(buffer)) != -1) {
                                out.write(buffer, 0, read);
                            }
                        }
                    }
                    if (!JacobLoader.isLoaded()) {
                        JacobLoader.loadJacob(dllFile);
                    }
                } else {
                    String uniqueName = dllName.replace(".dll", "_" + UUID.randomUUID() + ".dll");
                    dllFile = new File(tempDir, uniqueName);
                    Files.copy(dllStream, dllFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    System.setProperty("jacob.dll.path", dllFile.getAbsolutePath());
                    com.jacob.com.LibraryLoader.loadJacobLibrary();
                }

                ComThread.InitSTA();
                ActiveXComponent excel = new ActiveXComponent("Excel.Application");
                excel.setProperty("Visible", Boolean.TRUE.equals(visible));

                shared = new Session(excel);
                shared.global = createGlobal;
            } else {
                // Ya existe Excel: opcionalmente actualizamos la visibilidad
                try {
                    shared.excelApp.setProperty("Visible", Boolean.TRUE.equals(visible));
                } catch (Exception ignore) {
                    // si falla, no interrumpimos
                }
            }

            // Abrir el workbook en la instancia compartida
            Dispatch workbooks = shared.excelApp.getProperty("Workbooks").toDispatch();
            Dispatch wb = Dispatch.call(workbooks, "Open", workbookPath).toDispatch();
            shared.openWorkbooks.put(workbookKey, wb);

            // Registrar el sessionId (mapeando al MISMO Session compartido)
            SessionManager.addSession(sessionId, shared);

            return SessionValue.builder()
                    .withSessionObject(new ExcelSession(sessionId, shared, workbookKey))
                    .build();

        } catch (Exception e) {
            throw new BotCommandException("Failed to open workbook: " + e.getMessage());
        }
    }
}