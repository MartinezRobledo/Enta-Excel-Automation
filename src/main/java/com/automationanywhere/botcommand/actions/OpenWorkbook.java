package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.JacobLoader;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.Session;
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
import java.util.UUID;

@BotCommand
@CommandPkg(
        label = "Open Workbook",
        name = "openWorkbook",
        description = "Opens an Excel workbook and creates a dedicated session for it",
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

        String sessionId = "WB_" + Integer.toHexString(workbookPath.toLowerCase().hashCode());

        Session existingSession = SessionManager.getSession(sessionId);
        if (existingSession != null) {
            if (Boolean.TRUE.equals(attachIfExists)) {
                return SessionValue
                        .builder()
                        .withSessionObject(new ExcelSession(sessionId, existingSession))
                        .build();

            } else {
                throw new BotCommandException("Session already exists for workbook: " + workbookPath);
            }
        }

        boolean is64Bit = System.getProperty("os.arch").contains("64");
        String dllName = is64Bit ? "BridgeCOM64.dll" : "BridgeCOM32.dll";

        InputStream dllStream = this.getClass().getClassLoader().getResourceAsStream("bridges/" + dllName);
        if (dllStream == null) {
            throw new BotCommandException("DLL not found in resources/bridges/: " + dllName);
        }

        String tempDir = System.getProperty("java.io.tmpdir");
        File dllFile;

        try {
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

            Dispatch workbooks = excel.getProperty("Workbooks").toDispatch();
            Dispatch wb = Dispatch.call(workbooks, "Open", workbookPath).toDispatch();

            Session session = new Session(excel);
            session.global = createGlobal;
            session.openWorkbooks.put(workbookPath, wb);
            SessionManager.addSession(sessionId, session);

            return SessionValue
                    .builder()
                    .withSessionObject(new ExcelSession(sessionId, session))
                    .build();

        } catch (Exception e) {
            throw new BotCommandException("Failed to open workbook: " + e.getMessage());
        }
    }
}
