package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.JacobLoader;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.UUID;

@BotCommand
@CommandPkg(
        name = "OpenSession",
        label = "Open Excel Session",
        node_label = "Open Excel COM Session",
        description = "Opens an Excel session using COM automation and JACOB",
        icon = "excel.svg"
)
public class OpenSession {

    @Execute
    public void openSession(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Session Name", description = "Name of the Excel session", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty
            String sessionName,

            @Idx(index = "2", type = AttributeType.CHECKBOX)
            @Pkg(label = "Use Global Session", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean global
    ) throws Exception {

        boolean is64Bit = System.getProperty("os.arch").contains("64");
        String dllName = is64Bit ? "BridgeCOM64.dll" : "BridgeCOM32.dll";

        InputStream dllStream = this.getClass().getClassLoader().getResourceAsStream("bridges/" + dllName);
        if (dllStream == null) {
            throw new Exception("DLL not found in resources/bridges/: " + dllName);
        }

        String tempDir = System.getProperty("java.io.tmpdir");
        File dllFile;

        if (Boolean.TRUE.equals(global)) {
            // DLL única global
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
            // DLL local con nombre único
            String uniqueName = dllName.replace(".dll", "_" + UUID.randomUUID() + ".dll");
            dllFile = new File(tempDir, uniqueName);
            Files.copy(dllStream, dllFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
            System.setProperty("jacob.dll.path", dllFile.getAbsolutePath());
            com.jacob.com.LibraryLoader.loadJacobLibrary();
        }

        // Inicializar COM y Excel
        ComThread.InitSTA();
        ActiveXComponent excel = new ActiveXComponent("Excel.Application");
        excel.setProperty("Visible", false);

        Session session = new Session(excel);
        session.global = global;
        SessionManager.addSession(sessionName, session);
    }
}
