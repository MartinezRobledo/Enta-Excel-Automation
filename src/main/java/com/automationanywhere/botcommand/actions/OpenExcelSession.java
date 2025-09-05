package com.automationanywhere.botcommand.actions;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;

import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.ExcelSessionManager;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

@BotCommand
@CommandPkg(
        name = "OpenExcelSession",
        label = "Open Excel Session",
        node_label = "Open Excel COM Session",
        description = "Opens an Excel session using COM automation and JACOB",
        icon = "excel.svg"
)
public class OpenExcelSession {

    @Execute
    public void openSession(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Session Name", description = "Name of the Excel session", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty
            String sessionName
    ) throws Exception {

        // Detect architecture
        boolean is64Bit = System.getProperty("os.arch").contains("64");
        String dllName = is64Bit ? "BridgeCOM64.dll" : "BridgeCOM32.dll";

        // Check if DLL exists in resources/bridges/
        InputStream dllStream = this.getClass().getClassLoader().getResourceAsStream("bridges/" + dllName);
        if (dllStream == null) {
            throw new Exception("DLL not found in resources/bridges/: " + dllName);
        }

        // Copy DLL to temp folder
        String tempDir = System.getProperty("java.io.tmpdir");
        File dllFile = new File(tempDir, dllName);
        if (!dllFile.exists()) {
            try (FileOutputStream out = new FileOutputStream(dllFile)) {
                byte[] buffer = new byte[1024];
                int read;
                while ((read = dllStream.read(buffer)) != -1) {
                    out.write(buffer, 0, read);
                }
            }
        }

        // Load DLL dynamically
        System.setProperty("jacob.dll.path", dllFile.getAbsolutePath());
        System.load(dllFile.getAbsolutePath());

        // Initialize COM and start Excel
        ComThread.InitSTA();
        ActiveXComponent excel = new ActiveXComponent("Excel.Application");
        excel.setProperty("Visible", false);

        // Guardar la sesión en algún helper para futuras acciones si se desea
        // Crear un ExcelSession con el ActiveXComponent
        ExcelSession session = new ExcelSession(excel);

        // Guardar en el manager
        ExcelSessionManager.addSession(sessionName, session);

    }
}
