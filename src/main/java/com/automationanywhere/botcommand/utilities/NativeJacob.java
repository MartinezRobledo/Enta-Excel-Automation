package com.automationanywhere.botcommand.utilities;

import com.automationanywhere.botcommand.exception.BotCommandException;

import java.io.InputStream;
import java.io.IOException;
import java.nio.file.*;
import java.util.UUID;

public final class NativeJacob {
    private static volatile boolean LOADED = false;

    private NativeJacob() {}

    public static synchronized void load(boolean global) {
        if (LOADED) return;

        boolean is64 = System.getProperty("os.arch").contains("64");
        String dllName = is64 ? "jacob-1.21-x64.dll" : "jacob-1.21-x32.dll";
        String resourcePath = "bridges/" + dllName;

        try (InputStream in = NativeJacob.class.getClassLoader().getResourceAsStream(resourcePath)) {
            if (in == null) {
                throw new BotCommandException("DLL not found in classpath: " + resourcePath);
            }

            Path tempDir = Paths.get(System.getProperty("java.io.tmpdir"));
            Path target;
            if (global) {
                // Nombre estable para reusar entre invocaciones
                target = tempDir.resolve(dllName);
                if (!Files.exists(target)) {
                    Files.copy(in, target, StandardCopyOption.REPLACE_EXISTING);
                }
            } else {
                // Nombre único para evitar archivos bloqueados
                String unique = dllName.replace(".dll", "_" + UUID.randomUUID() + ".dll");
                target = tempDir.resolve(unique);
                Files.copy(in, target, StandardCopyOption.REPLACE_EXISTING);
                target.toFile().deleteOnExit();
            }

            // Carga explícita por ruta absoluta (evita java.library.path)
            System.load(target.toAbsolutePath().toString());
            LOADED = true;

        } catch (IOException e) {
            throw new BotCommandException("Failed to load Jacob DLL: " + e.getMessage(), e);
        }
    }
}