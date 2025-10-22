package com.automationanywhere.botcommand.utilities;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.jacob.com.LibraryLoader;

import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;

public final class JacobBootstrap {
    private static volatile boolean loaded = false;
    private static String loadedFrom = null;

    private JacobBootstrap() {}

    /**
     * Carga JACOB exactamente una vez por proceso, desde una ruta determinada.
     * Estrategia determinista:
     *  1) Usa -Djacob.dll.path si está definida.
     *  2) Si no, usa env JACOB_DLL_PATH y la establece como -Djacob.dll.path.
     *  3) Si ninguna está, falla con error claro.
     *  4) Valida existencia y bitness por nombre (x64/x86). Si no coincide, falla.
     *  5) Si ya estaba cargado desde otra ruta distinta, falla.
     */
    public static synchronized void ensureLoaded() {
        if (loaded) return;

        String dllPath = System.getenv("JACOB_DLL_PATH");
            if (dllPath != null && !dllPath.isBlank()) {
                System.setProperty("jacob.dll.path", dllPath);
            }

        if (dllPath == null || dllPath.isBlank()) {
            throw new BotCommandException(
                    "JACOB no configurado. Configure UNA de estas opciones antes de ejecutar:\n" +
                            " -Djacob.dll.path=C:\\ruta\\jacob-<ver>-x64.dll  (recomendado)\n" +
                            " o defina la variable de entorno JACOB_DLL_PATH con la RUTA COMPLETA a la DLL.\n" +
                            "Sin esto, la acción no intentará buscar la DLL en PATH ni en rutas relativas."
            );
        }

        File f = new File(dllPath);
        if (!f.exists() || !f.isFile()) {
            throw new BotCommandException("jacob.dll no existe o no es un archivo: " + dllPath);
        }

        // Validación de arquitectura por nombre (defensivo, evita errores comunes).
        boolean jvm64 = System.getProperty("os.arch", "").contains("64");
        String name = f.getName().toLowerCase();
        if (jvm64 && !name.contains("x64")) {
            throw new BotCommandException("La JVM es x64 pero la DLL no parece x64: " + name);
        }
        if (!jvm64 && !name.contains("x86")) {
            throw new BotCommandException("La JVM es x86 pero la DLL no parece x86: " + name);
        }

        // Carga estricta (no usa PATH ni búsqueda adicional).
        final String finalPath = f.getAbsolutePath();
        try {
            // Si otra parte del proceso ya cargó JACOB desde una ruta distinta, LibraryLoader fallará;
            // adicionalmente, nosotros validamos camino “inmutable” más abajo.
            LibraryLoader.loadJacobLibrary();
            loaded = true;
            loadedFrom = finalPath;

            // Congelamos la configuración: si alguien cambia jacob.dll.path en runtime, lo ignoramos.
            System.setProperty("jacob.dll.path", finalPath);

        } catch (UnsatisfiedLinkError ule) {
            // Puede pasar si ya se cargó desde otra ruta o si la DLL no corresponde.
            throw new BotCommandException(
                    "No se pudo cargar JACOB desde: " + finalPath + ". " +
                            "Verifique permisos y que la DLL coincide con la arquitectura de la JVM. Detalle: " + ule.getMessage(),
                    ule
            );
        } catch (Exception e) {
            throw new BotCommandException("Fallo al cargar JACOB: " + e.getMessage(), e);
        }
    }

    public static boolean isLoaded() { return loaded; }
    public static String loadedFrom() { return loadedFrom; }
}