package com.automationanywhere.botcommand.utilities;

import com.jacob.com.LibraryLoader;

import java.io.File;

public class JacobLoader {
    private static boolean loaded = false;

    public static synchronized void loadJacob(File dllFile) {
        if (!loaded) {
            System.setProperty("jacob.dll.path", dllFile.getAbsolutePath());
            LibraryLoader.loadJacobLibrary();
            loaded = true;
        }
    }

    public static boolean isLoaded() {
        return loaded;
    }
}
