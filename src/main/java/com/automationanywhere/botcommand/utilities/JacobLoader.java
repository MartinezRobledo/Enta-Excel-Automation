package com.automationanywhere.botcommand.utilities;

import com.jacob.com.LibraryLoader;

import java.io.File;

public class JacobLoader {
    private static volatile boolean loaded = false;

    public static synchronized boolean isLoaded() {
        return loaded;
    }

    public static synchronized void loadJacob(File dllFile) {
        if (loaded) return;
        System.setProperty("jacob.dll.path", dllFile.getAbsolutePath());
        LibraryLoader.loadJacobLibrary();
        loaded = true;
    }
}