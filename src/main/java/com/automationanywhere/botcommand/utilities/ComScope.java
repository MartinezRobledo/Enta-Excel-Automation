// src/main/java/com/automationanywhere/botcommand/utilities/ComScope.java
package com.automationanywhere.botcommand.utilities;

import com.jacob.com.ComThread;

/** Maneja InitSTA/Release con try-with-resources */
public class ComScope implements AutoCloseable {
    private boolean inited = false;

    public ComScope() {
        ComThread.InitSTA();
        inited = true;
    }

    @Override
    public void close() {
        if (inited) {
            try { ComThread.Release(); } catch (Exception ignore) {}
        }
    }
}