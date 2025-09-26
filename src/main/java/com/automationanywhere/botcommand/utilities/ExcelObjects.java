// src/main/java/com/automationanywhere/botcommand/utilities/ExcelObjects.java
package com.automationanywhere.botcommand.utilities;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.jacob.com.Dispatch;

import java.io.File;

public class ExcelObjects {

    public static Session requireSession(ExcelSession excelSession) {
        if (excelSession == null) throw new BotCommandException("Workbook Session is null.");
        Session s = excelSession.getSession();
        if (s == null || s.excelApp == null) throw new BotCommandException("Session not found or closed.");
        return s;
    }

    /**
     * IMPORTANTE: NO usa el Dispatch guardado en session.openWorkbooks porque puede venir de otro hilo.
     * Re-resuelve el Workbook en ESTE hilo, buscando por FullName contra workbookKey.
     */
    public static Dispatch requireWorkbook(Session session, ExcelSession excelSession) {
        String key = excelSession.getWorkbookKey();
        if (key == null || key.trim().isEmpty()) {
            throw new BotCommandException("Invalid workbook reference in ExcelSession.");
        }
        String normKey = new File(key).getAbsolutePath(); // normaliza
        try {
            Dispatch workbooks = session.excelApp.getProperty("Workbooks").toDispatch();
            int count = com.jacob.com.Dispatch.get(workbooks, "Count").getInt();
            for (int i = 1; i <= count; i++) {
                Dispatch wb = com.jacob.com.Dispatch.call(workbooks, "Item", i).toDispatch();
                String full = com.jacob.com.Dispatch.get(wb, "FullName").toString();
                String normFull = new File(full).getAbsolutePath();
                if (normFull.equalsIgnoreCase(normKey)) {
                    return wb; // Workbook resuelto EN ESTE HILO
                }
            }
        } catch (Exception e) {
            throw new BotCommandException("Failed to resolve workbook on current thread: " + e.getMessage(), e);
        }
        throw new BotCommandException("Workbook not open in session: " + normKey);
    }

    public static Dispatch requireSheet(Dispatch wb, String selectBy, String sheetName, Double sheetIndexNullable) {
        Integer idx = (sheetIndexNullable == null) ? null : sheetIndexNullable.intValue();
        return SessionHelper.getSheet(wb, selectBy, sheetName, idx);
    }
}