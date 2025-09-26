package com.automationanywhere.botcommand.utilities;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.automationanywhere.botcommand.exception.BotCommandException;

import java.io.File;

public class SessionHelper {

    // Normaliza una ruta a una key canónica (absoluta, lowercase). Si es null/empty, devuelve null.
    public static String toWorkbookKey(String maybePath) {
        if (maybePath == null || maybePath.trim().isEmpty()) return null;
        try {
            File f = new File(maybePath);
            String abs = f.getAbsolutePath();
            return abs.toLowerCase();
        } catch (Exception e) {
            return maybePath.toLowerCase();
        }
    }

    // Obtener libro abierto por nombre de key (path) o por nombre visible
    // Compatible: si la clave exacta no está, busca por nombre de hoja de Excel.
    public static Dispatch getWorkbook(Session session, String workbookNameOrPath) {
        if (session == null) {
            throw new BotCommandException("Session is null.");
        }
        // 1) Buscar por key normalizada (path)
        String key = toWorkbookKey(workbookNameOrPath);
        if (key != null && session.openWorkbooks.containsKey(key)) {
            return session.openWorkbooks.get(key);
        }

        // 2) Buscar por nombre visible (Name)
        for (Dispatch wb : session.openWorkbooks.values()) {
            String name = Dispatch.get(wb, "Name").toString();
            if (name.equalsIgnoreCase(workbookNameOrPath)) {
                return wb;
            }
        }
        throw new BotCommandException("Workbook not found in session: " + workbookNameOrPath);
    }

    // Obtener hoja por nombre o índice
    public static Dispatch getSheet(Dispatch wb, String selectBy, String sheetName, Integer sheetIndex) {
        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        int count = Dispatch.get(sheets, "Count").getInt();
        Dispatch sheet;

        if ("index".equalsIgnoreCase(selectBy)) {
            if (sheetIndex < 1 || sheetIndex > count) {
                throw new BotCommandException("Sheet index out of range: " + sheetIndex);
            }
            sheet = Dispatch.call(sheets, "Item", sheetIndex).toDispatch();
        } else {
            sheet = null;
            for (int i = 1; i <= count; i++) {
                Dispatch s = Dispatch.call(sheets, "Item", i).toDispatch();
                String name = Dispatch.get(s, "Name").toString();
                if (name.equalsIgnoreCase(sheetName)) {
                    sheet = s;
                    break;
                }
            }
            if (sheet == null) {
                throw new BotCommandException("Sheet not found: " + sheetName);
            }
        }
        return sheet;
    }

    // Copiar y renombrar hoja, opcionalmente sobreescribiendo
    public static Dispatch copySheet(Dispatch originSheet, Dispatch destWb, String destSheetName, boolean overwriteIfExists) {
        Dispatch destSheets = Dispatch.get(destWb, "Sheets").toDispatch();
        int destSheetCount = Dispatch.get(destSheets, "Count").getInt();

        // Validar si ya existe la hoja destino
        Dispatch existingSheet = null;
        for (int i = 1; i <= destSheetCount; i++) {
            Dispatch s = Dispatch.call(destSheets, "Item", i).toDispatch();
            String name = Dispatch.get(s, "Name").toString();
            if (name.equalsIgnoreCase(destSheetName)) {
                existingSheet = s;
                break;
            }
        }

        if (existingSheet != null) {
            if (overwriteIfExists) {
                Dispatch.call(existingSheet, "Delete");
            } else {
                throw new BotCommandException("Sheet already exists in destination and overwrite is not allowed.");
            }
        }

        // Copiar al final
        int newDestSheetIndex = Dispatch.get(destSheets, "Count").getInt();
        Dispatch lastSheet = Dispatch.call(destSheets, "Item", newDestSheetIndex).toDispatch();
        Dispatch.call(originSheet, "Copy", new Variant(), lastSheet);

        // Renombrar hoja copiada
        Dispatch newSheet = Dispatch.call(destSheets, "Item", newDestSheetIndex + 1).toDispatch();
        Dispatch.put(newSheet, "Name", destSheetName);

        // Guardar cambios
        Dispatch.call(destWb, "Save");

        // Devolvemos la hoja destino
        return newSheet;
    }
}