package com.automationanywhere.botcommand.utilities;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.automationanywhere.botcommand.exception.BotCommandException;

import java.util.Map;

public class ExcelSessionHelper {

    // Obtener libro abierto por nombre
    public static Dispatch getWorkbook(ExcelSession session, String workbookName) {
        Dispatch wb = session.openWorkbooks.get(workbookName);
        if (wb == null) {
            throw new BotCommandException("Workbook not found in session: " + workbookName);
        }
        return wb;
    }

    // Obtener hoja por nombre o índice
    public static Dispatch getSheet(Dispatch wb, String selectBy, String sheetName, Integer sheetIndex) {
        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        int count = Dispatch.get(sheets, "Count").getInt();
        Dispatch sheet = null;

        if ("index".equalsIgnoreCase(selectBy)) {
            if (sheetIndex < 1 || sheetIndex > count) {
                throw new BotCommandException("Sheet index out of range: " + sheetIndex);
            }
            sheet = Dispatch.call(sheets, "Item", sheetIndex).toDispatch();
        } else {
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

    // Copiar hoja a otro libro, renombrando si es necesario y opcionalmente sobreescribiendo
    // Antes: public static void copySheet(...)
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

        // DEVOLVEMOS la hoja destino
        return newSheet;
    }
}
