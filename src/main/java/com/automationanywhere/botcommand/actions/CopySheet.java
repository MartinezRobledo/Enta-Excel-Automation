package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Copy Sheet",
        name = "copySheet",
        description = "Duplica una hoja como Excel (Worksheet.Copy), en el mismo u otro workbook",
        icon = "excel.svg"
)
public class CopySheet {

    @Execute
    public void action(
            // --- ORIGEN ---
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Source Workbook Session")
            @NotEmpty @SessionObject ExcelSession sourceExcelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name",  value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select origin sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String selectOriginSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Sheet Name") String originSheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Origin Sheet Index (1-based)")
            @NumberInteger @GreaterThanEqualTo("1") Double originSheetIndex,

            // --- DESTINO ---
            @Idx(index = "3", type = AttributeType.SESSION)
            @Pkg(label = "Destination Workbook Session")
            @NotEmpty @SessionObject ExcelSession destExcelSession,

            @Idx(index = "4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "At end (after last sheet)", value = "end")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "Before sheet", value = "before")),
                    @Idx.Option(index = "4.3", pkg = @Pkg(label = "After sheet",  value = "after"))
            })
            @Pkg(label = "Insert position", default_value = "end", default_value_type = DataType.STRING)
            @SelectModes String positionMode,

            // Ancla opcional si BEFORE / AFTER
            @Idx(index = "4.2.1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.2.1.1", pkg = @Pkg(label = "By Name",  value = "name")),
                    @Idx.Option(index = "4.2.1.2", pkg = @Pkg(label = "By Index", value = "index"))
            })
            @Pkg(label = "Select anchor (BEFORE) by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String beforeBy,

            @Idx(index = "4.2.1.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Anchor Sheet Name (BEFORE)") String beforeName,

            @Idx(index = "4.2.1.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Anchor Sheet Index (1-based, BEFORE)")
            @NumberInteger @GreaterThanEqualTo("1") Double beforeIndex,

            @Idx(index = "4.3.1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.3.1.1", pkg = @Pkg(label = "By Name",  value = "name")),
                    @Idx.Option(index = "4.3.1.2", pkg = @Pkg(label = "By Index", value = "index"))
            })
            @Pkg(label = "Select anchor (AFTER) by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes String afterBy,

            @Idx(index = "4.3.1.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Anchor Sheet Name (AFTER)") String afterName,

            @Idx(index = "4.3.1.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Anchor Sheet Index (1-based, AFTER)")
            @NumberInteger @GreaterThanEqualTo("1") Double afterIndex,

            // Nombre final opcional (si se omite, toma el de origen)
            @Idx(index = "5", type = AttributeType.TEXT)
            @Pkg(label = "Final Sheet Name (optional; default = origin name)")
            String finalSheetName,

            // NUEVO: Overwrite si el nombre ya existe
            @Idx(index = "6", type = AttributeType.CHECKBOX)
            @Pkg(label = "Overwrite if name exists", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean overwriteIfExists
    ) {

        // 1) Sesiones y workbooks
        Session srcSession = ExcelObjects.requireSession(sourceExcelSession);
        Dispatch wbSrc     = ExcelObjects.requireWorkbook(srcSession, sourceExcelSession);

        Session dstSession = ExcelObjects.requireSession(destExcelSession);
        Dispatch wbDst     = ExcelObjects.requireWorkbook(dstSession, destExcelSession);

        // 2) Origen
        Dispatch srcSheets = Dispatch.get(wbSrc, "Sheets").toDispatch();
        Dispatch srcSheet  = "index".equalsIgnoreCase(selectOriginSheetBy)
                ? Dispatch.call(srcSheets, "Item", originSheetIndex.intValue()).toDispatch()
                : Dispatch.call(srcSheets, "Item", originSheetName).toDispatch();

        String originName  = Dispatch.get(srcSheet, "Name").getString();
        String desiredName = (finalSheetName != null && !finalSheetName.trim().isEmpty())
                ? finalSheetName.trim()
                : originName;

        // Validación básica de nombre
        if (desiredName.length() > 31 || desiredName.matches(".*[:\\\\/?*\\[\\]].*")) {
            throw new BotCommandException("Invalid sheet name (<=31 chars, no : \\ / ? * [ ]).");
        }

        // 3) Destino: buscar existencia de desiredName
        Dispatch dstSheets = Dispatch.get(wbDst, "Sheets").toDispatch();
        int dstCount = Dispatch.get(dstSheets, "Count").getInt();
        Dispatch existingSameName = findByName(dstSheets, desiredName, dstCount);

        // 4) Definir Before/After
        Variant before = Variant.VT_MISSING;
        Variant after  = Variant.VT_MISSING;

        if (existingSameName != null && Boolean.TRUE.equals(overwriteIfExists)) {
            // OVERWRITE: copiamos BEFORE de la hoja existente, para preservar posición
            before = new Variant(existingSameName);
        } else {
            // Posicionamiento normal
            if ("end".equalsIgnoreCase(positionMode)) {
                Dispatch last = Dispatch.call(dstSheets, "Item", dstCount).toDispatch();
                after = new Variant(last);
            } else if ("before".equalsIgnoreCase(positionMode)) {
                Dispatch anchor = resolveAnchor(dstSheets, beforeBy, beforeName, beforeIndex, dstCount);
                if (anchor == null) throw new BotCommandException("Anchor (BEFORE) not found.");
                before = new Variant(anchor);
            } else if ("after".equalsIgnoreCase(positionMode)) {
                Dispatch anchor = resolveAnchor(dstSheets, afterBy, afterName, afterIndex, dstCount);
                if (anchor == null) throw new BotCommandException("Anchor (AFTER) not found.");
                after = new Variant(anchor);
            } else {
                throw new BotCommandException("Invalid position mode: " + positionMode);
            }
        }

        // 5) Optimización + alerts
        Dispatch app = Dispatch.get(wbSrc, "Application").toDispatch();
        boolean prevUpd = getBool(app, "ScreenUpdating");
        boolean prevEvt = getBool(app, "EnableEvents");
        boolean prevAlr = getBool(app, "DisplayAlerts");

        try {
            putBool(app, "ScreenUpdating", false);
            putBool(app, "EnableEvents",  false);
            // DisplayAlerts lo manejamos manualmente al borrar

            // 6) Copiar hoja
            Dispatch.call(srcSheet, "Copy", before, after);

            // La nueva hoja queda ActiveSheet en destino
            Dispatch newSheet = Dispatch.get(wbDst, "ActiveSheet").toDispatch();

            if (existingSameName != null && Boolean.TRUE.equals(overwriteIfExists)) {
                // Desactivar alerts para borrar sin prompt
                putBool(app, "DisplayAlerts", false);
                try {
                    // Borrar la hoja que queremos “pisar”
                    Dispatch.call(existingSameName, "Delete");
                } finally {
                    // Restaurar alerts
                    putBool(app, "DisplayAlerts", prevAlr);
                }
            } else if (existingSameName != null) {
                // No overwrite -> fallar como pedís
                throw new BotCommandException("Destination already has a sheet named '" + desiredName + "'.");
            }

            // 7) Renombrar a desiredName (ya no existe colisión si hubo overwrite)
            Dispatch.put(newSheet, "Name", desiredName);

        } catch (Exception e) {
            throw new BotCommandException("DuplicateSheet failed: " + e.getMessage(), e);
        } finally {
            putBool(app, "DisplayAlerts", prevAlr);
            putBool(app, "EnableEvents",  prevEvt);
            putBool(app, "ScreenUpdating",prevUpd);
        }
    }

    // --- Helpers ---

    private static Dispatch resolveAnchor(Dispatch dstSheets, String by, String name, Double index, int count) {
        if ("index".equalsIgnoreCase(by)) {
            if (index == null) return null;
            int i = index.intValue();
            if (i < 1 || i > count) return null;
            return Dispatch.call(dstSheets, "Item", i).toDispatch();
        } else {
            if (name == null || name.trim().isEmpty()) return null;
            String wanted = name.trim();
            for (int i = 1; i <= count; i++) {
                Dispatch s  = Dispatch.call(dstSheets, "Item", i).toDispatch();
                String nm   = Dispatch.get(s, "Name").getString();
                if (nm != null && nm.equalsIgnoreCase(wanted)) return s;
            }
            return null;
        }
    }

    private static Dispatch findByName(Dispatch dstSheets, String name, int count) {
        for (int i = 1; i <= count; i++) {
            Dispatch s = Dispatch.call(dstSheets, "Item", i).toDispatch();
            String nm  = Dispatch.get(s, "Name").getString();
            if (nm != null && nm.equalsIgnoreCase(name)) return s;
        }
        return null;
    }

    private static boolean getBool(Dispatch app, String prop) {
        try { return Dispatch.get(app, prop).getBoolean(); } catch (Exception e) { return true; }
    }
    private static void putBool(Dispatch app, String prop, boolean v) {
        try { Dispatch.put(app, prop, v); } catch (Exception ignore) {}
    }
}
