package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
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
        label = "Rename Worksheet",
        name = "renameWorksheet",
        description = "Rename a worksheet by name or index. If the target name already exists, the action fails.",
        icon = "excel.svg"
)
public class RenameWorksheet {

    @Execute
    public void action(
            // ==== Sesión / Workbook ====
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession excelSession,

            // ==== Selección de hoja origen ====
            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name (origin)")
            @NotEmpty
            String sheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Sheet Index (origin)", description = "1-based index")
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double sheetIndex,

            // ==== Modo de renombrado ====
            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "New name", value = "new")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Replace in current name", value = "replace"))
            })
            @Pkg(label = "Rename mode", default_value = "new", default_value_type = DataType.STRING)
            @SelectModes
            String renameMode,

            // --- Modo: Nuevo nombre ---
            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "New Sheet Name")
            String newSheetName,

            // --- Modo: Reemplazo parcial ---
            @Idx(index = "3.2.1", type = AttributeType.TEXT)
            @Pkg(label = "Search for (in current name) - empty allowed")
            String searchFor,

            @Idx(index = "3.2.2", type = AttributeType.TEXT)
            @Pkg(label = "Replace with (can be empty)")
            String replaceWith
    ) {
        // 1) Sesión + workbook
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch workbook = ExcelObjects.requireWorkbook(session, excelSession);

        // 2) Resolver hoja origen
        Dispatch sheet = ExcelObjects.requireSheet(workbook, selectSheetBy, sheetName, sheetIndex);
        try {
            Dispatch.call(sheet, "Activate");
        } catch (Exception ignore) {
        }

        // 3) Nombre actual
        String currentName = Dispatch.get(sheet, "Name").getString();

        // 4) Determinar nombre destino según modo
        String desired;
        if ("replace".equalsIgnoreCase(renameMode)) {
            String search = (searchFor == null) ? "" : searchFor;
            String repl = (replaceWith == null) ? "" : replaceWith;

            // Reemplazo literal, sensible a mayúsculas/minúsculas, en todas las apariciones
            if (search.isEmpty()) {
                // Si search está vacío, no tiene sentido reemplazar; no cambiar nada
            }
            desired = currentName.replace(search, repl);

            // Si no hay cambios, devolver sin renombrar (idempotente)
            if (desired.equals(currentName)) {
            }
        } else {
            // Modo por defecto: "new"
            desired = (newSheetName == null) ? "" : newSheetName.trim();
            if (desired.isEmpty()) {
                throw new BotCommandException("New Sheet Name cannot be empty.");
            }
        }

        // 5) Validaciones de nombre Excel
        validateExcelSheetName(desired);

        // 6) Política de duplicados: FALLAR si ya existe otro con ese nombre (case-insensitive)
        // Permitir cambio solo de mayúsculas/minúsculas sobre la misma hoja
        if (!desired.equalsIgnoreCase(currentName) && sheetExistsByName(workbook, desired)) {
            throw new BotCommandException("A worksheet named '" + desired + "' already exists in the workbook.");
        }

        // 7) Renombrar
        try {
            Dispatch.put(sheet, "Name", desired);
        } catch (Exception e) {
            throw new BotCommandException("Failed to rename worksheet to '" + desired + "': " + e.getMessage(), e);
        }

    }

    // ===== Helpers =====

    private static void validateExcelSheetName(String name) {
        if (name == null || name.trim().isEmpty()) {
            throw new BotCommandException("Sheet name cannot be empty.");
        }
        if (name.length() > 31) {
            throw new BotCommandException("Sheet name '" + name + "' exceeds 31 characters (Excel limit).");
        }
        if (containsInvalidChars(name)) {
            throw new BotCommandException(
                    "Invalid sheet name. The name cannot contain any of the following characters: : \\ / ? * [ ]"
            );
        }
    }

    private static boolean containsInvalidChars(String name) {
        return name.contains(":") || name.contains("\\") || name.contains("/")
                || name.contains("?") || name.contains("*")
                || name.contains("[") || name.contains("]");
    }

    private static boolean sheetExistsByName(Dispatch workbook, String name) {
        Dispatch sheets = Dispatch.get(workbook, "Worksheets").toDispatch();
        try {
            // Address by name (case-insensitive in Excel COM). If found, exists.
            Dispatch found = Dispatch.call(sheets, "Item", new Variant(name)).toDispatch();
            // Asegurar que realmente coincide (por si Excel resolvió algo raro)
            String resolved = Dispatch.get(found, "Name").getString();
            return resolved.equalsIgnoreCase(name);
        } catch (Exception notFound) {
            return false;
        }
    }
}