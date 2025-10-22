package com.automationanywhere.botcommand.conditional;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.Locale;

@BotCommand(commandType = BotCommand.CommandType.Condition)
@CommandPkg(
        label = "Worksheet Exists?",
        name = "worksheetExists",
        description = "Retorna TRUE si existe una hoja con el nombre indicado en el workbook de la sesión.",
        node_label = "La hoja '{{worksheetName}}' existe",
        icon = "excel.svg"
)
public class WorksheetExists {

    @ConditionTest
    public Boolean worksheetExists(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty @SessionObject ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Worksheet Name")
            @NotEmpty String worksheetName
    ) {
        if (worksheetName == null || worksheetName.trim().isEmpty()) {
            throw new BotCommandException("Worksheet Name no puede estar vacío.");
        }

        // Obtener sesión y workbook abiertos
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);

        String target = worksheetName.trim().toLowerCase(Locale.ROOT);

        try {
            Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
            int count = Dispatch.get(sheets, "Count").getInt();

            for (int i = 1; i <= count; i++) {
                Dispatch sh = Dispatch.call(sheets, "Item", i).toDispatch();
                String name = safeVariantToString(Dispatch.get(sh, "Name"))
                        .trim().toLowerCase(Locale.ROOT);
                if (name.equals(target)) {
                    return true;   // encontrada
                }
            }
            return false;          // no encontrada
        } catch (Exception e) {
            throw new BotCommandException("No se pudo consultar las hojas del workbook: " + e.getMessage(), e);
        }
    }

    // ------- Helper -------
    private static String safeVariantToString(Variant v) {
        if (v == null || v.isNull()) return "";
        Object o = v.toJavaObject();
        return o != null ? o.toString() : "";
    }
}
