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

@BotCommand(commandType = BotCommand.CommandType.Condition) // <- Condicional
@CommandPkg(
        label = "Worksheet Not Exists?",
        name = "worksheetNotExistsCond",
        description = "TRUE si NO existe una hoja con el nombre indicado",
        node_label = "Sheet '{{worksheetName}}' not exists",
        icon = "excel.svg"
)
public class WorksheetNotExists {

    @ConditionTest // <- Método evaluador (debe devolver Boolean)
    public Boolean test(
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

        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);

        String target = worksheetName.trim().toLowerCase(Locale.ROOT);

        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        int count = Dispatch.get(sheets, "Count").getInt();

        for (int i = 1; i <= count; i++) {
            Dispatch sh = Dispatch.call(sheets, "Item", i).toDispatch();
            String name = safeVariantToString(Dispatch.get(sh, "Name"))
                    .trim().toLowerCase(Locale.ROOT);
            if (name.equals(target)) return false; // existe => la inversa es FALSE
        }
        return true; // no se encontró => TRUE
    }

    private static String safeVariantToString(Variant v) {
        if (v == null || v.isNull()) return "";
        Object o = v.toJavaObject();
        return o != null ? o.toString() : "";
    }
}