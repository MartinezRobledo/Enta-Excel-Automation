
package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;

@BotCommand
@CommandPkg(
    label = "Delete Worksheet",
    name = "deleteWorksheet",
    description = "Elimina una hoja específica de un libro de Excel",
    icon = "excel.svg"
)
public class DeleteWorksheet {

    @Execute
    public void action(
        @Idx(index = "1", type = AttributeType.SESSION)
        @Pkg(label = "Workbook Session")
        @NotEmpty @SessionObject ExcelSession excelSession,

        @Idx(index = "2", type = AttributeType.TEXT)
        @Pkg(label = "Nombre de la hoja a eliminar", default_value_type = DataType.STRING)
        @NotEmpty String sheetName
    ) {
        // Obtener sesión y workbook
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch workbook = ExcelObjects.requireWorkbook(session, excelSession);

        try {
            Dispatch sheets = Dispatch.get(workbook, "Worksheets").toDispatch();
            Dispatch sheetToDelete = Dispatch.call(sheets, "Item", sheetName).toDispatch();
            Dispatch.call(sheetToDelete, "Delete");
        } catch (Exception e) {
            throw new BotCommandException("No se pudo eliminar la hoja '" + sheetName + "'. Verificá que exista en el libro activo. Detalle: " + safeMsg(e), e);
        }
    }

    private static String safeMsg(Exception e) {
        String m = (e == null ? "" : e.getMessage());
        return (m == null ? "" : m);
    }
}
