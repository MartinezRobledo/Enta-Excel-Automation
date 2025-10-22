
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
    label = "Delete Pivot Table",
    name = "deletePivotTable",
    description = "Elimina una Tabla Dinámica por su nombre en una hoja específica.",
    icon = "excel.svg"
)
public class DeletePivotTable {

    @Execute
    public void action(
        @Idx(index = "1", type = AttributeType.SESSION)
        @Pkg(label = "Workbook Session")
        @NotEmpty @SessionObject ExcelSession excelSession,

        @Idx(index = "2", type = AttributeType.TEXT)
        @Pkg(label = "Nombre de la hoja", default_value_type = DataType.STRING)
        @NotEmpty String sheetName,

        @Idx(index = "3", type = AttributeType.TEXT)
        @Pkg(label = "Nombre de la Tabla Dinámica", default_value_type = DataType.STRING)
        @NotEmpty String pivotName
    ) {
        // Obtener sesión y workbook
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);

        // Obtener hoja
        Dispatch sheet = requireSheetByName(wb, sheetName);

        // Obtener colección de PivotTables
        Dispatch pivotTables = Dispatch.get(sheet, "PivotTables").toDispatch();

        // Verificar existencia y eliminar
        try {
            Dispatch pivotTable = Dispatch.call(pivotTables, "Item", pivotName).toDispatch();
            Dispatch range = Dispatch.get(pivotTable, "TableRange2").toDispatch(); // Obtener rango antes de borrar
            Dispatch.call(pivotTable, "ClearAllFilters"); // Limpia filtros si hay
            Dispatch.call(pivotTable, "PivotFields"); // Asegura acceso
            Dispatch.call(range, "Clear"); // Borra la tabla dinámica
        } catch (Exception e) {
            throw new BotCommandException("No se pudo eliminar la Tabla Dinámica '" + pivotName + "'. Verificá que exista en la hoja '" + sheetName + "'. Detalle: " + safeMsg(e), e);
        }
    }

    private static Dispatch requireSheetByName(Dispatch wb, String name) {
        try {
            Dispatch sheets = Dispatch.get(wb, "Worksheets").toDispatch();
            return Dispatch.call(sheets, "Item", name).toDispatch();
        } catch (Exception e) {
            throw new BotCommandException("No existe la hoja '" + name + "' en el libro activo.");
        }
    }

    private static String safeMsg(Exception e) {
        String m = (e == null ? "" : e.getMessage());
        return (m == null ? "" : m);
    }
}
