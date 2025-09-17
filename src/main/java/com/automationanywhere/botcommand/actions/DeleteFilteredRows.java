package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Delete Filtered Rows",
        name = "deleteRows",
        description = "Elimina todas las filas visibles",
        icon = "excel.svg"
)
public class DeleteFilteredRows {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Session ID", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty
            String sessionId,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Workbook Path")
            @NotEmpty
            String workbookName,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty
            String sheetName,

            @Idx(index = "4", type = AttributeType.NUMBER)
            @Pkg(label = "Header Row Number (e.g., 1)")
            @NotEmpty
            Double headerRowNumber
    ) {

        if (headerRowNumber == null || headerRowNumber < 1)
            throw new BotCommandException("Header row number must be a positive integer.");

        Session session = SessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Excel session not found: " + sessionId);

        Dispatch workbook = session.openWorkbooks.get(workbookName);
        if (workbook == null)
            throw new BotCommandException("Workbook not open: " + workbookName);

        Dispatch sheet = Dispatch.call(Dispatch.get(workbook, "Sheets").toDispatch(), "Item", sheetName).toDispatch();

        try {
            Dispatch usedRange = Dispatch.get(sheet, "UsedRange").toDispatch();

            // Determinar última fila
            Dispatch rows = Dispatch.get(usedRange, "Rows").toDispatch();
            int lastUsedRow = Dispatch.get(rows, "Count").getInt() + Dispatch.get(usedRange, "Row").getInt() - 1;

            int firstRowToDelete = headerRowNumber.intValue() + 1;
            if (firstRowToDelete > lastUsedRow) return; // No hay filas para borrar

            // Crear rango a partir de la fila siguiente a los headers
            Dispatch rangeToDelete = Dispatch.call(sheet, "Range",
                    "A" + firstRowToDelete + ":A" + lastUsedRow).toDispatch();

            // Obtener solo las filas visibles del filtro
            Dispatch visibleCells;
            try {
                visibleCells = Dispatch.call(rangeToDelete, "SpecialCells", new Variant(12)).toDispatch(); // xlCellTypeVisible
            } catch (Exception e) {
                return; // No hay filas visibles, nada que borrar
            }

            // Borrar filas completas
            Dispatch visibleRows = Dispatch.get(visibleCells, "EntireRow").toDispatch();
            Dispatch.call(visibleRows, "Delete");

        } catch (Exception e) {
            throw new BotCommandException("Error deleting filtered rows: " + e.getMessage());
        }
    }
}
