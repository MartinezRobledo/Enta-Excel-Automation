package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
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
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select origin sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Sheet Name")
            @NotEmpty
            String originSheetName,

            @Idx(index = "2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Origin Sheet Index (1-based)")
            @NotEmpty
            Double originSheetIndex,

            @Idx(index = "3", type = AttributeType.NUMBER)
            @Pkg(label = "Header Row Number (e.g., 1)")
            @NotEmpty
            Double headerRowNumber
    ) {

        if (headerRowNumber == null || headerRowNumber < 1)
            throw new BotCommandException("Header row number must be a positive integer.");

        Session session = excelSession.getSession();
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found o closed");

        if (session.openWorkbooks.isEmpty())
            throw new BotCommandException("No workbook is open in this session.");

        Dispatch wb = session.openWorkbooks.values().iterator().next();

        Dispatch sheets = Dispatch.get(wb, "Sheets").toDispatch();
        Dispatch sheet = "index".equalsIgnoreCase(selectSheetBy)
                ? Dispatch.call(sheets, "Item", originSheetIndex.intValue()).toDispatch()
                : Dispatch.call(sheets, "Item", originSheetName).toDispatch();

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
