package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;

@BotCommand
@CommandPkg(
        label = "Copy Range Content",
        name = "copyRangeContent",
        description = "Copies a specific range from one sheet to another workbook sheet",
        icon = "excel.svg"
)
public class CopyRange {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Source Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession sourceExcelSession,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Source Sheet Name")
            @NotEmpty
            String sourceSheetName,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Range to copy (e.g., C10:BM36)")
            @NotEmpty
            String sourceRange,

            @Idx(index = "4", type = AttributeType.SESSION)
            @Pkg(label = "Source Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession destExcelSession,

            @Idx(index = "5", type = AttributeType.TEXT)
            @Pkg(label = "Destination Sheet Name")
            @NotEmpty
            String destSheetName,

            @Idx(index = "6", type = AttributeType.TEXT)
            @Pkg(label = "Destination start cell (e.g., A1)")
            @NotEmpty
            String destStartCell
    ) {
        Session sourceSession = sourceExcelSession.getSession();
        if (sourceSession == null || sourceSession.excelApp == null)
            throw new BotCommandException("Source Session not found o closed");

        if (sourceSession.openWorkbooks.isEmpty())
            throw new BotCommandException("Source workbook: No workbook is open in this session.");

        Dispatch wb1 = sourceSession.openWorkbooks.values().iterator().next();

        Session destSession = destExcelSession.getSession();
        if (destSession == null || destSession.excelApp == null)
            throw new BotCommandException("Destination Session not found o closed");

        if (sourceSession.openWorkbooks.isEmpty())
            throw new BotCommandException("Destination workbook: No workbook is open in this session.");

        Dispatch wb2 = destSession.openWorkbooks.values().iterator().next();

        Dispatch sourceSheets = Dispatch.get(wb1, "Sheets").toDispatch();
        Dispatch sourceSheet = Dispatch.call(sourceSheets, "Item", sourceSheetName).toDispatch();

        Dispatch destSheets = Dispatch.get(wb2, "Sheets").toDispatch();
        Dispatch destSheet = null;
        int destSheetCount = Dispatch.get(destSheets, "Count").getInt();

        // Buscar o crear hoja destino
        for (int i = 1; i <= destSheetCount; i++) {
            Dispatch s = Dispatch.call(destSheets, "Item", i).toDispatch();
            if (Dispatch.get(s, "Name").toString().equalsIgnoreCase(destSheetName)) {
                destSheet = s;
                break;
            }
        }
        if (destSheet == null) {
            destSheet = Dispatch.call(destSheets, "Add").toDispatch();
            Dispatch.put(destSheet, "Name", destSheetName);
        }

        try {
            // Obtener rango a copiar
            Dispatch sourceRangeDispatch = Dispatch.call(sourceSheet, "Range", sourceRange).toDispatch();
            Dispatch destStartDispatch = Dispatch.call(destSheet, "Range", destStartCell).toDispatch();

            // Copiar rango al destino
            Dispatch.call(sourceRangeDispatch, "Copy", destStartDispatch);

        } catch (Exception e) {
            throw new BotCommandException("Error copying range content: " + e.getMessage(), e);
        }
    }
}
