package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import static com.automationanywhere.botcommand.utilities.ExcelHelpers.numberToColumnLetter;

@BotCommand
@CommandPkg(
        label = "Copy Sheet Content",
        name = "copySheetContent",
        description = "Copies the content of a sheet to another sheet without saving",
        icon = "excel.svg"
)
public class CopySheet {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Source Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession sourceExcelSession,

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
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double originSheetIndex,

            @Idx(index = "3", type = AttributeType.SESSION)
            @Pkg(label = "Destination Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession destExcelSession,

            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Destination Sheet Name")
            @NotEmpty
            String destSheetName,

            @Idx(index = "5", type = AttributeType.CHECKBOX)
            @Pkg(label = "Overwrite destination sheet", default_value = "true", default_value_type = DataType.BOOLEAN)
            @SelectModes
            Boolean overwrite
    ) {
        // 1) Sesión + workbook correctos
        Session sourceSession = ExcelObjects.requireSession(sourceExcelSession);
        Dispatch wb1 = ExcelObjects.requireWorkbook(sourceSession, sourceExcelSession);

        Session destSession = ExcelObjects.requireSession(destExcelSession);
        Dispatch wb2 = ExcelObjects.requireWorkbook(destSession, destExcelSession);

        Dispatch sourceSheets = Dispatch.get(wb1, "Sheets").toDispatch();
        Dispatch originSheet = "index".equalsIgnoreCase(selectSheetBy)
                ? Dispatch.call(sourceSheets, "Item", originSheetIndex.intValue()).toDispatch()
                : Dispatch.call(sourceSheets, "Item", originSheetName).toDispatch();

        Dispatch destSheets = Dispatch.get(wb2, "Sheets").toDispatch();
        Dispatch destSheet = null;
        int destSheetCount = Dispatch.get(destSheets, "Count").getInt();

        for (int i = 1; i <= destSheetCount; i++) {
            Dispatch s = Dispatch.call(destSheets, "Item", i).toDispatch();
            if (Dispatch.get(s, "Name").toString().equalsIgnoreCase(destSheetName)) {
                destSheet = s;
                break;
            }
        }

        try {
            Dispatch sourceUsedRange = Dispatch.get(originSheet, "UsedRange").toDispatch();
            Dispatch destStart;

            if (overwrite) {
                if (destSheet == null) {
                    // Crear hoja si no existe
                    destSheet = Dispatch.call(destSheets, "Add").toDispatch();
                    Dispatch.put(destSheet, "Name", destSheetName);
                } else {
                    // Limpiar hoja existente
                    Dispatch usedRange = Dispatch.get(destSheet, "UsedRange").toDispatch();
                    Dispatch.call(usedRange, "Clear");
                }
                destStart = Dispatch.call(destSheet, "Range", "A1").toDispatch();
            }

            // Copiar contenido
            Dispatch.call(sourceUsedRange, "Copy", "A1");

        } catch (Exception e) {
            throw new BotCommandException("Error copying sheet content: " + e.getMessage());
        }
    }
}