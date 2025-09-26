package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;

@BotCommand
@CommandPkg(
        label = "Paste Image To Excel",
        name = "pasteImageToExcel",
        description = "Pega una imagen en la hoja de Excel en una celda específica, ajustando tamaño",
        icon = "excel.svg"
)
public class PasteImageToExcel {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty
            @SessionObject
            ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            @NotEmpty
            String sheetName,

            @Idx(index = "3", type = AttributeType.FILE)
            @Pkg(label = "Image File Path")
            @NotEmpty
            String imagePath,

            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Start Cell (e.g., B2)", default_value = "A1", default_value_type = DataType.STRING)
            @NotEmpty
            String startCell,

            @Idx(index = "5", type = AttributeType.NUMBER)
            @Pkg(label = "Width in cells", default_value = "3", default_value_type = DataType.NUMBER)
            @NotEmpty
            Double widthCells,

            @Idx(index = "6", type = AttributeType.NUMBER)
            @Pkg(label = "Height in cells", default_value = "4", default_value_type = DataType.NUMBER)
            @NotEmpty
            Double heightCells
    ) {
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);

        Dispatch sheet = Dispatch.call(Dispatch.get(wb, "Sheets").toDispatch(), "Item", sheetName).toDispatch();
        Dispatch range = Dispatch.call(sheet, "Range", startCell).toDispatch();

        // Usar getDouble() para evitar error
        double left = Dispatch.get(range, "Left").getDouble();
        double top = Dispatch.get(range, "Top").getDouble();
        double cellWidth = Dispatch.get(range, "Width").getDouble();
        double cellHeight = Dispatch.get(range, "Height").getDouble();

        double imgWidth = cellWidth * widthCells;
        double imgHeight = cellHeight * heightCells;

        Dispatch shapes = Dispatch.get(sheet, "Shapes").toDispatch();

        // Usar Variant para parámetros booleanos
        Dispatch.call(shapes, "AddPicture",
                new Variant(imagePath),   // Filename
                new Variant(false),       // LinkToFile
                new Variant(true),        // SaveWithDocument
                new Variant(left),
                new Variant(top),
                new Variant(imgWidth),
                new Variant(imgHeight)
        );
    }
}
