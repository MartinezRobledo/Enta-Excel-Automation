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
import com.jacob.com.Variant;

@BotCommand
@CommandPkg(
        label = "Copy Range",
        name = "copyRange",
        description = "Copia un rango de celdas de una hoja a otra dentro del mismo libro de Excel",
        icon = "excel.svg"
)
public class CopyRange {

    // Constante para xlPasteValues
    private static final int xlPasteValues = -4163;

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session")
            @NotEmpty @SessionObject ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Hoja origen", default_value_type = DataType.STRING)
            @NotEmpty String sourceSheetName,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Rango A1 a copiar", default_value_type = DataType.STRING)
            @NotEmpty String sourceRangeA1,

            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Hoja destino", default_value_type = DataType.STRING)
            @NotEmpty String destSheetName,

            @Idx(index = "5", type = AttributeType.TEXT)
            @Pkg(label = "Celda destino (top-left)", default_value_type = DataType.STRING)
            @NotEmpty String destTopLeft,

            @Idx(index = "6", type = AttributeType.CHECKBOX)
            @Pkg(label = "Copiar solo valores (sin formato ni fórmulas)", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean soloValores,

            @Idx(index = "7", type = AttributeType.CHECKBOX)
            @Pkg(label = "Crear hoja destino si no existe", default_value_type = DataType.BOOLEAN, default_value = "false")
            Boolean crearHojaDestinoSiNoExiste
    ) {
        // Obtener sesión y workbook
        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch workbook = ExcelObjects.requireWorkbook(session, excelSession);

        try {
            // Obtener colección de hojas de cálculo (Worksheets: solo hojas, no gráficos)
            Dispatch sheets = Dispatch.get(workbook, "Worksheets").toDispatch();

            // Hoja origen
            Dispatch srcSheet;
            try {
                srcSheet = Dispatch.call(sheets, "Item", sourceSheetName).toDispatch();
            } catch (Exception e) {
                throw new BotCommandException("No existe la hoja origen '" + sourceSheetName + "'.", e);
            }

            // Hoja destino (crear si no existe, según checkbox)
            Dispatch dstSheet;
            try {
                dstSheet = Dispatch.call(sheets, "Item", destSheetName).toDispatch();
            } catch (Exception notFound) {
                if (Boolean.TRUE.equals(crearHojaDestinoSiNoExiste)) {
                    // Crear nueva hoja y nombrarla
                    dstSheet = Dispatch.call(sheets, "Add").toDispatch();
                    try {
                        Dispatch.put(dstSheet, "Name", destSheetName);
                    } catch (Exception nameEx) {
                        throw new BotCommandException("No se pudo crear la hoja destino '" + destSheetName + "'. " +
                                "Verifique que el nombre sea válido y que el libro no tenga la estructura protegida. Detalle: " + safeMsg(nameEx), nameEx);
                    }
                } else {
                    throw new BotCommandException("La hoja destino '" + destSheetName + "' no existe. " +
                            "Active 'Crear hoja destino si no existe' para crearla automáticamente.");
                }
            }

            // Obtener rango origen
            Dispatch srcRange = Dispatch.call(srcSheet, "Range", sourceRangeA1).toDispatch();

            // Copiar rango (Excel hace el trabajo pesado)
            Dispatch.call(srcRange, "Copy");

            // Obtener celda destino (top-left) y pegar
            Dispatch destCell = Dispatch.call(dstSheet, "Range", destTopLeft).toDispatch();
            if (Boolean.TRUE.equals(soloValores)) {
                Dispatch.call(destCell, "PasteSpecial", new Variant(xlPasteValues));
            } else {
                Dispatch.call(destCell, "PasteSpecial");
            }

            // Limpiar el modo de copia (libera portapapeles de Excel)
            try {
                Dispatch app = Dispatch.get(workbook, "Application").toDispatch();
                Dispatch.put(app, "CutCopyMode", false);
            } catch (Exception ignore) {}

        } catch (Exception e) {
            throw new BotCommandException(
                    "No se pudo copiar el rango '" + sourceRangeA1 + "' de '" + sourceSheetName +
                            "' a '" + destSheetName + "' en '" + destTopLeft + "'. Detalle: " + safeMsg(e), e);
        }
    }

    private static String safeMsg(Exception e) {
        String m = (e == null ? "" : e.getMessage());
        return (m == null ? "" : m);
    }
}