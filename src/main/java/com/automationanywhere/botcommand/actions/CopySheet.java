package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelSessionManager;
import com.automationanywhere.botcommand.utilities.ExcelSession;
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
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Session ID", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty
            String sessionId,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Source Workbook Path")
            @NotEmpty
            String sourceWorkbookName,

            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select origin sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectSheetBy,

            @Idx(index = "3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Sheet Name")
            @NotEmpty
            String originSheetName,

            @Idx(index = "3.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Origin Sheet Index (1-based)")
            @NumberInteger
            @GreaterThanEqualTo("1")
            @NotEmpty
            Double originSheetIndex,

            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Destination Workbook Path")
            @NotEmpty
            String destWorkbookName,

            @Idx(index = "5", type = AttributeType.TEXT)
            @Pkg(label = "Destination Sheet Name")
            @NotEmpty
            String destSheetName,

            @Idx(index = "6", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "6.1", pkg = @Pkg(label = "Overwrite", value = "overwrite")),
                    @Idx.Option(index = "6.2", pkg = @Pkg(label = "Append", value = "append"))
            })
            @Pkg(label = "Select copy mode", default_value = "overwrite", default_value_type = DataType.STRING)
            @SelectModes
            String selectCopyMode,

            @Idx(index = "7", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "7.1", pkg = @Pkg(label = "Si", value = "Si")),
                    @Idx.Option(index = "7.2", pkg = @Pkg(label = "No", value = "No"))
            })
            @Pkg(label = "Ignore columns?", default_value = "No", default_value_type = DataType.STRING)
            @SelectModes
            String applyIgnoreCol,

            @Idx(index = "7.1.1", type = AttributeType.LIST)
            @Pkg(label = "Columns to ignore (letters; e.g., A,B,C)")
            List<Object> ignoreColumns,

            //Nueva funcionalidad a impactar
            @Idx(index = "8", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "8.1", pkg = @Pkg(label = "Si", value = "Si")),
                    @Idx.Option(index = "8.2", pkg = @Pkg(label = "No", value = "No"))
            })
            @Pkg(label = "Ignore headers?", default_value = "Si", default_value_type = DataType.STRING)
            @SelectModes
            String preserveHeaders,

            @Idx(index = "8.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Headers range (e.g., A3:H3)")
            @NotEmpty
            String rangeHeaders
    ) {
        ExcelSession session = ExcelSessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found: " + sessionId);

        Dispatch sourceWb = session.openWorkbooks.get(sourceWorkbookName);
        if (sourceWb == null)
            throw new BotCommandException("Source workbook not open: " + sourceWorkbookName);

        Dispatch destWb = session.openWorkbooks.get(destWorkbookName);
        if (destWb == null)
            throw new BotCommandException("Destination workbook not open: " + destWorkbookName);

        Dispatch sourceSheets = Dispatch.get(sourceWb, "Sheets").toDispatch();
        Dispatch originSheet = "index".equalsIgnoreCase(selectSheetBy)
                ? Dispatch.call(sourceSheets, "Item", originSheetIndex.intValue()).toDispatch()
                : Dispatch.call(sourceSheets, "Item", originSheetName).toDispatch();

        Dispatch destSheets = Dispatch.get(destWb, "Sheets").toDispatch();
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

            if ("overwrite".equalsIgnoreCase(selectCopyMode)) {
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
            } else if ("append".equalsIgnoreCase(selectCopyMode)) {
                if (destSheet == null) {
                    throw new BotCommandException("Destination sheet does not exist. Cannot append.");
                }
                Dispatch usedRange = Dispatch.get(destSheet, "UsedRange").toDispatch();
                int lastRow = Dispatch.get(Dispatch.get(usedRange, "Rows").toDispatch(), "Count").getInt();
                lastRow = lastRow > 0 ? lastRow + 1 : 1;
                destStart = Dispatch.call(destSheet, "Cells", lastRow, 1).toDispatch();
            } else {
                throw new BotCommandException("Invalid copy mode: " + selectCopyMode);
            }

            // Copiar contenido
            Dispatch.call(sourceUsedRange, "Copy", destStart);

            // Eliminar columnas ignoradas si corresponde
            if ("Si".equalsIgnoreCase(applyIgnoreCol) && ignoreColumns != null && !ignoreColumns.isEmpty()) {
                List<String> colsToDelete = new ArrayList<>();
                for (Object o : ignoreColumns) {
                    if (o == null) continue;
                    String s = o.toString();
                    if (!s.trim().isEmpty()) {
                        for (String part : s.split(",")) {
                            if (!part.trim().isEmpty()) colsToDelete.add(part.trim().toUpperCase());
                        }
                    }
                }

                Collections.sort(colsToDelete, (a, b) -> Integer.compare(excelColumnLetterToNumber(b), excelColumnLetterToNumber(a)));
                for (String colLetter : colsToDelete) {
                    try {
                        Dispatch colRange = Dispatch.call(destSheet, "Columns", colLetter).toDispatch();
                        Dispatch.call(colRange, "Delete");
                    } catch (Exception ex) {
                        // Ignorar columnas inválidas
                    }
                }
            }

            // Nueva funcionalidad: ignorar encabezados
            if ("Si".equalsIgnoreCase(preserveHeaders) && rangeHeaders != null && !rangeHeaders.trim().isEmpty()) {
                try {
                    // Parsear la letra de columna y la fila de inicio del rango original
                    String[] parts = rangeHeaders.split(":");
                    if (parts.length != 2) {
                        throw new BotCommandException("Rango de encabezados inválido: " + rangeHeaders);
                    }

                    String startCell = parts[0].toUpperCase();
                    String endCell = parts[1].toUpperCase();

                    int startCol = excelColumnLetterToNumber(startCell.replaceAll("\\d", ""));
                    int startRow = Integer.parseInt(startCell.replaceAll("\\D", ""));
                    int endCol = excelColumnLetterToNumber(endCell.replaceAll("\\d", ""));
                    int endRow = Integer.parseInt(endCell.replaceAll("\\D", ""));

                    // Obtener la fila de destino donde se pegó la primera celda
                    int destStartRow = Dispatch.get(destStart, "Row").getInt();

                    // Ajustar el rango de headers a la fila de destino
                    String destRange = numberToColumnLetter(startCol) + destStartRow + ":" + numberToColumnLetter(endCol) + (destStartRow + (endRow - startRow));

                    Dispatch headerRange = Dispatch.call(destSheet, "Range", destRange).toDispatch();
                    Dispatch.call(headerRange, "Delete", -4162); // xlShiftUp
                } catch (Exception ex) {
                    throw new BotCommandException("Error al borrar el rango de encabezados: " + ex.getMessage());
                }
            }


        } catch (Exception e) {
            throw new BotCommandException("Error copying sheet content: " + e.getMessage());
        }
    }

    private static int excelColumnLetterToNumber(String col) {
        int res = 0;
        for (int i = 0; i < col.length(); i++) {
            res = res * 26 + (col.charAt(i) - 'A' + 1);
        }
        return res;
    }
}
