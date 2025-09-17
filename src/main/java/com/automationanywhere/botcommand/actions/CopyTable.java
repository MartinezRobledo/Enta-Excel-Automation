package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.SessionManager;
import com.automationanywhere.botcommand.utilities.Session;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.*;
import java.util.stream.Collectors;

import static com.automationanywhere.botcommand.utilities.ExcelHelpers.splitRangeByIgnoredColumns;

@BotCommand
@CommandPkg(
        label = "Copy Table Content",
        name = "copyTableContent",
        description = "Copies a table (with headers range) from one sheet to another with options for headers, filtering and columns ignoring",
        icon = "excel.svg",
        return_type = DataType.BOOLEAN,
        return_required = false
)
public class CopyTable {

    @Execute
    public Value<Boolean> action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Session ID", default_value_type = DataType.STRING, default_value = "Default")
            @NotEmpty String sessionId,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Source Workbook Path") @NotEmpty String sourceWorkbookName,

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
            @NotEmpty
            Double originSheetIndex,

            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Headers Range (e.g., C9:BM9)") @NotEmpty String headersRange,

            @Idx(index = "5", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "All rows", value = "all")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "Only visible rows", value = "visible"))
            })
            @Pkg(label = "Rows to copy", default_value = "all", default_value_type = DataType.STRING)
            @SelectModes
            String rowsMode,

            @Idx(index = "6", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "6.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "6.2", pkg = @Pkg(label = "No", value = "no"))
            })
            @Pkg(label = "Include headers?", default_value = "no", default_value_type = DataType.STRING)
            @SelectModes
            String includeHeaders,

            @Idx(index = "7", type = AttributeType.TEXT)
            @Pkg(label = "Destination Workbook Path") @NotEmpty String destWorkbookName,

            @Idx(index = "8", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "8.1", pkg = @Pkg(label = "Name", value = "name")),
                    @Idx.Option(index = "8.2", pkg = @Pkg(label = "Index", value = "index"))
            })
            @Pkg(label = "Select origin sheet by", default_value = "name", default_value_type = DataType.STRING)
            @SelectModes
            String selectDestSheetBy,

            @Idx(index = "8.1.1", type = AttributeType.TEXT)
            @Pkg(label = "Origin Sheet Name")
            @NotEmpty
            String destSheetName,

            @Idx(index = "8.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "Origin Sheet Index (1-based)")
            @NotEmpty
            Double destSheetIndex,

            @Idx(index = "9", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "9.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "9.2", pkg = @Pkg(label = "No", value = "no"))
            })
            @Pkg(label = "Create sheet if not exists?", default_value = "yes", default_value_type = DataType.STRING)
            @SelectModes
            String createSheet,

            @Idx(index = "10", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "10.1", pkg = @Pkg(label = "Overwrite", value = "overwrite")),
                    @Idx.Option(index = "10.2", pkg = @Pkg(label = "Append", value = "append")),
                    @Idx.Option(index = "10.3", pkg = @Pkg(label = "Manual", value = "manual"))
            })
            @Pkg(label = "Copy mode", default_value = "overwrite", default_value_type = DataType.STRING)
            @SelectModes
            String copyMode,

            @Idx(index = "10.3.1", type = AttributeType.TEXT)
            @Pkg(label = "Manual destination cell (e.g., A2)") String manualCell,

            @Idx(index = "11", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "11.1", pkg = @Pkg(label = "Yes", value = "yes")),
                    @Idx.Option(index = "11.2", pkg = @Pkg(label = "No", value = "no"))
            })
            @Pkg(label = "Ignore columns?", default_value = "no", default_value_type = DataType.STRING)
            @SelectModes
            String applyIgnoreCol,

            @Idx(index = "11.1.1", type = AttributeType.LIST)
            @Pkg(label = "Columns to ignore (letters; e.g., A,B,C)")
            List<Object> ignoreColumns


    ) {

        Session session = SessionManager.getSession(sessionId);
        if (session == null || session.excelApp == null)
            throw new BotCommandException("Session not found: " + sessionId);

        Dispatch sourceWb = session.openWorkbooks.get(sourceWorkbookName);
        Dispatch destWb = session.openWorkbooks.get(destWorkbookName);
        if (sourceWb == null || destWb == null)
            throw new BotCommandException("Workbook not open (source or destination).");

        Dispatch sourceSheets = Dispatch.get(sourceWb, "Sheets").toDispatch();
        Dispatch sourceSheet;

        if ("index".equalsIgnoreCase(selectSheetBy)) {
            sourceSheet = Dispatch.call(sourceSheets, "Item", originSheetIndex.intValue()).toDispatch();
        } else {
            sourceSheet = Dispatch.call(sourceSheets, "Item", originSheetName).toDispatch();
        }

        Dispatch destSheets = Dispatch.get(destWb, "Sheets").toDispatch();
        Dispatch destSheet = null;

        // buscar hoja por nombre o índice
        for (int i = 1; i <= Dispatch.get(destSheets, "Count").getInt(); i++) {
            Dispatch s = Dispatch.call(destSheets, "Item", i).toDispatch();
            String sheetName = Dispatch.get(s, "Name").toString();
            if ("index".equalsIgnoreCase(selectDestSheetBy) && i == destSheetIndex.intValue()) {
                destSheet = s; break;
            } else if ("name".equalsIgnoreCase(selectDestSheetBy) && sheetName.equalsIgnoreCase(destSheetName)) {
                destSheet = s; break;
            }
        }

        if (destSheet == null) {
            if ("yes".equalsIgnoreCase(createSheet)) {
                destSheet = Dispatch.call(destSheets, "Add").toDispatch();
                Dispatch.put(destSheet, "Name", destSheetName);
            } else {
                throw new BotCommandException("Destination sheet does not exist.");
            }
        }

        try {
            // Determinar rango completo desde headers
            String startCell = headersRange.split(":")[0];
            String endCell = headersRange.split(":")[1];
            int headerRow = Integer.parseInt(endCell.replaceAll("\\D",""));

            // Encontrar última fila con datos
            Dispatch usedRange = Dispatch.get(sourceSheet, "UsedRange").toDispatch();
            Dispatch rows = Dispatch.get(usedRange, "Rows").toDispatch();
            int lastRow = Dispatch.get(rows, "Count").getInt();

            // Ajustar si se copian headers o no
            int startRow = "yes".equalsIgnoreCase(includeHeaders) ? headerRow : headerRow + 1;

            String copyRange = startCell.replaceAll("\\d","") + startRow + ":" + endCell.replaceAll("\\d","") + lastRow;

            Dispatch rangeToCopy = Dispatch.call(sourceSheet, "Range", copyRange).toDispatch();

            // Si solo visibles
            boolean isFilteredCopy = false;
            if ("visible".equalsIgnoreCase(rowsMode)) {
                rangeToCopy = Dispatch.call(rangeToCopy, "SpecialCells", new Variant(12)).toDispatch(); // xlCellTypeVisible=12
                isFilteredCopy = true;
            }

            // Si no hay datos en el rango
            if (rangeToCopy == null) {
                return new BooleanValue(false);
            }

            // Determinar celda de inicio en destino
            Dispatch destStart;
            if ("overwrite".equalsIgnoreCase(copyMode)) {
                Dispatch usedRangeDest = Dispatch.get(destSheet, "UsedRange").toDispatch();
                Dispatch.call(usedRangeDest, "Clear");
                destStart = Dispatch.call(destSheet, "Range", "A1").toDispatch();
            } else if ("append".equalsIgnoreCase(copyMode)) {
                Dispatch destUsedRange = Dispatch.get(destSheet, "UsedRange").toDispatch();
                int lastDestRow = Dispatch.get(Dispatch.get(destUsedRange, "Rows").toDispatch(), "Count").getInt();
                destStart = Dispatch.call(destSheet, "Cells", lastDestRow+1, 1).toDispatch();
            } else {
                destStart = Dispatch.call(destSheet, "Range", manualCell).toDispatch();
            }

            // Copiar ignorando columnas (compactado)
            if ("yes".equalsIgnoreCase(applyIgnoreCol) && ignoreColumns != null && !ignoreColumns.isEmpty()) {
                List<String> ranges = splitRangeByIgnoredColumns(
                        copyRange,
                        ignoreColumns.stream().map(Object::toString).collect(Collectors.toList())
                );

                if (ranges.isEmpty()) {
                    return new BooleanValue(false);
                }

                int destColOffset = 0;
                for (String subRange : ranges) {
                    Dispatch subRangeObj = Dispatch.call(sourceSheet, "Range", subRange).toDispatch();
                    int blockCols = Dispatch.get(Dispatch.get(subRangeObj, "Columns").toDispatch(), "Count").getInt();
                    int blockRows = Dispatch.get(Dispatch.get(subRangeObj, "Rows").toDispatch(), "Count").getInt();

                    int destRow = Dispatch.get(destStart, "Row").getInt();
                    int destCol = Dispatch.get(destStart, "Column").getInt() + destColOffset;
                    Dispatch adjustedDest = Dispatch.call(destSheet, "Cells", destRow, destCol).toDispatch();
                    Dispatch adjustedDestResized = Dispatch.call(adjustedDest, "Resize", blockRows, blockCols).toDispatch();

                    // Descombinar celdas
                    Dispatch.put(subRangeObj, "MergeCells", false);

                    if (isFilteredCopy) {
                        try {
                            Dispatch visibleCells = Dispatch.call(subRangeObj, "SpecialCells", new Variant(12)).toDispatch(); // xlCellTypeVisible
                            Dispatch.call(visibleCells, "Copy", adjustedDestResized);
                        } catch (Exception e) {
                            // No hay filas visibles en este subrango
                        }
                    } else {
                        Dispatch.call(subRangeObj, "Copy", adjustedDestResized);
                    }

                    destColOffset += blockCols;
                }
            } else {
                // Copiar todo sin ignorar columnas
                Dispatch rangeObj = rangeToCopy;
                Dispatch.put(rangeObj, "MergeCells", false);

                if (isFilteredCopy) {
                    try {
                        Dispatch visibleCells = Dispatch.call(rangeObj, "SpecialCells", new Variant(12)).toDispatch();
                        Dispatch.call(visibleCells, "Copy", destStart);
                    } catch (Exception e) {
                        return new BooleanValue(false); // No hay filas visibles
                    }
                } else {
                    Dispatch.call(rangeObj, "Copy", destStart);
                }
            }

            //Se copiaron valores
            return new BooleanValue(true);

        } catch (Exception e) {
            throw new BotCommandException("Error in CopyTable: " + e.getMessage(), e);
        }
    }
}
